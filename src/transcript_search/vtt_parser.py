"""
Parser for VTT (WebVTT) and plain-text transcript formats.

Supported formats
-----------------
VTT  — standard WebVTT files produced by Microsoft Teams.
       Each cue block looks like::

           00:00:01.000 --> 00:00:04.000
           <v John Smith>Hello everyone

TEXT — simple line-based text files where each line is a segment.
       Optionally prefixed with a timestamp in [HH:MM:SS] notation and/or
       a "Speaker: " prefix.
"""

from __future__ import annotations

import re
from typing import List

from .models import TranscriptSegment

# ---------------------------------------------------------------------------
# VTT helpers
# ---------------------------------------------------------------------------

# Matches a VTT timestamp range:  00:00:01.000 --> 00:00:04.000
_VTT_TIMESTAMP_RE = re.compile(
    r"^(\d{2}:\d{2}:\d{2}[.,]\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}[.,]\d{3})"
)

# Strips cue settings and HTML tags from cue payloads
_HTML_TAG_RE = re.compile(r"<[^>]+>")

# Extracts <v SpeakerName> voice span
_VOICE_SPAN_RE = re.compile(r"<v\s+([^>]+)>")


def _strip_html(text: str) -> str:
    return _HTML_TAG_RE.sub("", text).strip()


def _extract_speaker(text: str) -> tuple[str | None, str]:
    """Return (speaker, cleaned_text) from a VTT cue payload."""
    match = _VOICE_SPAN_RE.search(text)
    if match:
        speaker = match.group(1).strip()
        cleaned = _strip_html(text)
        return speaker, cleaned
    return None, _strip_html(text)


def parse_vtt(content: str) -> List[TranscriptSegment]:
    """
    Parse a WebVTT transcript into a list of :class:`TranscriptSegment`.

    Parameters
    ----------
    content:
        Raw VTT file content as a string.

    Returns
    -------
    List[TranscriptSegment]
        Segments in order of appearance; segments with empty text are skipped.
    """
    segments: List[TranscriptSegment] = []
    lines = content.splitlines()
    index = 0
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # Skip WEBVTT header and NOTE blocks
        if line.startswith("WEBVTT") or line.startswith("NOTE"):
            i += 1
            continue

        ts_match = _VTT_TIMESTAMP_RE.match(line)
        if ts_match:
            start_time = ts_match.group(1)
            end_time = ts_match.group(2)

            # Collect cue payload lines (until blank line or EOF)
            i += 1
            payload_lines: List[str] = []
            while i < len(lines) and lines[i].strip():
                payload_lines.append(lines[i])
                i += 1

            raw_payload = " ".join(payload_lines)
            if raw_payload.strip():
                speaker, text = _extract_speaker(raw_payload)
                if text:
                    segments.append(
                        TranscriptSegment(
                            index=index,
                            start_time=start_time,
                            end_time=end_time,
                            speaker=speaker,
                            text=text,
                        )
                    )
                    index += 1
        else:
            i += 1

    return segments


# ---------------------------------------------------------------------------
# Plain-text helpers
# ---------------------------------------------------------------------------

# Optional leading timestamp in several common forms:
#   [00:01:23]  |  [00:01:23.456]  |  (00:01:23)
_TEXT_TIMESTAMP_RE = re.compile(
    r"^[\[\(](\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?)\s*[\]\)]\s*"
)

# Optional "Speaker Name: " prefix (only if followed by non-space text)
_TEXT_SPEAKER_RE = re.compile(r"^([A-Za-z][^:]{0,50}):\s+(?=\S)")


def parse_text(content: str) -> List[TranscriptSegment]:
    """
    Parse a plain-text transcript into a list of :class:`TranscriptSegment`.

    Each non-empty line becomes a segment.  Lines may optionally start with:
    * a timestamp like ``[00:01:23]`` or ``(00:01:23)``
    * a speaker prefix like ``John Smith: ``

    Parameters
    ----------
    content:
        Raw text content as a string.

    Returns
    -------
    List[TranscriptSegment]
        Segments in order of appearance; blank lines are skipped.
    """
    segments: List[TranscriptSegment] = []
    index = 0

    for raw_line in content.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        start_time: str | None = None
        speaker: str | None = None

        ts_match = _TEXT_TIMESTAMP_RE.match(line)
        if ts_match:
            start_time = ts_match.group(1)
            line = line[ts_match.end():]

        sp_match = _TEXT_SPEAKER_RE.match(line)
        if sp_match:
            speaker = sp_match.group(1).strip()
            line = line[sp_match.end():]

        text = line.strip()
        if text:
            segments.append(
                TranscriptSegment(
                    index=index,
                    start_time=start_time,
                    end_time=None,
                    speaker=speaker,
                    text=text,
                )
            )
            index += 1

    return segments


_DEFAULT_VTT_CONTENT_TYPE = "application/vnd.ms-teams.transcript.vtt"


def parse_transcript(content: str, content_type: str = _DEFAULT_VTT_CONTENT_TYPE) -> List[TranscriptSegment]:
    """
    Dispatch to the appropriate parser based on *content_type*.

    Parameters
    ----------
    content:
        Raw transcript content.
    content_type:
        MIME type returned by the Graph API.  VTT types are routed to
        :func:`parse_vtt`; everything else falls back to :func:`parse_text`.

    Returns
    -------
    List[TranscriptSegment]
    """
    if "vtt" in content_type.lower():
        return parse_vtt(content)
    return parse_text(content)
