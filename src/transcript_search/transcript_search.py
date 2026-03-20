"""
Meeting transcript search module.

This module provides :class:`TranscriptSearchClient`, which uses the Microsoft
Graph API to:

* List all Teams online-meeting transcripts for a user within a date range.
* Download and parse transcript content (VTT or plain text).
* Perform full-text keyword searches across transcripts, returning matched
  segments together with ±*context_lines* of surrounding context.
* Handle meetings with no transcript by flagging them as
  :attr:`TranscriptStatus.NOT_RECORDED`.
* Surface permission errors as :attr:`TranscriptStatus.ACCESS_DENIED` rather
  than raising uncaught exceptions.

Graph API endpoints used
------------------------
``/users/{userId}/onlineMeetings``
    Enumerate meetings in the requested date window.
``/users/{userId}/onlineMeetings/{meetingId}/transcripts``
    List transcripts for a specific meeting.
``/users/{userId}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content``
    Download raw transcript content (VTT or text).
"""

from __future__ import annotations

import logging
import re
from datetime import datetime, timezone
from typing import Dict, Iterable, List, Optional, Sequence

from .graph_client import GraphClient, GraphAPIError, GraphNotFoundError, GraphPermissionError
from .models import (
    MeetingTranscriptInfo,
    TranscriptMatch,
    TranscriptSearchResult,
    TranscriptSegment,
    TranscriptStatus,
)
from .vtt_parser import parse_transcript

logger = logging.getLogger(__name__)

# Default number of context lines to include around a match
_DEFAULT_CONTEXT_LINES = 5


def _to_iso(dt: datetime) -> str:
    """Return an ISO-8601 string with UTC timezone suitable for OData filters."""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def _search_segments(
    segments: List[TranscriptSegment],
    keywords: Sequence[str],
    context_lines: int = _DEFAULT_CONTEXT_LINES,
) -> List[TranscriptMatch]:
    """
    Search *segments* for any of *keywords* (case-insensitive).

    Each hit produces a :class:`TranscriptMatch` containing the matched segment
    together with up to *context_lines* segments before and after it.

    Parameters
    ----------
    segments:
        Ordered list of transcript segments.
    keywords:
        One or more keywords to look for.  A segment is a hit if it contains
        **any** of the supplied keywords.
    context_lines:
        Number of segments to include before and after each match.

    Returns
    -------
    List[TranscriptMatch]
        Matches in order of appearance in the transcript.
    """
    if not segments or not keywords:
        return []

    patterns = [re.compile(re.escape(kw), re.IGNORECASE) for kw in keywords]
    matches: List[TranscriptMatch] = []

    for idx, segment in enumerate(segments):
        for pattern in patterns:
            if pattern.search(segment.text):
                before = segments[max(0, idx - context_lines): idx]
                after = segments[idx + 1: idx + 1 + context_lines]
                matches.append(
                    TranscriptMatch(
                        keyword=pattern.pattern,
                        matched_segment=segment,
                        context_before=before,
                        context_after=after,
                    )
                )
                break  # One match per segment is enough

    return matches


class TranscriptSearchClient:
    """
    High-level client for listing and searching Teams meeting transcripts.

    Parameters
    ----------
    graph_client:
        An authenticated :class:`~transcript_search.graph_client.GraphClient`
        instance.
    context_lines:
        Default number of context segments to include around keyword hits.
    """

    def __init__(
        self,
        graph_client: GraphClient,
        context_lines: int = _DEFAULT_CONTEXT_LINES,
    ) -> None:
        self._client = graph_client
        self._context_lines = context_lines

    # ------------------------------------------------------------------
    # Listing meetings
    # ------------------------------------------------------------------

    def list_meetings(
        self,
        user_id: str,
        start_date: datetime,
        end_date: datetime,
    ) -> List[Dict]:
        """
        Return raw meeting objects from Graph for *user_id* between the given dates.

        Parameters
        ----------
        user_id:
            Azure AD object ID or UPN of the user.
        start_date:
            Inclusive start of the date window (UTC).
        end_date:
            Inclusive end of the date window (UTC).

        Returns
        -------
        List[Dict]
            Raw Graph ``onlineMeeting`` objects.
        """
        start_str = _to_iso(start_date)
        end_str = _to_iso(end_date)
        filter_q = (
            f"startDateTime ge {start_str} and startDateTime le {end_str}"
        )
        path = f"/users/{user_id}/onlineMeetings?$filter={filter_q}"
        return list(self._client.get_paginated(path))

    # ------------------------------------------------------------------
    # Transcript availability
    # ------------------------------------------------------------------

    def list_transcripts_for_meeting(
        self,
        user_id: str,
        meeting_id: str,
    ) -> List[Dict]:
        """
        Return raw transcript objects for a single *meeting_id*.

        Returns an empty list when the meeting has no transcripts.
        Raises :class:`GraphPermissionError` on 403, transparently re-raises
        other errors.
        """
        path = f"/users/{user_id}/onlineMeetings/{meeting_id}/transcripts"
        try:
            return list(self._client.get_paginated(path))
        except GraphNotFoundError:
            logger.debug("No transcripts endpoint for meeting %s (404).", meeting_id)
            return []

    def get_meeting_transcript_info(
        self,
        user_id: str,
        meeting: Dict,
    ) -> MeetingTranscriptInfo:
        """
        Build a :class:`MeetingTranscriptInfo` for the given raw *meeting* dict.

        Catches permission errors and marks status as
        :attr:`TranscriptStatus.ACCESS_DENIED` rather than propagating them.
        """
        meeting_id = meeting.get("id", "")
        subject = meeting.get("subject", "") or "(no subject)"
        join_url = meeting.get("joinWebUrl") or meeting.get("joinUrl")

        start_raw = meeting.get("startDateTime")
        end_raw = meeting.get("endDateTime")
        start_time = _parse_dt(start_raw) if start_raw else None
        end_time = _parse_dt(end_raw) if end_raw else None

        organizer: Optional[str] = None
        participants = meeting.get("participants", {})
        org_data = participants.get("organizer", {}).get("identity", {})
        user_data = org_data.get("user", {})
        organizer = user_data.get("displayName") or user_data.get("id")

        try:
            transcripts = self.list_transcripts_for_meeting(user_id, meeting_id)
        except GraphPermissionError as exc:
            logger.warning(
                "Permission denied accessing transcripts for meeting %s: %s",
                meeting_id,
                exc,
            )
            return MeetingTranscriptInfo(
                meeting_id=meeting_id,
                subject=subject,
                start_time=start_time,
                end_time=end_time,
                organizer=organizer,
                join_url=join_url,
                transcript_status=TranscriptStatus.ACCESS_DENIED,
                error_message=str(exc),
            )
        except GraphAPIError as exc:
            logger.error("Error fetching transcripts for meeting %s: %s", meeting_id, exc)
            return MeetingTranscriptInfo(
                meeting_id=meeting_id,
                subject=subject,
                start_time=start_time,
                end_time=end_time,
                organizer=organizer,
                join_url=join_url,
                transcript_status=TranscriptStatus.ERROR,
                error_message=str(exc),
            )

        transcript_ids = [t.get("id", "") for t in transcripts if t.get("id")]
        status = (
            TranscriptStatus.AVAILABLE if transcript_ids else TranscriptStatus.NOT_RECORDED
        )

        return MeetingTranscriptInfo(
            meeting_id=meeting_id,
            subject=subject,
            start_time=start_time,
            end_time=end_time,
            organizer=organizer,
            join_url=join_url,
            transcript_status=status,
            transcript_ids=transcript_ids,
        )

    # ------------------------------------------------------------------
    # Listing transcripts across a date range
    # ------------------------------------------------------------------

    def list_available_transcripts(
        self,
        user_id: str,
        start_date: datetime,
        end_date: datetime,
    ) -> List[MeetingTranscriptInfo]:
        """
        Return transcript availability info for all meetings in the date window.

        Parameters
        ----------
        user_id:
            Azure AD object ID or UPN.
        start_date:
            Start of date window (UTC).
        end_date:
            End of date window (UTC).

        Returns
        -------
        List[MeetingTranscriptInfo]
            One entry per meeting.  Meetings without transcripts have
            ``transcript_status == TranscriptStatus.NOT_RECORDED``.
        """
        meetings = self.list_meetings(user_id, start_date, end_date)
        results: List[MeetingTranscriptInfo] = []
        for meeting in meetings:
            info = self.get_meeting_transcript_info(user_id, meeting)
            results.append(info)
        return results

    # ------------------------------------------------------------------
    # Downloading transcript content
    # ------------------------------------------------------------------

    def download_transcript(
        self,
        user_id: str,
        meeting_id: str,
        transcript_id: str,
        content_type: str = "application/vnd.ms-teams.transcript.vtt",
    ) -> List[TranscriptSegment]:
        """
        Download and parse a single transcript.

        Parameters
        ----------
        user_id:
            Azure AD object ID or UPN.
        meeting_id:
            Graph online meeting ID.
        transcript_id:
            Graph transcript ID.
        content_type:
            MIME type for the transcript format.  Defaults to VTT.

        Returns
        -------
        List[TranscriptSegment]
            Parsed transcript segments.

        Raises
        ------
        GraphPermissionError
            When transcript access is denied (403).
        GraphAPIError
            For other API errors.
        """
        accept_header = content_type
        path = (
            f"/users/{user_id}/onlineMeetings/{meeting_id}"
            f"/transcripts/{transcript_id}/content"
        )
        raw = self._client.get_raw(path, accept=accept_header)
        return parse_transcript(raw, content_type)

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def search_transcripts(
        self,
        user_id: str,
        keywords: Sequence[str],
        start_date: datetime,
        end_date: datetime,
        context_lines: Optional[int] = None,
    ) -> List[TranscriptSearchResult]:
        """
        Search all transcripts in the date window for the given keywords.

        For each meeting the method:

        1. Checks whether a transcript is available.
        2. Downloads and parses each transcript.
        3. Searches for any of the supplied *keywords* (case-insensitive).
        4. Returns matched segments with ±*context_lines* surrounding segments.

        Meetings with no transcript are included in the results with
        ``transcript_status == TranscriptStatus.NOT_RECORDED`` and an empty
        ``matches`` list so callers can distinguish "no results" from
        "no transcript".

        Parameters
        ----------
        user_id:
            Azure AD object ID or UPN.
        keywords:
            One or more search terms.  A segment matches if it contains
            **any** keyword.
        start_date:
            Start of date window (UTC).
        end_date:
            End of date window (UTC).
        context_lines:
            Override the instance-level ``context_lines`` setting.

        Returns
        -------
        List[TranscriptSearchResult]
            One entry per meeting, ordered by meeting start time.
        """
        ctx = context_lines if context_lines is not None else self._context_lines
        transcript_infos = self.list_available_transcripts(user_id, start_date, end_date)
        results: List[TranscriptSearchResult] = []

        for info in transcript_infos:
            if info.transcript_status != TranscriptStatus.AVAILABLE:
                results.append(TranscriptSearchResult(meeting=info))
                continue

            all_matches: List[TranscriptMatch] = []
            for transcript_id in info.transcript_ids:
                try:
                    segments = self.download_transcript(
                        user_id, info.meeting_id, transcript_id
                    )
                except GraphPermissionError as exc:
                    logger.warning(
                        "Permission denied downloading transcript %s for meeting %s: %s",
                        transcript_id,
                        info.meeting_id,
                        exc,
                    )
                    info.transcript_status = TranscriptStatus.ACCESS_DENIED
                    info.error_message = str(exc)
                    break
                except GraphAPIError as exc:
                    logger.error(
                        "Error downloading transcript %s for meeting %s: %s",
                        transcript_id,
                        info.meeting_id,
                        exc,
                    )
                    info.transcript_status = TranscriptStatus.ERROR
                    info.error_message = str(exc)
                    break
                else:
                    all_matches.extend(
                        _search_segments(segments, keywords, context_lines=ctx)
                    )

            results.append(TranscriptSearchResult(meeting=info, matches=all_matches))

        return results

    def search_meeting_transcript(
        self,
        user_id: str,
        meeting_id: str,
        transcript_id: str,
        keywords: Sequence[str],
        context_lines: Optional[int] = None,
    ) -> List[TranscriptMatch]:
        """
        Search a specific, already-known transcript for *keywords*.

        Parameters
        ----------
        user_id:
            Azure AD object ID or UPN.
        meeting_id:
            Graph online meeting ID.
        transcript_id:
            Graph transcript ID.
        keywords:
            Search terms.
        context_lines:
            Number of context segments; defaults to the instance setting.

        Returns
        -------
        List[TranscriptMatch]
        """
        ctx = context_lines if context_lines is not None else self._context_lines
        segments = self.download_transcript(user_id, meeting_id, transcript_id)
        return _search_segments(segments, keywords, context_lines=ctx)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_dt(value: str) -> Optional[datetime]:
    """Parse an ISO-8601 datetime string, returning ``None`` on failure."""
    try:
        from dateutil.parser import parse as du_parse  # type: ignore
        return du_parse(value)
    except Exception:
        try:
            return datetime.fromisoformat(value.rstrip("Z"))
        except Exception:
            logger.debug("Could not parse datetime: %r", value)
            return None
