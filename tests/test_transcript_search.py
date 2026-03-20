"""
Unit tests for the meeting transcript search module.

Tests cover:
- VTT parsing (well-formed, edge cases)
- Plain-text transcript parsing
- Keyword search with context lines
- TranscriptSearchClient: listing, searching, and error handling
"""

from __future__ import annotations

import sys
import os
from datetime import datetime, timezone
from typing import Dict, List
from unittest.mock import MagicMock, patch, call

import pytest

# Ensure the src package is importable when running from the repo root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from transcript_search.graph_client import (
    GraphAPIError,
    GraphNotFoundError,
    GraphPermissionError,
)
from transcript_search.models import (
    MeetingTranscriptInfo,
    TranscriptMatch,
    TranscriptSearchResult,
    TranscriptSegment,
    TranscriptStatus,
)
from transcript_search.transcript_search import (
    TranscriptSearchClient,
    _search_segments,
)
from transcript_search.vtt_parser import parse_text, parse_vtt, parse_transcript


# ===========================================================================
# Sample data
# ===========================================================================

SAMPLE_VTT = """\
WEBVTT

00:00:01.000 --> 00:00:04.000
<v John Smith>Hello everyone, welcome to the quarterly review.

00:00:05.000 --> 00:00:09.000
<v Jane Doe>Thank you John. Let's start with the PTU budget discussion.

00:00:10.000 --> 00:00:14.000
<v John Smith>Sure. The PTU allocation for Q2 was 120 units.

00:00:15.000 --> 00:00:18.000
<v Jane Doe>And WSIB claims have increased by 15 percent this quarter.

00:00:19.000 --> 00:00:22.000
<v John Smith>We need to review the WSIB policy changes before the deadline.

00:00:23.000 --> 00:00:27.000
<v Jane Doe>Agreed. Any other business items on the agenda?

00:00:28.000 --> 00:00:31.000
<v John Smith>Yes, the annual compliance training is due next month.
"""

SAMPLE_VTT_NO_SPEAKER = """\
WEBVTT

00:00:00.000 --> 00:00:03.500
Welcome to the meeting.

00:00:04.000 --> 00:00:07.000
Today we will discuss the project roadmap.
"""

SAMPLE_TEXT_TRANSCRIPT = """\
[00:00:01] John Smith: Hello everyone, welcome to the quarterly review.
[00:00:05] Jane Doe: Thank you John. Let's start with the PTU budget discussion.
[00:00:10] John Smith: Sure. The PTU allocation for Q2 was 120 units.
[00:00:15] Jane Doe: And WSIB claims have increased by 15 percent this quarter.
[00:00:19] John Smith: We need to review the WSIB policy changes before the deadline.
[00:00:23] Jane Doe: Agreed. Any other business items on the agenda?
[00:00:28] John Smith: Yes, the annual compliance training is due next month.
"""

SAMPLE_MEETING: Dict = {
    "id": "meeting-001",
    "subject": "Q2 Quarterly Review",
    "startDateTime": "2024-04-15T09:00:00Z",
    "endDateTime": "2024-04-15T10:00:00Z",
    "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/...",
    "participants": {
        "organizer": {
            "identity": {
                "user": {
                    "displayName": "John Smith",
                    "id": "user-abc-123",
                }
            }
        }
    },
}

SAMPLE_TRANSCRIPT_OBJ: Dict = {
    "id": "transcript-001",
    "createdDateTime": "2024-04-15T10:05:00Z",
}


# ===========================================================================
# VTT parser tests
# ===========================================================================


class TestParseVtt:
    def test_parses_correct_number_of_segments(self):
        segments = parse_vtt(SAMPLE_VTT)
        assert len(segments) == 7

    def test_segment_indices_are_sequential(self):
        segments = parse_vtt(SAMPLE_VTT)
        assert [s.index for s in segments] == list(range(7))

    def test_first_segment_content(self):
        segments = parse_vtt(SAMPLE_VTT)
        first = segments[0]
        assert first.start_time == "00:00:01.000"
        assert first.end_time == "00:00:04.000"
        assert first.speaker == "John Smith"
        assert "Hello everyone" in first.text

    def test_speaker_extracted_correctly(self):
        segments = parse_vtt(SAMPLE_VTT)
        speakers = {s.speaker for s in segments}
        assert speakers == {"John Smith", "Jane Doe"}

    def test_no_speaker_falls_back_to_none(self):
        segments = parse_vtt(SAMPLE_VTT_NO_SPEAKER)
        assert all(s.speaker is None for s in segments)

    def test_no_html_tags_in_text(self):
        segments = parse_vtt(SAMPLE_VTT)
        for seg in segments:
            assert "<" not in seg.text
            assert ">" not in seg.text

    def test_empty_vtt_returns_empty_list(self):
        assert parse_vtt("WEBVTT\n\n") == []

    def test_malformed_vtt_skips_bad_blocks(self):
        content = "WEBVTT\n\nNot a timestamp line\nsome text\n"
        segments = parse_vtt(content)
        assert segments == []

    def test_str_representation_includes_time_and_speaker(self):
        seg = TranscriptSegment(
            index=0,
            start_time="00:01:00.000",
            end_time="00:01:05.000",
            speaker="Alice",
            text="Hello",
        )
        assert "[00:01:00.000]" in str(seg)
        assert "Alice:" in str(seg)
        assert "Hello" in str(seg)


# ===========================================================================
# Plain-text parser tests
# ===========================================================================


class TestParseText:
    def test_parses_correct_number_of_segments(self):
        segments = parse_text(SAMPLE_TEXT_TRANSCRIPT)
        assert len(segments) == 7

    def test_timestamp_extracted(self):
        segments = parse_text(SAMPLE_TEXT_TRANSCRIPT)
        assert segments[0].start_time == "00:00:01"

    def test_speaker_extracted(self):
        segments = parse_text(SAMPLE_TEXT_TRANSCRIPT)
        assert segments[0].speaker == "John Smith"

    def test_text_content_correct(self):
        segments = parse_text(SAMPLE_TEXT_TRANSCRIPT)
        assert "Hello everyone" in segments[0].text

    def test_blank_lines_skipped(self):
        content = "Line one\n\n\nLine two\n"
        segments = parse_text(content)
        assert len(segments) == 2

    def test_no_timestamp_no_speaker(self):
        content = "Just a plain line of text."
        segments = parse_text(content)
        assert len(segments) == 1
        assert segments[0].start_time is None
        assert segments[0].speaker is None
        assert segments[0].text == "Just a plain line of text."

    def test_line_without_speaker_but_with_timestamp(self):
        content = "[00:05:00] Just the text here."
        segments = parse_text(content)
        assert segments[0].start_time == "00:05:00"
        assert segments[0].speaker is None
        assert segments[0].text == "Just the text here."


# ===========================================================================
# parse_transcript dispatcher tests
# ===========================================================================


class TestParseTranscript:
    def test_dispatches_to_vtt_for_vtt_content_type(self):
        result = parse_transcript(SAMPLE_VTT, "application/vnd.ms-teams.transcript.vtt")
        assert len(result) == 7

    def test_dispatches_to_text_for_unknown_content_type(self):
        result = parse_transcript(SAMPLE_TEXT_TRANSCRIPT, "text/plain")
        assert len(result) == 7

    def test_default_content_type_uses_vtt(self):
        result = parse_transcript(SAMPLE_VTT)
        assert len(result) == 7


# ===========================================================================
# _search_segments tests
# ===========================================================================


class TestSearchSegments:
    @pytest.fixture()
    def segments(self):
        return parse_vtt(SAMPLE_VTT)

    def test_finds_single_keyword(self, segments):
        matches = _search_segments(segments, ["PTU"])
        # Two segments mention PTU
        assert len(matches) == 2

    def test_finds_keyword_case_insensitive(self, segments):
        matches_upper = _search_segments(segments, ["PTU"])
        matches_lower = _search_segments(segments, ["ptu"])
        assert len(matches_upper) == len(matches_lower)

    def test_no_match_returns_empty(self, segments):
        matches = _search_segments(segments, ["XYZNOTFOUND"])
        assert matches == []

    def test_context_lines_clipped_at_start(self, segments):
        # First segment matches 'Hello everyone'; no context_before possible
        matches = _search_segments(segments, ["Hello everyone"], context_lines=3)
        assert len(matches) == 1
        assert matches[0].context_before == []
        assert len(matches[0].context_after) == 3

    def test_context_lines_clipped_at_end(self, segments):
        # Last segment; no context_after possible
        last_keyword = "compliance training"
        matches = _search_segments(segments, [last_keyword], context_lines=5)
        assert len(matches) == 1
        assert matches[0].context_after == []

    def test_context_lines_count(self, segments):
        # "PTU budget discussion" is segment index 1
        matches = _search_segments(segments, ["PTU budget"], context_lines=2)
        assert len(matches) == 1
        match = matches[0]
        assert len(match.context_before) == 1  # Only segment 0 available before
        assert len(match.context_after) == 2

    def test_all_segments_property(self, segments):
        matches = _search_segments(segments, ["PTU budget"], context_lines=2)
        assert len(matches) == 1
        m = matches[0]
        total = len(m.context_before) + 1 + len(m.context_after)
        assert len(m.all_segments) == total

    def test_empty_segments_returns_empty(self):
        assert _search_segments([], ["PTU"]) == []

    def test_empty_keywords_returns_empty(self):
        segments = parse_vtt(SAMPLE_VTT)
        assert _search_segments(segments, []) == []

    def test_multiple_keywords_single_match_per_segment(self, segments):
        # A segment containing both "PTU" and "Q2" should produce only 1 match entry
        matches = _search_segments(segments, ["PTU", "Q2"])
        ptu_q2_segment_matches = [
            m for m in matches if "PTU" in m.matched_segment.text and "Q2" in m.matched_segment.text
        ]
        for m in ptu_q2_segment_matches:
            count = sum(
                1 for mm in matches if mm.matched_segment.index == m.matched_segment.index
            )
            assert count == 1


# ===========================================================================
# TranscriptSearchClient tests
# ===========================================================================


def _make_client(mock_graph):
    """Construct a TranscriptSearchClient backed by a mock GraphClient."""
    return TranscriptSearchClient(mock_graph, context_lines=5)


class TestListAvailableTranscripts:
    def test_returns_not_recorded_for_meeting_without_transcripts(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),  # list_meetings
            iter([]),                # list_transcripts_for_meeting → empty
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert len(results) == 1
        assert results[0].transcript_status == TranscriptStatus.NOT_RECORDED

    def test_returns_available_for_meeting_with_transcripts(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            iter([SAMPLE_TRANSCRIPT_OBJ]),
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].transcript_status == TranscriptStatus.AVAILABLE
        assert "transcript-001" in results[0].transcript_ids

    def test_access_denied_on_403(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            GraphPermissionError(403, "Access denied"),
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].transcript_status == TranscriptStatus.ACCESS_DENIED

    def test_error_status_on_generic_api_error(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            GraphAPIError(500, "Internal Server Error"),
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].transcript_status == TranscriptStatus.ERROR

    def test_no_meetings_returns_empty_list(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.return_value = iter([])
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results == []

    def test_meeting_metadata_parsed_correctly(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            iter([SAMPLE_TRANSCRIPT_OBJ]),
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        info = results[0]
        assert info.meeting_id == "meeting-001"
        assert info.subject == "Q2 Quarterly Review"
        assert info.organizer == "John Smith"
        assert info.join_url == "https://teams.microsoft.com/l/meetup-join/..."

    def test_404_transcript_endpoint_treated_as_no_transcripts(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            GraphNotFoundError(404, "Not found"),
        ]
        client = _make_client(mock_graph)
        results = client.list_available_transcripts(
            "user-1",
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].transcript_status == TranscriptStatus.NOT_RECORDED


# ===========================================================================
# TranscriptSearchClient.search_transcripts tests
# ===========================================================================


class TestSearchTranscripts:
    def _build_client_with_vtt(self, meetings, transcripts_per_meeting, vtt_content):
        """Helper: create a mock client that returns *meetings* and serves VTT."""
        mock_graph = MagicMock()
        paginated_calls = [iter(meetings)]
        for t_list in transcripts_per_meeting:
            paginated_calls.append(iter(t_list))
        mock_graph.get_paginated.side_effect = paginated_calls
        mock_graph.get_raw.return_value = vtt_content
        return _make_client(mock_graph)

    def test_returns_matches_for_keyword_found_in_transcript(self):
        client = self._build_client_with_vtt(
            [SAMPLE_MEETING],
            [[SAMPLE_TRANSCRIPT_OBJ]],
            SAMPLE_VTT,
        )
        results = client.search_transcripts(
            "user-1",
            ["PTU"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert len(results) == 1
        result = results[0]
        assert result.has_matches
        assert all("PTU" in m.matched_segment.text for m in result.matches)

    def test_no_matches_for_absent_keyword(self):
        client = self._build_client_with_vtt(
            [SAMPLE_MEETING],
            [[SAMPLE_TRANSCRIPT_OBJ]],
            SAMPLE_VTT,
        )
        results = client.search_transcripts(
            "user-1",
            ["XYZNOTFOUND"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert len(results) == 1
        assert not results[0].has_matches

    def test_not_recorded_meeting_included_with_no_matches(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            iter([]),  # No transcripts
        ]
        client = _make_client(mock_graph)
        results = client.search_transcripts(
            "user-1",
            ["PTU"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].meeting.transcript_status == TranscriptStatus.NOT_RECORDED
        assert not results[0].has_matches

    def test_permission_error_during_download_marks_meeting(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            iter([SAMPLE_TRANSCRIPT_OBJ]),
        ]
        mock_graph.get_raw.side_effect = GraphPermissionError(403, "Forbidden")
        client = _make_client(mock_graph)
        results = client.search_transcripts(
            "user-1",
            ["PTU"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].meeting.transcript_status == TranscriptStatus.ACCESS_DENIED
        assert not results[0].has_matches

    def test_api_error_during_download_marks_meeting(self):
        mock_graph = MagicMock()
        mock_graph.get_paginated.side_effect = [
            iter([SAMPLE_MEETING]),
            iter([SAMPLE_TRANSCRIPT_OBJ]),
        ]
        mock_graph.get_raw.side_effect = GraphAPIError(500, "Server Error")
        client = _make_client(mock_graph)
        results = client.search_transcripts(
            "user-1",
            ["PTU"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        assert results[0].meeting.transcript_status == TranscriptStatus.ERROR

    def test_multiple_keywords_matched(self):
        client = self._build_client_with_vtt(
            [SAMPLE_MEETING],
            [[SAMPLE_TRANSCRIPT_OBJ]],
            SAMPLE_VTT,
        )
        results = client.search_transcripts(
            "user-1",
            ["PTU", "WSIB"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
        )
        # PTU appears in 2 segments, WSIB in 2 segments → 4 matches
        assert len(results[0].matches) == 4

    def test_context_lines_returned(self):
        client = self._build_client_with_vtt(
            [SAMPLE_MEETING],
            [[SAMPLE_TRANSCRIPT_OBJ]],
            SAMPLE_VTT,
        )
        results = client.search_transcripts(
            "user-1",
            ["PTU budget"],
            datetime(2024, 4, 1, tzinfo=timezone.utc),
            datetime(2024, 4, 30, tzinfo=timezone.utc),
            context_lines=2,
        )
        match = results[0].matches[0]
        # "PTU budget discussion" is the 2nd segment (index 1); only 1 before it
        assert len(match.context_before) <= 2
        assert len(match.context_after) <= 2


# ===========================================================================
# TranscriptSearchClient.search_meeting_transcript tests
# ===========================================================================


class TestSearchMeetingTranscript:
    def test_direct_search_finds_keyword(self):
        mock_graph = MagicMock()
        mock_graph.get_raw.return_value = SAMPLE_VTT
        client = _make_client(mock_graph)
        matches = client.search_meeting_transcript(
            "user-1", "meeting-001", "transcript-001", ["WSIB"]
        )
        assert len(matches) == 2
        for m in matches:
            assert "WSIB" in m.matched_segment.text

    def test_direct_search_propagates_permission_error(self):
        mock_graph = MagicMock()
        mock_graph.get_raw.side_effect = GraphPermissionError(403, "Forbidden")
        client = _make_client(mock_graph)
        with pytest.raises(GraphPermissionError):
            client.search_meeting_transcript(
                "user-1", "meeting-001", "transcript-001", ["WSIB"]
            )


# ===========================================================================
# TranscriptSearchResult model tests
# ===========================================================================


class TestTranscriptSearchResultModel:
    def test_has_matches_false_when_empty(self):
        info = MeetingTranscriptInfo(
            meeting_id="x",
            subject="Test",
            start_time=None,
            end_time=None,
            organizer=None,
            join_url=None,
            transcript_status=TranscriptStatus.NOT_RECORDED,
        )
        result = TranscriptSearchResult(meeting=info)
        assert not result.has_matches

    def test_has_matches_true_when_populated(self):
        seg = TranscriptSegment(index=0, start_time=None, end_time=None, speaker=None, text="hi")
        match = TranscriptMatch(keyword="hi", matched_segment=seg)
        info = MeetingTranscriptInfo(
            meeting_id="x",
            subject="Test",
            start_time=None,
            end_time=None,
            organizer=None,
            join_url=None,
            transcript_status=TranscriptStatus.AVAILABLE,
            transcript_ids=["t1"],
        )
        result = TranscriptSearchResult(meeting=info, matches=[match])
        assert result.has_matches
