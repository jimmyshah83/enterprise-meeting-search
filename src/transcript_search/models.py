"""
Data models for the meeting transcript search module.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import List, Optional


class TranscriptStatus(str, Enum):
    """Status of a transcript for a given meeting."""

    AVAILABLE = "available"
    NOT_RECORDED = "not recorded"
    ACCESS_DENIED = "access denied"
    ERROR = "error"


@dataclass
class TranscriptSegment:
    """A single timed segment from a transcript."""

    index: int
    start_time: Optional[str]
    end_time: Optional[str]
    speaker: Optional[str]
    text: str

    def __str__(self) -> str:
        speaker_prefix = f"{self.speaker}: " if self.speaker else ""
        time_prefix = f"[{self.start_time}] " if self.start_time else ""
        return f"{time_prefix}{speaker_prefix}{self.text}"


@dataclass
class TranscriptMatch:
    """A keyword match inside a transcript, with surrounding context lines."""

    keyword: str
    matched_segment: TranscriptSegment
    context_before: List[TranscriptSegment] = field(default_factory=list)
    context_after: List[TranscriptSegment] = field(default_factory=list)

    @property
    def all_segments(self) -> List[TranscriptSegment]:
        """Return context_before + matched_segment + context_after."""
        return self.context_before + [self.matched_segment] + self.context_after


@dataclass
class MeetingTranscriptInfo:
    """Metadata about a meeting and its transcript availability."""

    meeting_id: str
    subject: str
    start_time: Optional[datetime]
    end_time: Optional[datetime]
    organizer: Optional[str]
    join_url: Optional[str]
    transcript_status: TranscriptStatus
    transcript_ids: List[str] = field(default_factory=list)
    error_message: Optional[str] = None


@dataclass
class TranscriptSearchResult:
    """Aggregated search result for a single meeting."""

    meeting: MeetingTranscriptInfo
    matches: List[TranscriptMatch] = field(default_factory=list)

    @property
    def has_matches(self) -> bool:
        return bool(self.matches)
