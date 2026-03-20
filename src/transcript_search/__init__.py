"""
Meeting transcript search package for Microsoft Teams via Microsoft Graph API.
"""

from .transcript_search import TranscriptSearchClient
from .models import (
    MeetingTranscriptInfo,
    TranscriptSegment,
    TranscriptMatch,
    TranscriptSearchResult,
    TranscriptStatus,
)

__all__ = [
    "TranscriptSearchClient",
    "MeetingTranscriptInfo",
    "TranscriptSegment",
    "TranscriptMatch",
    "TranscriptSearchResult",
    "TranscriptStatus",
]
