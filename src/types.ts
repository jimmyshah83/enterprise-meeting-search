/**
 * Shared types for the enterprise meeting search system.
 */

/** Filter criteria accepted by the unified search request. */
export interface SearchRequest {
  /** Keywords to match across all fields. */
  keywords?: string[];
  /** Inclusive date range filter. */
  dateRange?: { start: Date; end: Date };
  /** Email addresses of attendees or organizers to filter by. */
  attendees?: string[];
  /** Organizer email address filter. */
  organizer?: string;
  /** Meeting type identifiers (e.g. "Teams", "Zoom", "InPerson"). */
  meetingType?: string[];
  /** Per-domain timeout in milliseconds (default: 5000). */
  domainTimeoutMs?: number;
  /** Overall orchestration timeout in milliseconds (default: 10000). */
  overallTimeoutMs?: number;
}

// ---------------------------------------------------------------------------
// Domain result types
// ---------------------------------------------------------------------------

export interface CalendarEvent {
  id: string;
  title: string;
  startTime: Date;
  endTime: Date;
  organizer: string;
  attendees: string[];
  meetingType: string;
  source: 'calendar';
}

export interface TranscriptResult {
  id: string;
  meetingId: string;
  content: string;
  speakers: string[];
  timestamp: Date;
  source: 'transcript';
}

export interface ChatMessage {
  id: string;
  /** The meeting this chat thread belongs to, if known. */
  meetingId?: string;
  threadId: string;
  content: string;
  sender: string;
  timestamp: Date;
  source: 'chat';
}

export interface EmailResult {
  id: string;
  subject: string;
  from: string;
  to: string[];
  body: string;
  timestamp: Date;
  /** Associated meeting ID derived from calendar invite references. */
  meetingId?: string;
  source: 'email';
}

export interface FileResult {
  id: string;
  name: string;
  url: string;
  /** Associated meeting ID if the file was shared during a meeting. */
  meetingId?: string;
  modifiedBy: string;
  modifiedAt: Date;
  source: 'file';
}

/** Union of all domain-specific result types. */
export type DomainResult =
  | CalendarEvent
  | TranscriptResult
  | ChatMessage
  | EmailResult
  | FileResult;

// ---------------------------------------------------------------------------
// Aggregated / unified types
// ---------------------------------------------------------------------------

/**
 * A single unified search result entry grouped by meeting / calendar event.
 * All artifacts (transcripts, chats, emails, files) are correlated to the
 * calendar event via their `meetingId` field.
 */
export interface UnifiedMeetingResult {
  meetingId: string;
  calendarEvent?: CalendarEvent;
  transcripts: TranscriptResult[];
  chats: ChatMessage[];
  emails: EmailResult[];
  files: FileResult[];
}

/**
 * Diagnostic information returned when no results are found across all
 * domains, helping the caller understand why the search came up empty.
 */
export interface DiagnosticSummary {
  message: string;
  reasons: string[];
  suggestions: string[];
  domainStatuses: Record<DomainName, DomainStatus>;
}

export type DomainName = 'calendar' | 'transcripts' | 'chats' | 'email' | 'files';

export type DomainStatus =
  | { status: 'success'; resultCount: number }
  | { status: 'timeout' }
  | { status: 'error'; error: string };

/** The top-level response returned by the orchestrator. */
export interface SearchResponse {
  results: UnifiedMeetingResult[];
  /** Present only when results is empty. */
  diagnostics?: DiagnosticSummary;
  totalResults: number;
  searchDurationMs: number;
}
