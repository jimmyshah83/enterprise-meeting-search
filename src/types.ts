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

export type DomainName = 'calendar' | 'transcripts' | 'chats' | 'email' | 'files';

export type DomainStatus =
  | { status: 'success'; resultCount: number }
  | { status: 'timeout' }
  | { status: 'error'; error: string };

// ---------------------------------------------------------------------------
// Diagnostic types
// ---------------------------------------------------------------------------

/**
 * Likelihood level for a probable cause or suggestion, ranked from highest
 * to lowest confidence.
 */
export type Likelihood = 'high' | 'medium' | 'low';

/**
 * A single probable cause explaining why no results were found.
 */
export interface ProbableCause {
  /** Short identifier for the cause (machine-readable). */
  code: string;
  /** Human-readable explanation. */
  description: string;
  /** Estimated likelihood that this cause applies. */
  likelihood: Likelihood;
}

/**
 * A single actionable suggestion the user can act on to retry their search.
 */
export interface RetrySuggestion {
  /** Short identifier for the suggestion (machine-readable). */
  code: string;
  /** Human-readable action the user should take. */
  action: string;
  /**
   * Which additional piece of information would help narrow down results
   * (e.g. "date", "organizer", "alternate title").
   */
  informationNeeded?: string;
  /** Estimated usefulness / priority of this suggestion. */
  likelihood: Likelihood;
}

/**
 * Structured diagnostic output returned when all search domains yield no
 * results.  Designed for both direct API JSON consumption and conversion to
 * human-readable conversational text via {@link formatDiagnosticText}.
 */
export interface DiagnosticSummary {
  /** One-sentence summary suitable for use as an API error message. */
  message: string;
  /**
   * Probable causes ranked from most to least likely.
   * Each entry explains a potential reason the search returned empty.
   */
  probableCauses: ProbableCause[];
  /**
   * Actionable retry suggestions ranked from most to least impactful.
   * Each entry tells the user something concrete to try next.
   */
  retrySuggestions: RetrySuggestion[];
  /** Per-domain execution status (success / timeout / error). */
  domainStatuses: Record<DomainName, DomainStatus>;
  /**
   * @deprecated Use `probableCauses` and `retrySuggestions` instead.
   * Kept for backwards compatibility with early orchestrator usage.
   */
  reasons: string[];
  /**
   * @deprecated Use `retrySuggestions` instead.
   * Kept for backwards compatibility with early orchestrator usage.
   */
  suggestions: string[];
}

/** The top-level response returned by the orchestrator. */
export interface SearchResponse {
  results: UnifiedMeetingResult[];
  /** Present only when results is empty. */
  diagnostics?: DiagnosticSummary;
  totalResults: number;
  searchDurationMs: number;
}
