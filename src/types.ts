/**
 * Shared types for the enterprise meeting search system.
 */

/** Filter criteria accepted by the unified search request. */
export interface SearchRequest {
  keywords?: string[];
  dateRange?: { start: Date; end: Date };
  attendees?: string[];
  organizer?: string;
  meetingType?: string[];
  domainTimeoutMs?: number;
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
  meetingId?: string;
  source: 'email';
}

export interface FileResult {
  id: string;
  name: string;
  url: string;
  meetingId?: string;
  modifiedBy: string;
  modifiedAt: Date;
  source: 'file';
}

export type DomainResult =
  | CalendarEvent
  | TranscriptResult
  | ChatMessage
  | EmailResult
  | FileResult;

// ---------------------------------------------------------------------------
// Aggregated / unified types
// ---------------------------------------------------------------------------

export interface UnifiedMeetingResult {
  meetingId: string;
  calendarEvent?: CalendarEvent;
  transcripts: TranscriptResult[];
  chats: ChatMessage[];
  emails: EmailResult[];
  files: FileResult[];
}

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

export interface SearchResponse {
  results: UnifiedMeetingResult[];
  diagnostics?: DiagnosticSummary;
  totalResults: number;
  searchDurationMs: number;
}

// ---------------------------------------------------------------------------
// Meeting reconstruction types
// ---------------------------------------------------------------------------

export type ConfidenceLevel = 'high' | 'medium' | 'low' | 'inferred';
export type ReconstructionSourceType = 'transcript' | 'chat' | 'email' | 'file' | 'calendar';

export interface ReconstructedElement<T> {
  value: T;
  confidence: ConfidenceLevel;
  source: ReconstructionSourceType;
  sourceDetails?: string;
}

export interface ActionItem {
  description: string;
  assignee?: string;
  dueDate?: string;
}

export interface MeetingReconstruction {
  meetingId: string;
  title: ReconstructedElement<string> | null;
  dateTime: ReconstructedElement<string> | null;
  attendees: ReconstructedElement<string[]> | null;
  topics: ReconstructedElement<string[]> | null;
  decisions: ReconstructedElement<string[]> | null;
  actionItems: ReconstructedElement<ActionItem[]> | null;
  overallConfidence: ConfidenceLevel;
  gaps: string[];
  reconstructedAt: Date;
}

// ---------------------------------------------------------------------------
// API-specific types
// ---------------------------------------------------------------------------

export type SearchStatus = 'pending' | 'running' | 'completed' | 'failed';

export interface AsyncSearchRecord {
  searchId: string;
  status: SearchStatus;
  request: SearchRequest;
  response?: SearchResponse;
  error?: string;
  createdAt: Date;
  completedAt?: Date;
}

export type ReconstructionStatus = 'pending' | 'running' | 'completed' | 'failed';

export interface AsyncReconstructionRecord {
  reconstructionId: string;
  meetingId: string;
  status: ReconstructionStatus;
  result?: MeetingReconstruction;
  error?: string;
  createdAt: Date;
  completedAt?: Date;
}
