/**
 * Confidence level for a reconstructed meeting element.
 * - high: sourced directly from a reliable primary source (e.g. transcript)
 * - medium: sourced from a secondary source (e.g. email thread)
 * - low: sourced from a tertiary source (e.g. chat messages)
 * - inferred: derived indirectly from context with low certainty
 */
export type ConfidenceLevel = 'high' | 'medium' | 'low' | 'inferred';

/** The domain from which a reconstructed element was derived. */
export type SourceType = 'transcript' | 'chat' | 'email' | 'file';

// ---------------------------------------------------------------------------
// Input types – search results produced by the cross-domain search orchestrator
// ---------------------------------------------------------------------------

/** A meeting transcript returned by the orchestrator. */
export interface TranscriptResult {
  source: 'transcript';
  meetingId?: string;
  title?: string;
  date?: string;
  participants?: string[];
  /** Full transcript text. */
  content: string;
  /** Individual speaker turns within the transcript. */
  segments?: Array<{
    speaker: string;
    text: string;
    timestamp?: string;
  }>;
}

/** A single chat message returned by the orchestrator. */
export interface ChatMessage {
  source: 'chat';
  channelId?: string;
  channelName?: string;
  messageId?: string;
  sender: string;
  timestamp: string;
  text: string;
  /** Optional reference to the meeting this message is associated with. */
  meetingReference?: string;
}

/** An email thread returned by the orchestrator. */
export interface EmailThread {
  source: 'email';
  subject?: string;
  date?: string;
  from?: string;
  to?: string[];
  body: string;
  threadId?: string;
}

/** A file (e.g. meeting notes document) returned by the orchestrator. */
export interface FileResult {
  source: 'file';
  filename: string;
  path?: string;
  content?: string;
  meetingDate?: string;
  author?: string;
}

/** Union of all domain-specific search result types. */
export type SearchResult = TranscriptResult | ChatMessage | EmailThread | FileResult;

/**
 * The correlated set of search results produced by the cross-domain
 * search orchestrator and consumed by the reconstruction engine.
 */
export interface CorrelatedSearchResults {
  query: string;
  results: SearchResult[];
}

// ---------------------------------------------------------------------------
// Output types – the structured meeting reconstruction
// ---------------------------------------------------------------------------

/**
 * A single reconstructed meeting element with provenance metadata.
 */
export interface ReconstructedElement<T> {
  /** The reconstructed value. */
  value: T;
  /** How confident we are in this element. */
  confidence: ConfidenceLevel;
  /** The domain from which this element was derived. */
  source: SourceType;
  /** Optional human-readable detail about the exact source (e.g. email subject line). */
  sourceDetails?: string;
}

/**
 * A fully structured meeting reconstruction produced by
 * {@link MeetingReconstructionEngine}.  Every element is nullable to
 * support partial data scenarios; missing elements are explained in `gaps`.
 */
export interface MeetingReconstruction {
  /** Likely meeting title. */
  title: ReconstructedElement<string> | null;
  /** Meeting date/time string as found in the source data. */
  dateTime: ReconstructedElement<string> | null;
  /** List of meeting attendees / participants. */
  attendees: ReconstructedElement<string[]> | null;
  /** Key discussion topics identified in the meeting. */
  topics: ReconstructedElement<string[]> | null;
  /** Decisions made during the meeting. */
  decisions: ReconstructedElement<string[]> | null;
  /** Action items assigned during the meeting. */
  actionItems: ReconstructedElement<string[]> | null;
  /** Human-readable descriptions of information that could not be reconstructed. */
  gaps: string[];
  /** Aggregate confidence across all reconstructed elements. */
  overallConfidence: ConfidenceLevel;
}
