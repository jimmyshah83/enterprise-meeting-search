/**
 * Structured search parameters produced by the natural language parser.
 */
export interface DateRange {
  start: Date;
  end: Date;
}

export type MeetingType = 'teams' | 'external' | 'channel' | 'unknown';

export interface SearchParams {
  /** Optional date range resolved from natural language hints. */
  date_range?: DateRange;
  /** Attendee names extracted from the hint. */
  attendees: string[];
  /** Remaining keyword terms not consumed by other parsers. */
  keywords: string[];
  /** Meeting type hint, if detected. */
  meeting_type: MeetingType;
}
