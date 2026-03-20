/**
 * Structured representation of a calendar meeting or Teams event.
 */
export interface Meeting {
  /** Unique event identifier from Microsoft Graph */
  id: string;
  /** Meeting subject / title */
  title: string;
  /** ISO 8601 start datetime (UTC) */
  datetime: string;
  /** ISO 8601 end datetime (UTC) */
  endDatetime: string;
  /** Meeting organizer details */
  organizer: Participant;
  /** All invited attendees */
  attendees: AttendeeInfo[];
  /** Teams / Skype join URL, null when not an online meeting */
  joinUrl: string | null;
  /** Whether a transcript exists for this meeting */
  hasTranscript: boolean;
  /** Whether a recording exists for this meeting */
  hasRecording: boolean;
  /** Plain-text body / description of the event */
  body?: string;
  /** Location string (room, address, or online) */
  location?: string;
  /** Whether the event spans all day */
  isAllDay: boolean;
  /** Raw categories/tags assigned to the event */
  categories: string[];
}

/** A person (organizer or attendee) */
export interface Participant {
  name: string;
  email: string;
}

/** An attendee with their RSVP status */
export interface AttendeeInfo extends Participant {
  /** accepted | declined | tentative | none */
  status: string;
}

/**
 * Options accepted by CalendarSearchClient.searchMeetings().
 */
export interface SearchOptions {
  /** Free-text keyword to match against title, body, and metadata */
  keyword?: string;
  /** Start of date range (ISO 8601 or JS Date) */
  dateFrom?: string | Date;
  /** End of date range (ISO 8601 or JS Date) */
  dateTo?: string | Date;
  /** Filter by organizer name or email (substring match) */
  organizer?: string;
  /** Filter by attendee name or email (substring match) */
  attendee?: string;
  /** Enable fuzzy / approximate keyword matching (default: false) */
  fuzzy?: boolean;
  /** Max events to fetch per page (1–999, default: 50) */
  pageSize?: number;
  /**
   * Opaque pagination token returned by a previous call.
   * Pass this back to retrieve the next page.
   */
  nextLink?: string;
}

/**
 * Result returned by CalendarSearchClient.searchMeetings().
 */
export interface SearchResult {
  /** Meetings matching the search criteria */
  meetings: Meeting[];
  /**
   * Pagination token for the next page of results.
   * Undefined when there are no more results.
   */
  nextLink?: string;
  /** Total number of items reported by Graph (may not always be present) */
  totalCount?: number;
}

/**
 * Configuration required to build a Graph API client.
 */
export interface GraphClientConfig {
  /** Azure AD tenant ID */
  tenantId: string;
  /** Application (client) ID */
  clientId: string;
  /** Client secret (for client-credentials flow) */
  clientSecret: string;
}
