import { Client } from '@microsoft/microsoft-graph-client';
import {
  Meeting,
  Participant,
  AttendeeInfo,
  SearchOptions,
  SearchResult,
} from './types';

/** Default page size used when the caller does not specify one. */
const DEFAULT_PAGE_SIZE = 50;

/** Maximum page size accepted by the Graph calendarView endpoint. */
const MAX_PAGE_SIZE = 999;

/**
 * Transforms a raw Graph API event object into the structured {@link Meeting}
 * interface consumed by the rest of the application.
 */
function transformEvent(event: Record<string, unknown>): Meeting {
  const organizer = event['organizer'] as
    | Record<string, Record<string, string>>
    | undefined;
  const organizerAddr = organizer?.emailAddress ?? {};

  const rawAttendees = (event['attendees'] as Array<
    Record<string, unknown>
  >) ?? [];
  const attendees: AttendeeInfo[] = rawAttendees.map((a) => {
    const addr = (a['emailAddress'] as Record<string, string>) ?? {};
    const statusRecord = (a['status'] as Record<string, string>) ?? {};
    return {
      name: addr['name'] ?? '',
      email: addr['address'] ?? '',
      status: statusRecord['response'] ?? 'none',
    };
  });

  const onlineMeeting = event['onlineMeeting'] as
    | Record<string, string>
    | undefined;
  const joinUrl = onlineMeeting?.joinUrl ?? null;

  const bodyObj = event['body'] as Record<string, string> | undefined;
  const bodyText = bodyObj?.content ?? '';

  const startObj = event['start'] as Record<string, string> | undefined;
  const endObj = event['end'] as Record<string, string> | undefined;

  const locationObj = event['location'] as Record<string, string> | undefined;
  const locationText = locationObj?.displayName ?? '';

  const categories = (event['categories'] as string[]) ?? [];

  // Transcript / recording presence is inferred from the event body or
  // the onlineMeetingProvider value.  A definitive check would require
  // separate Graph calls to /communications/callRecords; here we use the
  // presence of the joinUrl (indicating an online meeting that *could*
  // have a recording) and look for known keywords in the body.
  const bodyLower = bodyText.toLowerCase();
  const hasRecording =
    bodyLower.includes('recording') || bodyLower.includes('recorded');
  const hasTranscript =
    bodyLower.includes('transcript') || bodyLower.includes('transcription');

  return {
    id: (event['id'] as string) ?? '',
    title: (event['subject'] as string) ?? '',
    datetime: startObj?.dateTime ?? (event['createdDateTime'] as string) ?? '',
    endDatetime: endObj?.dateTime ?? '',
    organizer: {
      name: organizerAddr['name'] ?? '',
      email: organizerAddr['address'] ?? '',
    } as Participant,
    attendees,
    joinUrl,
    hasTranscript,
    hasRecording,
    body: bodyText,
    location: locationText,
    isAllDay: (event['isAllDay'] as boolean) ?? false,
    categories,
  };
}

/**
 * Builds an OData `$filter` expression from the caller-supplied options.
 * Returns an empty string when no filter constraints apply.
 */
function buildFilter(options: SearchOptions): string {
  const clauses: string[] = [];

  if (options.organizer) {
    const escaped = options.organizer.replace(/'/g, "''");
    clauses.push(
      `(organizer/emailAddress/address eq '${escaped}' or contains(organizer/emailAddress/name, '${escaped}'))`,
    );
  }

  if (options.attendee) {
    const escaped = options.attendee.replace(/'/g, "''");
    clauses.push(
      `attendees/any(a: a/emailAddress/address eq '${escaped}' or contains(a/emailAddress/name, '${escaped}'))`,
    );
  }

  return clauses.join(' and ');
}

/**
 * Checks whether a {@link Meeting} matches the keyword contained in
 * `options.keyword` using either exact (substring) or fuzzy matching.
 *
 * Fuzzy matching is performed by splitting the keyword into tokens and
 * requiring that every token appears (case-insensitively) in the
 * concatenated searchable text.
 */
function matchesKeyword(meeting: Meeting, options: SearchOptions): boolean {
  if (!options.keyword) return true;

  const haystack = [
    meeting.title,
    meeting.body ?? '',
    meeting.organizer.name,
    meeting.organizer.email,
    meeting.location ?? '',
    meeting.categories.join(' '),
    ...meeting.attendees.map((a) => `${a.name} ${a.email}`),
  ]
    .join(' ')
    .toLowerCase();

  if (options.fuzzy) {
    // Fuzzy: every whitespace-delimited token must appear somewhere
    const tokens = options.keyword.toLowerCase().split(/\s+/).filter(Boolean);
    return tokens.every((token) => haystack.includes(token));
  }

  // Exact: the full keyword phrase must appear as a substring
  return haystack.includes(options.keyword.toLowerCase());
}

/**
 * Core search client.  Instantiate with a Microsoft Graph {@link Client} and
 * then call {@link searchMeetings} to query the calendar.
 *
 * @example
 * ```typescript
 * import { createGraphClient } from './graphClient';
 * import { CalendarSearchClient } from './calendarSearch';
 *
 * const graphClient = createGraphClient({ tenantId, clientId, clientSecret });
 * const searcher = new CalendarSearchClient(graphClient, 'user@example.com');
 *
 * const result = await searcher.searchMeetings({
 *   keyword: 'PTU',
 *   dateFrom: '2024-01-01',
 *   dateTo: '2024-03-31',
 * });
 * console.log(result.meetings);
 * ```
 */
export class CalendarSearchClient {
  /**
   * @param client   - Authenticated Graph client (from `createGraphClient`)
   * @param userId   - UPN or object ID of the mailbox to search.
   *                   Use `'me'` for delegated (signed-in user) access.
   */
  constructor(
    private readonly client: Client,
    private readonly userId: string = 'me',
  ) {}

  /**
   * Searches the user's calendar for events that match the supplied options.
   *
   * - When `dateFrom` / `dateTo` are provided the `calendarView` endpoint is
   *   used (most efficient for date-bounded queries).
   * - When only a keyword is provided, `events` is queried with `$search`.
   * - Organizer / attendee constraints are applied via `$filter`.
   * - Keyword post-filtering (including fuzzy) is applied in-process after
   *   fetching each page so that server-side results are narrowed further.
   *
   * @returns A {@link SearchResult} containing the matched meetings and an
   *          optional `nextLink` token for fetching subsequent pages.
   */
  async searchMeetings(options: SearchOptions = {}): Promise<SearchResult> {
    const pageSize = Math.min(
      options.pageSize ?? DEFAULT_PAGE_SIZE,
      MAX_PAGE_SIZE,
    );

    const filter = buildFilter(options);
    const select = [
      'id',
      'subject',
      'start',
      'end',
      'organizer',
      'attendees',
      'body',
      'onlineMeeting',
      'location',
      'isAllDay',
      'categories',
      'createdDateTime',
    ].join(',');

    // If a nextLink is provided we follow it directly (pagination).
    if (options.nextLink) {
      return this.fetchPage(options.nextLink, options);
    }

    let request: ReturnType<Client['api']>;

    if (options.dateFrom || options.dateTo) {
      // calendarView requires both startDateTime and endDateTime
      const startDateTime = options.dateFrom
        ? new Date(options.dateFrom).toISOString()
        : new Date(0).toISOString();
      const endDateTime = options.dateTo
        ? new Date(options.dateTo).toISOString()
        : new Date('2099-12-31').toISOString();

      const path =
        this.userId === 'me'
          ? '/me/calendarView'
          : `/users/${this.userId}/calendarView`;

      request = this.client
        .api(path)
        .query({ startDateTime, endDateTime })
        .top(pageSize)
        .select(select)
        .header('Prefer', 'outlook.timezone="UTC"');
    } else {
      const path =
        this.userId === 'me' ? '/me/events' : `/users/${this.userId}/events`;

      request = this.client
        .api(path)
        .top(pageSize)
        .select(select)
        .orderby('start/dateTime DESC')
        .header('Prefer', 'outlook.timezone="UTC"');

      // $search and $filter cannot be combined on the events endpoint;
      // use $search when a keyword is present, $filter otherwise.
      if (options.keyword && !options.fuzzy) {
        request = request.search(`"${options.keyword}"`);
      } else if (filter) {
        request = request.filter(filter);
      }
    }

    // When using calendarView with additional OData filters
    if ((options.dateFrom || options.dateTo) && filter) {
      request = request.filter(filter);
    }

    const response = await request.get();
    return this.processResponse(response, options);
  }

  /** Follows a raw Graph `@odata.nextLink` URL. */
  private async fetchPage(
    nextLink: string,
    options: SearchOptions,
  ): Promise<SearchResult> {
    const response = await this.client.api(nextLink).get();
    return this.processResponse(response, options);
  }

  /**
   * Converts a raw Graph response into a {@link SearchResult}, applying
   * in-process keyword filtering.
   */
  private processResponse(
    response: Record<string, unknown>,
    options: SearchOptions,
  ): SearchResult {
    const rawEvents = (response['value'] as Array<Record<string, unknown>>) ?? [];

    let meetings = rawEvents.map(transformEvent);

    // Apply in-process keyword filtering (handles fuzzy mode and also
    // provides an extra safety net for the exact-match path).
    if (options.keyword) {
      meetings = meetings.filter((m) => matchesKeyword(m, options));
    }

    const graphNextLink = response['@odata.nextLink'] as string | undefined;
    const odataCount = response['@odata.count'] as number | undefined;

    return {
      meetings,
      nextLink: graphNextLink,
      totalCount: odataCount,
    };
  }
}
