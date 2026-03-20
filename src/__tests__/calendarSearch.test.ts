/**
 * Integration tests for CalendarSearchClient.
 *
 * All Microsoft Graph HTTP calls are mocked so that these tests run without
 * any live network access or credentials.
 */

import { CalendarSearchClient } from '../calendarSearch';
import { Client } from '@microsoft/microsoft-graph-client';
import { Meeting, SearchResult } from '../types';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build a minimal raw Graph event object for use in mock responses. */
function makeRawEvent(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: 'event-1',
    subject: 'Default Meeting',
    start: { dateTime: '2024-02-01T09:00:00', timeZone: 'UTC' },
    end: { dateTime: '2024-02-01T10:00:00', timeZone: 'UTC' },
    isAllDay: false,
    categories: [],
    organizer: {
      emailAddress: { name: 'Alice Smith', address: 'alice@example.com' },
    },
    attendees: [
      {
        emailAddress: { name: 'Bob Jones', address: 'bob@example.com' },
        status: { response: 'accepted' },
      },
    ],
    body: { content: 'Discuss the PTU deployment plan.', contentType: 'text' },
    onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc' },
    location: { displayName: 'Teams' },
    ...overrides,
  };
}

/** Build a raw Graph event that has no online meeting (room booking etc.). */
function makeRawOfflineEvent(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return makeRawEvent({
    id: 'event-offline',
    subject: 'Offline Meeting',
    onlineMeeting: null,
    body: { content: 'In-person sync.', contentType: 'text' },
    ...overrides,
  });
}

/**
 * Creates a CalendarSearchClient whose underlying Graph `api()` method is
 * replaced by a jest mock that returns `mockResponse`.
 */
function makeClient(
  mockResponse: Record<string, unknown>,
  userId = 'me',
): { client: CalendarSearchClient; apiMock: jest.Mock } {
  // We build a chainable mock: every method on the request builder returns
  // `this` so that chaining works; the final `.get()` resolves with the response.
  const requestBuilder: Record<string, jest.Mock> = {};
  const chainableMethod = (): Record<string, jest.Mock> => requestBuilder;
  const getMock = jest.fn().mockResolvedValue(mockResponse);

  for (const method of [
    'query',
    'top',
    'select',
    'filter',
    'search',
    'orderby',
    'header',
  ]) {
    requestBuilder[method] = jest.fn().mockReturnValue(requestBuilder);
  }
  requestBuilder['get'] = getMock;

  const apiMock = jest.fn().mockReturnValue(requestBuilder);
  const graphClient = { api: apiMock } as unknown as Client;

  return { client: new CalendarSearchClient(graphClient, userId), apiMock };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('CalendarSearchClient.searchMeetings()', () => {
  // -------------------------------------------------------------------------
  // Basic event transformation
  // -------------------------------------------------------------------------

  describe('event transformation', () => {
    it('maps a raw Graph event to a structured Meeting object', async () => {
      const { client } = makeClient({ value: [makeRawEvent()] });

      const result: SearchResult = await client.searchMeetings({});

      expect(result.meetings).toHaveLength(1);
      const meeting: Meeting = result.meetings[0];

      expect(meeting.id).toBe('event-1');
      expect(meeting.title).toBe('Default Meeting');
      expect(meeting.datetime).toBe('2024-02-01T09:00:00');
      expect(meeting.endDatetime).toBe('2024-02-01T10:00:00');
      expect(meeting.isAllDay).toBe(false);
      expect(meeting.categories).toEqual([]);

      expect(meeting.organizer).toEqual({
        name: 'Alice Smith',
        email: 'alice@example.com',
      });

      expect(meeting.attendees).toHaveLength(1);
      expect(meeting.attendees[0]).toEqual({
        name: 'Bob Jones',
        email: 'bob@example.com',
        status: 'accepted',
      });

      expect(meeting.joinUrl).toBe(
        'https://teams.microsoft.com/l/meetup-join/abc',
      );
      expect(meeting.body).toContain('PTU');
      expect(meeting.location).toBe('Teams');
    });

    it('sets joinUrl to null for offline/room events', async () => {
      const { client } = makeClient({ value: [makeRawOfflineEvent()] });
      const result = await client.searchMeetings({});
      expect(result.meetings[0].joinUrl).toBeNull();
    });

    it('detects recording keyword in body', async () => {
      const event = makeRawEvent({
        body: { content: 'This meeting was recorded.', contentType: 'text' },
      });
      const { client } = makeClient({ value: [event] });
      const result = await client.searchMeetings({});
      expect(result.meetings[0].hasRecording).toBe(true);
      expect(result.meetings[0].hasTranscript).toBe(false);
    });

    it('detects transcript keyword in body', async () => {
      const event = makeRawEvent({
        body: {
          content: 'A transcript is available.',
          contentType: 'text',
        },
      });
      const { client } = makeClient({ value: [event] });
      const result = await client.searchMeetings({});
      expect(result.meetings[0].hasTranscript).toBe(true);
    });

    it('handles missing optional fields gracefully', async () => {
      const sparse = {
        id: 'sparse-1',
        subject: 'Sparse',
        start: { dateTime: '2024-03-01T08:00:00' },
        end: {},
        organizer: {},
        attendees: [],
        body: {},
        isAllDay: false,
        categories: [],
      };
      const { client } = makeClient({ value: [sparse] });
      const result = await client.searchMeetings({});
      const m = result.meetings[0];
      expect(m.joinUrl).toBeNull();
      expect(m.organizer).toEqual({ name: '', email: '' });
      expect(m.attendees).toEqual([]);
      expect(m.body).toBe('');
    });

    it('maps categories to the meeting object', async () => {
      const event = makeRawEvent({ categories: ['Project Alpha', 'Finance'] });
      const { client } = makeClient({ value: [event] });
      const result = await client.searchMeetings({});
      expect(result.meetings[0].categories).toEqual(['Project Alpha', 'Finance']);
    });
  });

  // -------------------------------------------------------------------------
  // Keyword search (exact)
  // -------------------------------------------------------------------------

  describe('keyword search – exact', () => {
    it('returns meetings whose title contains the keyword', async () => {
      const events = [
        makeRawEvent({ id: 'e1', subject: 'PTU Review', body: { content: '' } }),
        makeRawEvent({ id: 'e2', subject: 'Team Standup', body: { content: '' } }),
      ];
      const { client } = makeClient({ value: events });
      const result = await client.searchMeetings({ keyword: 'PTU' });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });

    it('returns meetings whose body contains the keyword', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          subject: 'Sync',
          body: { content: 'We need to discuss WSIB processes.', contentType: 'text' },
        }),
        makeRawEvent({
          id: 'e2',
          subject: 'Other',
          body: { content: 'Nothing relevant here.', contentType: 'text' },
        }),
      ];
      const { client } = makeClient({ value: events });
      const result = await client.searchMeetings({ keyword: 'WSIB' });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });

    it('matches keyword in organizer name', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          organizer: {
            emailAddress: { name: 'Charlie Brown', address: 'charlie@example.com' },
          },
          body: { content: '' },
        }),
        makeRawEvent({ id: 'e2', body: { content: '' } }),
      ];
      const { client } = makeClient({ value: events });
      const result = await client.searchMeetings({ keyword: 'Charlie' });
      expect(result.meetings).toHaveLength(1);
    });

    it('is case-insensitive', async () => {
      const { client } = makeClient({
        value: [makeRawEvent({ subject: 'PTU Deployment' })],
      });
      const result = await client.searchMeetings({ keyword: 'ptu' });
      expect(result.meetings).toHaveLength(1);
    });

    it('returns empty array when keyword has no match', async () => {
      const { client } = makeClient({
        value: [makeRawEvent({ subject: 'Team Standup', body: { content: 'nothing' } })],
      });
      const result = await client.searchMeetings({ keyword: 'WSIB' });
      expect(result.meetings).toHaveLength(0);
    });

    it('matches keyword in attendee email', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          attendees: [
            {
              emailAddress: { name: 'Dana White', address: 'dana@wsib.ca' },
              status: { response: 'accepted' },
            },
          ],
          body: { content: '' },
        }),
        makeRawEvent({ id: 'e2', body: { content: '' } }),
      ];
      const { client } = makeClient({ value: events });
      const result = await client.searchMeetings({ keyword: 'wsib.ca' });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });
  });

  // -------------------------------------------------------------------------
  // Fuzzy keyword matching
  // -------------------------------------------------------------------------

  describe('keyword search – fuzzy', () => {
    it('returns meetings when all tokens match (different words)', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          subject: 'PTU Budget Review',
          body: { content: 'Quarterly alignment' },
        }),
        makeRawEvent({
          id: 'e2',
          subject: 'Weekly Standup',
          body: { content: 'No PTU topics' },
        }),
      ];
      const { client } = makeClient({ value: events });
      // Both tokens "ptu" and "budget" must be present
      const result = await client.searchMeetings({
        keyword: 'PTU budget',
        fuzzy: true,
      });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });

    it('returns no meetings when not all tokens match', async () => {
      const { client } = makeClient({
        value: [makeRawEvent({ subject: 'PTU Review', body: { content: '' } })],
      });
      const result = await client.searchMeetings({
        keyword: 'PTU budget',
        fuzzy: true,
      });
      expect(result.meetings).toHaveLength(0);
    });

    it('handles single-token fuzzy search identically to exact', async () => {
      const { client } = makeClient({
        value: [makeRawEvent({ subject: 'PTU Review' })],
      });
      const result = await client.searchMeetings({ keyword: 'PTU', fuzzy: true });
      expect(result.meetings).toHaveLength(1);
    });
  });

  // -------------------------------------------------------------------------
  // Date range filtering
  // -------------------------------------------------------------------------

  describe('date range filtering', () => {
    it('uses calendarView endpoint when dateFrom is provided', async () => {
      const { client, apiMock } = makeClient({ value: [] });
      await client.searchMeetings({ dateFrom: '2024-02-01' });
      const path: string = apiMock.mock.calls[0][0];
      expect(path).toBe('/me/calendarView');
    });

    it('uses calendarView endpoint when dateTo is provided', async () => {
      const { client, apiMock } = makeClient({ value: [] });
      await client.searchMeetings({ dateTo: '2024-02-28' });
      const path: string = apiMock.mock.calls[0][0];
      expect(path).toBe('/me/calendarView');
    });

    it('passes correct startDateTime and endDateTime query params', async () => {
      const requestBuilder: Record<string, jest.Mock> = {};
      const queryMock = jest.fn().mockReturnValue(requestBuilder);
      const getMock = jest.fn().mockResolvedValue({ value: [] });

      for (const m of ['top', 'select', 'filter', 'search', 'orderby', 'header']) {
        requestBuilder[m] = jest.fn().mockReturnValue(requestBuilder);
      }
      requestBuilder['query'] = queryMock;
      requestBuilder['get'] = getMock;

      const graphClient = {
        api: jest.fn().mockReturnValue(requestBuilder),
      } as unknown as Client;
      const searcher = new CalendarSearchClient(graphClient, 'me');

      await searcher.searchMeetings({
        dateFrom: '2024-02-01',
        dateTo: '2024-02-29',
      });

      expect(queryMock).toHaveBeenCalledWith(
        expect.objectContaining({
          startDateTime: expect.stringContaining('2024-02-01'),
          endDateTime: expect.stringContaining('2024-02-29'),
        }),
      );
    });

    it('falls back to /me/events when no date range is given', async () => {
      const { client, apiMock } = makeClient({ value: [] });
      await client.searchMeetings({ keyword: 'WSIB' });
      const path: string = apiMock.mock.calls[0][0];
      expect(path).toBe('/me/events');
    });

    it('accepts Date objects for dateFrom/dateTo', async () => {
      const { client, apiMock } = makeClient({ value: [] });
      await client.searchMeetings({
        dateFrom: new Date('2024-01-01'),
        dateTo: new Date('2024-03-31'),
      });
      const path: string = apiMock.mock.calls[0][0];
      expect(path).toBe('/me/calendarView');
    });
  });

  // -------------------------------------------------------------------------
  // Organizer and attendee filters
  // -------------------------------------------------------------------------

  describe('organizer / attendee filtering', () => {
    it('applies organizer filter via $filter', async () => {
      const requestBuilder: Record<string, jest.Mock> = {};
      const filterMock = jest.fn().mockReturnValue(requestBuilder);
      const getMock = jest.fn().mockResolvedValue({ value: [] });

      for (const m of ['query', 'top', 'select', 'search', 'orderby', 'header']) {
        requestBuilder[m] = jest.fn().mockReturnValue(requestBuilder);
      }
      requestBuilder['filter'] = filterMock;
      requestBuilder['get'] = getMock;

      const graphClient = {
        api: jest.fn().mockReturnValue(requestBuilder),
      } as unknown as Client;
      const searcher = new CalendarSearchClient(graphClient, 'me');

      await searcher.searchMeetings({
        organizer: 'alice@example.com',
        dateFrom: '2024-01-01',
        dateTo: '2024-03-31',
      });

      expect(filterMock).toHaveBeenCalledWith(
        expect.stringContaining('alice@example.com'),
      );
    });

    it('post-filters by organizer email at the in-process level', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          organizer: {
            emailAddress: { name: 'Alice', address: 'alice@example.com' },
          },
        }),
        makeRawEvent({
          id: 'e2',
          organizer: {
            emailAddress: { name: 'Bob', address: 'bob@example.com' },
          },
        }),
      ];
      const { client } = makeClient({ value: events });
      // Using keyword with fuzzy to avoid Graph-level $search and exercise
      // the in-process path
      const result = await client.searchMeetings({
        keyword: 'alice@example.com',
      });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });

    it('post-filters by attendee email at the in-process level', async () => {
      const events = [
        makeRawEvent({
          id: 'e1',
          attendees: [
            {
              emailAddress: { name: 'Carol', address: 'carol@corp.com' },
              status: { response: 'accepted' },
            },
          ],
          body: { content: '' },
          subject: 'Planning',
        }),
        makeRawEvent({
          id: 'e2',
          attendees: [],
          body: { content: '' },
          subject: 'Planning',
        }),
      ];
      const { client } = makeClient({ value: events });
      const result = await client.searchMeetings({ keyword: 'carol@corp.com' });
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('e1');
    });
  });

  // -------------------------------------------------------------------------
  // Pagination
  // -------------------------------------------------------------------------

  describe('pagination', () => {
    it('exposes nextLink when Graph returns @odata.nextLink', async () => {
      const { client } = makeClient({
        value: [makeRawEvent()],
        '@odata.nextLink':
          'https://graph.microsoft.com/v1.0/me/events?$skip=50',
      });
      const result = await client.searchMeetings({});
      expect(result.nextLink).toBe(
        'https://graph.microsoft.com/v1.0/me/events?$skip=50',
      );
    });

    it('does NOT expose nextLink when Graph omits it', async () => {
      const { client } = makeClient({ value: [makeRawEvent()] });
      const result = await client.searchMeetings({});
      expect(result.nextLink).toBeUndefined();
    });

    it('fetches subsequent pages using the nextLink token', async () => {
      const page2Events = [
        makeRawEvent({ id: 'page2-event-1', subject: 'Second Page Meeting' }),
      ];
      const { client, apiMock } = makeClient({ value: page2Events });

      const nextLinkUrl =
        'https://graph.microsoft.com/v1.0/me/events?$skip=50&$top=50';
      const result = await client.searchMeetings({ nextLink: nextLinkUrl });

      // The client should call the Graph API directly with the nextLink URL
      expect(apiMock).toHaveBeenCalledWith(nextLinkUrl);
      expect(result.meetings).toHaveLength(1);
      expect(result.meetings[0].id).toBe('page2-event-1');
    });

    it('exposes totalCount when Graph returns @odata.count', async () => {
      const { client } = makeClient({
        value: [makeRawEvent()],
        '@odata.count': 120,
      });
      const result = await client.searchMeetings({});
      expect(result.totalCount).toBe(120);
    });

    it('returns empty meetings array for an empty page', async () => {
      const { client } = makeClient({ value: [] });
      const result = await client.searchMeetings({});
      expect(result.meetings).toHaveLength(0);
    });

    it('handles multiple pages of results', async () => {
      // Simulate two pages: first with nextLink, second without
      const page1Events = Array.from({ length: 3 }, (_, i) =>
        makeRawEvent({ id: `e${i + 1}`, subject: `Meeting ${i + 1}` }),
      );
      const page2Events = Array.from({ length: 2 }, (_, i) =>
        makeRawEvent({ id: `e${i + 4}`, subject: `Meeting ${i + 4}` }),
      );

      const requestBuilder: Record<string, jest.Mock> = {};
      for (const m of ['query', 'top', 'select', 'filter', 'search', 'orderby', 'header']) {
        requestBuilder[m] = jest.fn().mockReturnValue(requestBuilder);
      }
      // First call returns page1 with nextLink; second call returns page2
      requestBuilder['get'] = jest
        .fn()
        .mockResolvedValueOnce({
          value: page1Events,
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events?$skip=3',
        })
        .mockResolvedValueOnce({ value: page2Events });

      const graphClient = {
        api: jest.fn().mockReturnValue(requestBuilder),
      } as unknown as Client;
      const searcher = new CalendarSearchClient(graphClient, 'me');

      const result1 = await searcher.searchMeetings({ pageSize: 3 });
      expect(result1.meetings).toHaveLength(3);
      expect(result1.nextLink).toBeDefined();

      const result2 = await searcher.searchMeetings({ nextLink: result1.nextLink });
      expect(result2.meetings).toHaveLength(2);
      expect(result2.nextLink).toBeUndefined();
    });
  });

  // -------------------------------------------------------------------------
  // userId routing
  // -------------------------------------------------------------------------

  describe('userId routing', () => {
    it('routes to /me/events for default userId', async () => {
      const { client, apiMock } = makeClient({ value: [] });
      await client.searchMeetings({});
      expect(apiMock.mock.calls[0][0]).toBe('/me/events');
    });

    it('routes to /users/{id}/events for a specific userId', async () => {
      const { client, apiMock } = makeClient({ value: [] }, 'user123');
      await client.searchMeetings({});
      expect(apiMock.mock.calls[0][0]).toBe('/users/user123/events');
    });

    it('routes to /users/{id}/calendarView when date range provided', async () => {
      const { client, apiMock } = makeClient({ value: [] }, 'user456');
      await client.searchMeetings({
        dateFrom: '2024-01-01',
        dateTo: '2024-12-31',
      });
      expect(apiMock.mock.calls[0][0]).toBe('/users/user456/calendarView');
    });
  });

  // -------------------------------------------------------------------------
  // Edge cases
  // -------------------------------------------------------------------------

  describe('edge cases', () => {
    it('returns empty result set when value is undefined', async () => {
      const { client } = makeClient({});
      const result = await client.searchMeetings({});
      expect(result.meetings).toHaveLength(0);
    });

    it('handles pageSize capped at MAX (999)', async () => {
      const requestBuilder: Record<string, jest.Mock> = {};
      const topMock = jest.fn().mockReturnValue(requestBuilder);
      const getMock = jest.fn().mockResolvedValue({ value: [] });

      for (const m of ['query', 'select', 'filter', 'search', 'orderby', 'header']) {
        requestBuilder[m] = jest.fn().mockReturnValue(requestBuilder);
      }
      requestBuilder['top'] = topMock;
      requestBuilder['get'] = getMock;

      const graphClient = {
        api: jest.fn().mockReturnValue(requestBuilder),
      } as unknown as Client;
      const searcher = new CalendarSearchClient(graphClient, 'me');

      await searcher.searchMeetings({ pageSize: 5000 });
      expect(topMock).toHaveBeenCalledWith(999);
    });

    it('escapes single quotes in organizer filter', async () => {
      const requestBuilder: Record<string, jest.Mock> = {};
      const filterMock = jest.fn().mockReturnValue(requestBuilder);
      const getMock = jest.fn().mockResolvedValue({ value: [] });

      for (const m of ['query', 'top', 'select', 'search', 'orderby', 'header']) {
        requestBuilder[m] = jest.fn().mockReturnValue(requestBuilder);
      }
      requestBuilder['filter'] = filterMock;
      requestBuilder['get'] = getMock;

      const graphClient = {
        api: jest.fn().mockReturnValue(requestBuilder),
      } as unknown as Client;
      const searcher = new CalendarSearchClient(graphClient, 'me');

      await searcher.searchMeetings({
        organizer: "O'Brien",
        dateFrom: '2024-01-01',
      });

      // Single quote in name should be escaped to ''
      expect(filterMock).toHaveBeenCalledWith(expect.stringContaining("O''Brien"));
    });
  });
});
