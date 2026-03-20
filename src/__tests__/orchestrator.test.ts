/**
 * Unit tests for the cross-domain search orchestrator.
 *
 * Covers:
 *  - Parallel fan-out (all five domain adapters are called)
 *  - Result aggregation and grouping by meeting ID
 *  - Cross-domain correlation (transcript → meeting, chat → meeting, etc.)
 *  - Deduplication of identical IDs within a domain
 *  - Empty-result diagnostics (keyword, date-range, attendee, organizer, timeout, error)
 *  - Per-domain and overall timeout bounds
 */

import { search } from '../orchestrator';
import {
  CalendarEvent,
  ChatMessage,
  EmailResult,
  FileResult,
  SearchRequest,
  TranscriptResult,
} from '../types';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const makeCalendarEvent = (id: string, overrides: Partial<CalendarEvent> = {}): CalendarEvent => ({
  id,
  title: `Meeting ${id}`,
  startTime: new Date('2024-01-15T09:00:00Z'),
  endTime: new Date('2024-01-15T10:00:00Z'),
  organizer: 'organizer@example.com',
  attendees: ['alice@example.com', 'bob@example.com'],
  meetingType: 'Teams',
  source: 'calendar',
  ...overrides,
});

const makeTranscript = (
  id: string,
  meetingId: string,
  overrides: Partial<TranscriptResult> = {},
): TranscriptResult => ({
  id,
  meetingId,
  content: `Transcript for meeting ${meetingId}`,
  speakers: ['Alice', 'Bob'],
  timestamp: new Date('2024-01-15T10:01:00Z'),
  source: 'transcript',
  ...overrides,
});

const makeChat = (
  id: string,
  meetingId: string | undefined,
  overrides: Partial<ChatMessage> = {},
): ChatMessage => ({
  id,
  meetingId,
  threadId: `thread-${id}`,
  content: `Chat message ${id}`,
  sender: 'alice@example.com',
  timestamp: new Date('2024-01-15T09:30:00Z'),
  source: 'chat',
  ...overrides,
});

const makeEmail = (
  id: string,
  meetingId: string | undefined,
  overrides: Partial<EmailResult> = {},
): EmailResult => ({
  id,
  subject: `Email ${id}`,
  from: 'organizer@example.com',
  to: ['alice@example.com'],
  body: `Email body ${id}`,
  timestamp: new Date('2024-01-14T16:00:00Z'),
  meetingId,
  source: 'email',
  ...overrides,
});

const makeFile = (
  id: string,
  meetingId: string | undefined,
  overrides: Partial<FileResult> = {},
): FileResult => ({
  id,
  name: `file-${id}.docx`,
  url: `https://sharepoint.example.com/files/${id}`,
  meetingId,
  modifiedBy: 'alice@example.com',
  modifiedAt: new Date('2024-01-15T10:30:00Z'),
  source: 'file',
  ...overrides,
});

const baseRequest: SearchRequest = {
  keywords: ['Q4', 'budget'],
  dateRange: {
    start: new Date('2024-01-01T00:00:00Z'),
    end: new Date('2024-01-31T23:59:59Z'),
  },
};

// ---------------------------------------------------------------------------
// Fan-out – all adapters are called
// ---------------------------------------------------------------------------

describe('parallel fan-out', () => {
  it('calls all five domain adapters for every search', async () => {
    const calendarSpy = jest.fn().mockResolvedValue([]);
    const transcriptsSpy = jest.fn().mockResolvedValue([]);
    const chatsSpy = jest.fn().mockResolvedValue([]);
    const emailSpy = jest.fn().mockResolvedValue([]);
    const filesSpy = jest.fn().mockResolvedValue([]);

    await search(baseRequest, {
      calendar: calendarSpy,
      transcripts: transcriptsSpy,
      chats: chatsSpy,
      email: emailSpy,
      files: filesSpy,
    });

    expect(calendarSpy).toHaveBeenCalledTimes(1);
    expect(calendarSpy).toHaveBeenCalledWith(baseRequest);
    expect(transcriptsSpy).toHaveBeenCalledTimes(1);
    expect(chatsSpy).toHaveBeenCalledTimes(1);
    expect(emailSpy).toHaveBeenCalledTimes(1);
    expect(filesSpy).toHaveBeenCalledTimes(1);
  });
});

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

describe('result aggregation', () => {
  it('groups calendar events as the anchor for each meeting', async () => {
    const event1 = makeCalendarEvent('mtg-1');
    const event2 = makeCalendarEvent('mtg-2');

    const response = await search(baseRequest, {
      calendar: async () => [event1, event2],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.results).toHaveLength(2);
    expect(response.totalResults).toBe(2);
    const ids = response.results.map((r) => r.meetingId).sort();
    expect(ids).toEqual(['mtg-1', 'mtg-2']);
  });

  it('attaches each artifact type to its corresponding meeting bucket', async () => {
    const event = makeCalendarEvent('mtg-1');
    const transcript = makeTranscript('tr-1', 'mtg-1');
    const chat = makeChat('ch-1', 'mtg-1');
    const email = makeEmail('em-1', 'mtg-1');
    const file = makeFile('fi-1', 'mtg-1');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [transcript],
      chats: async () => [chat],
      email: async () => [email],
      files: async () => [file],
    });

    expect(response.results).toHaveLength(1);
    const result = response.results[0];
    expect(result.meetingId).toBe('mtg-1');
    expect(result.calendarEvent).toEqual(event);
    expect(result.transcripts).toEqual([transcript]);
    expect(result.chats).toEqual([chat]);
    expect(result.emails).toEqual([email]);
    expect(result.files).toEqual([file]);
  });

  it('creates synthetic buckets for orphaned artifacts without a meetingId', async () => {
    const orphanChat = makeChat('ch-orphan', undefined);

    const response = await search(baseRequest, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [orphanChat],
      email: async () => [],
      files: async () => [],
    });

    expect(response.results).toHaveLength(1);
    const result = response.results[0];
    expect(result.meetingId).toBe('chat::thread-ch-orphan');
    expect(result.chats).toContainEqual(orphanChat);
    expect(result.calendarEvent).toBeUndefined();
  });

  it('merges artifacts from multiple domains under the same meeting bucket', async () => {
    const event = makeCalendarEvent('mtg-10');
    const transcript = makeTranscript('tr-10', 'mtg-10');
    const chat1 = makeChat('ch-10a', 'mtg-10');
    const chat2 = makeChat('ch-10b', 'mtg-10');
    const email = makeEmail('em-10', 'mtg-10');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [transcript],
      chats: async () => [chat1, chat2],
      email: async () => [email],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-10');
    expect(result).toBeDefined();
    expect(result!.transcripts).toHaveLength(1);
    expect(result!.chats).toHaveLength(2);
    expect(result!.emails).toHaveLength(1);
    expect(result!.files).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Cross-domain correlation
// ---------------------------------------------------------------------------

describe('cross-domain correlation', () => {
  it('links transcript to its calendar event via meetingId', async () => {
    const event = makeCalendarEvent('mtg-A');
    const transcript = makeTranscript('tr-A', 'mtg-A');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [transcript],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-A');
    expect(result?.calendarEvent).toEqual(event);
    expect(result?.transcripts).toContainEqual(transcript);
  });

  it('links chat messages to their calendar event when meetingId matches', async () => {
    const event = makeCalendarEvent('mtg-B');
    const chat = makeChat('ch-B', 'mtg-B');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [],
      chats: async () => [chat],
      email: async () => [],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-B');
    expect(result?.chats).toContainEqual(chat);
    expect(result?.calendarEvent).toEqual(event);
  });

  it('links an email invite to its calendar event when meetingId matches', async () => {
    const event = makeCalendarEvent('mtg-C');
    const email = makeEmail('em-C', 'mtg-C');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [email],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-C');
    expect(result?.emails).toContainEqual(email);
    expect(result?.calendarEvent).toEqual(event);
  });

  it('links a shared file to its calendar event when meetingId matches', async () => {
    const event = makeCalendarEvent('mtg-D');
    const file = makeFile('fi-D', 'mtg-D');

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [file],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-D');
    expect(result?.files).toContainEqual(file);
    expect(result?.calendarEvent).toEqual(event);
  });

  it('correlates all five artifact types to the same calendar event', async () => {
    const id = 'mtg-FULL';
    const event = makeCalendarEvent(id);
    const transcript = makeTranscript('tr-F', id);
    const chat = makeChat('ch-F', id);
    const email = makeEmail('em-F', id);
    const file = makeFile('fi-F', id);

    const response = await search(baseRequest, {
      calendar: async () => [event],
      transcripts: async () => [transcript],
      chats: async () => [chat],
      email: async () => [email],
      files: async () => [file],
    });

    expect(response.results).toHaveLength(1);
    const result = response.results[0];
    expect(result.calendarEvent).toEqual(event);
    expect(result.transcripts[0]).toEqual(transcript);
    expect(result.chats[0]).toEqual(chat);
    expect(result.emails[0]).toEqual(email);
    expect(result.files[0]).toEqual(file);
  });
});

// ---------------------------------------------------------------------------
// Deduplication
// ---------------------------------------------------------------------------

describe('deduplication', () => {
  it('removes duplicate calendar events with the same id', async () => {
    const event = makeCalendarEvent('mtg-dup');

    const response = await search(baseRequest, {
      calendar: async () => [event, { ...event }],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const results = response.results.filter((r) => r.meetingId === 'mtg-dup');
    expect(results).toHaveLength(1);
  });

  it('removes duplicate transcripts with the same id', async () => {
    const transcript = makeTranscript('tr-dup', 'mtg-1');

    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('mtg-1')],
      transcripts: async () => [transcript, { ...transcript }],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-1');
    expect(result?.transcripts).toHaveLength(1);
  });

  it('removes duplicate chat messages with the same id', async () => {
    const chat = makeChat('ch-dup', 'mtg-1');

    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('mtg-1')],
      transcripts: async () => [],
      chats: async () => [chat, { ...chat }],
      email: async () => [],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-1');
    expect(result?.chats).toHaveLength(1);
  });

  it('removes duplicate emails with the same id', async () => {
    const email = makeEmail('em-dup', 'mtg-1');

    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('mtg-1')],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [email, { ...email }],
      files: async () => [],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-1');
    expect(result?.emails).toHaveLength(1);
  });

  it('removes duplicate files with the same id', async () => {
    const file = makeFile('fi-dup', 'mtg-1');

    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('mtg-1')],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [file, { ...file }],
    });

    const result = response.results.find((r) => r.meetingId === 'mtg-1');
    expect(result?.files).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Empty-result diagnostics
// ---------------------------------------------------------------------------

describe('empty-result diagnostics', () => {
  it('returns a diagnostic summary when all domains yield no results', async () => {
    const response = await search(baseRequest, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.results).toHaveLength(0);
    expect(response.totalResults).toBe(0);
    expect(response.diagnostics).toBeDefined();
    expect(response.diagnostics!.message).toContain('No results found');
  });

  it('includes keyword-related reason in diagnostics when keywords are provided', async () => {
    const response = await search({ keywords: ['synergy', 'roadmap'] }, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const reasons = response.diagnostics!.reasons.join(' ');
    expect(reasons).toContain('synergy');
    expect(reasons).toContain('roadmap');
  });

  it('includes date-range reason in diagnostics when dateRange is provided', async () => {
    const response = await search(
      {
        dateRange: {
          start: new Date('2024-03-01T00:00:00Z'),
          end: new Date('2024-03-31T23:59:59Z'),
        },
      },
      {
        calendar: async () => [],
        transcripts: async () => [],
        chats: async () => [],
        email: async () => [],
        files: async () => [],
      },
    );

    const reasons = response.diagnostics!.reasons.join(' ');
    expect(reasons).toContain('2024-03-01');
    expect(reasons).toContain('2024-03-31');
  });

  it('includes attendee reason in diagnostics when attendees are provided', async () => {
    const response = await search({ attendees: ['carol@example.com'] }, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const reasons = response.diagnostics!.reasons.join(' ');
    expect(reasons).toContain('carol@example.com');
  });

  it('includes organizer reason in diagnostics when organizer is provided', async () => {
    const response = await search({ organizer: 'dave@example.com' }, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const reasons = response.diagnostics!.reasons.join(' ');
    expect(reasons).toContain('dave@example.com');
  });

  it('does not include diagnostics when results are found', async () => {
    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('mtg-exists')],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.diagnostics).toBeUndefined();
    expect(response.results).toHaveLength(1);
  });

  it('surfaces domain status in diagnostics', async () => {
    const response = await search(baseRequest, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    const statuses = response.diagnostics!.domainStatuses;
    expect(statuses.calendar).toMatchObject({ status: 'success', resultCount: 0 });
    expect(statuses.transcripts).toMatchObject({ status: 'success', resultCount: 0 });
    expect(statuses.chats).toMatchObject({ status: 'success', resultCount: 0 });
    expect(statuses.email).toMatchObject({ status: 'success', resultCount: 0 });
    expect(statuses.files).toMatchObject({ status: 'success', resultCount: 0 });
  });
});

// ---------------------------------------------------------------------------
// Timeout handling
// ---------------------------------------------------------------------------

describe('timeout handling', () => {
  jest.setTimeout(15_000);

  it('handles a slow domain gracefully and marks it as timeout in diagnostics', async () => {
    const slow = () =>
      new Promise<CalendarEvent[]>((resolve) => setTimeout(() => resolve([]), 10_000));

    const response = await search(
      { ...baseRequest, domainTimeoutMs: 100 },
      {
        calendar: slow,
        transcripts: async () => [],
        chats: async () => [],
        email: async () => [],
        files: async () => [],
      },
    );

    // The slow domain should have timed out
    expect(response.diagnostics).toBeDefined();
    const calStatus = response.diagnostics!.domainStatuses.calendar;
    expect(calStatus.status).toBe('timeout');
  });

  it('respects the overall timeout and marks all domains as timed-out', async () => {
    const slow = () =>
      new Promise<never>((_, reject) =>
        setTimeout(() => reject(new Error('should not resolve')), 10_000),
      );

    const response = await search(
      { ...baseRequest, domainTimeoutMs: 500, overallTimeoutMs: 200 },
      {
        calendar: slow as unknown as () => Promise<CalendarEvent[]>,
        transcripts: slow as unknown as () => Promise<TranscriptResult[]>,
        chats: slow as unknown as () => Promise<ChatMessage[]>,
        email: slow as unknown as () => Promise<EmailResult[]>,
        files: slow as unknown as () => Promise<FileResult[]>,
      },
    );

    expect(response.diagnostics).toBeDefined();
    // All domains should be either timeout or error
    for (const status of Object.values(response.diagnostics!.domainStatuses)) {
      expect(['timeout', 'error']).toContain(status.status);
    }
  });

  it('marks an erroring domain as error in diagnostics', async () => {
    const response = await search(baseRequest, {
      calendar: async () => {
        throw new Error('Graph API unavailable');
      },
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.diagnostics).toBeDefined();
    const calStatus = response.diagnostics!.domainStatuses.calendar;
    expect(calStatus.status).toBe('error');
    if (calStatus.status === 'error') {
      expect(calStatus.error).toContain('Graph API unavailable');
    }
  });

  it('includes timeout domain names in diagnostic reasons', async () => {
    const slow = () =>
      new Promise<CalendarEvent[]>((resolve) => setTimeout(() => resolve([]), 10_000));

    const response = await search(
      { ...baseRequest, domainTimeoutMs: 100 },
      {
        calendar: slow,
        transcripts: async () => [],
        chats: async () => [],
        email: async () => [],
        files: async () => [],
      },
    );

    const reasons = response.diagnostics!.reasons.join(' ');
    expect(reasons).toContain('calendar');
  });
});

// ---------------------------------------------------------------------------
// Response shape
// ---------------------------------------------------------------------------

describe('response shape', () => {
  it('returns a non-negative searchDurationMs', async () => {
    const response = await search(baseRequest, {
      calendar: async () => [],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.searchDurationMs).toBeGreaterThanOrEqual(0);
  });

  it('totalResults matches the length of the results array', async () => {
    const response = await search(baseRequest, {
      calendar: async () => [makeCalendarEvent('m1'), makeCalendarEvent('m2')],
      transcripts: async () => [],
      chats: async () => [],
      email: async () => [],
      files: async () => [],
    });

    expect(response.totalResults).toBe(response.results.length);
    expect(response.totalResults).toBe(2);
  });
});
