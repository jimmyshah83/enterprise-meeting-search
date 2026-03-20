/**
 * Cross-domain search orchestrator.
 *
 * Fans out a single unified search request to all five domain modules
 * (calendar, transcripts, chats, email, files) concurrently, then
 * aggregates, deduplicates, and correlates the results into a unified
 * response grouped by meeting / calendar event.
 *
 * When no results are found it delegates to the diagnostic engine to
 * generate a structured {@link DiagnosticSummary} with probable causes and
 * actionable retry suggestions.
 */

import {
  CalendarEvent,
  ChatMessage,
  DomainName,
  DomainStatus,
  EmailResult,
  FileResult,
  SearchRequest,
  SearchResponse,
  TranscriptResult,
  UnifiedMeetingResult,
} from './types';
import { searchCalendar } from './domains/calendar';
import { searchTranscripts } from './domains/transcripts';
import { searchChats } from './domains/chats';
import { searchEmail } from './domains/email';
import { searchFiles } from './domains/files';
import { buildDiagnostics } from './diagnostics';

// ---------------------------------------------------------------------------
// Domain adapter types (allow injection for testing)
// ---------------------------------------------------------------------------

export interface DomainAdapters {
  calendar?: (req: SearchRequest) => Promise<CalendarEvent[]>;
  transcripts?: (req: SearchRequest) => Promise<TranscriptResult[]>;
  chats?: (req: SearchRequest) => Promise<ChatMessage[]>;
  email?: (req: SearchRequest) => Promise<EmailResult[]>;
  files?: (req: SearchRequest) => Promise<FileResult[]>;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/** Wraps a promise so it rejects after `ms` milliseconds. */
function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`timeout after ${ms}ms`)), ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err: unknown) => {
        clearTimeout(timer);
        reject(err);
      },
    );
  });
}

/** Builds a placeholder UnifiedMeetingResult for a given meeting ID. */
function emptyMeetingResult(meetingId: string): UnifiedMeetingResult {
  return {
    meetingId,
    transcripts: [],
    chats: [],
    emails: [],
    files: [],
  };
}

/**
 * Deduplicates an array of domain results by their `id` field.
 * The first occurrence wins.
 */
function deduplicateById<T extends { id: string }>(items: T[]): T[] {
  const seen = new Set<string>();
  return items.filter((item) => {
    if (seen.has(item.id)) return false;
    seen.add(item.id);
    return true;
  });
}

// ---------------------------------------------------------------------------
// Aggregation & correlation logic
// ---------------------------------------------------------------------------

/**
 * Merges all domain results into a map keyed by meeting ID.
 *
 * Strategy
 * --------
 * 1. Calendar events are the authoritative "anchor" for each meeting.
 * 2. Transcripts, chats, emails and files are attached to the meeting they
 *    reference via their `meetingId` field.
 * 3. Artifacts without a `meetingId` are grouped under a synthetic key so
 *    they are still surfaced in the results.
 */
function correlateResults(
  calendarEvents: CalendarEvent[],
  transcripts: TranscriptResult[],
  chats: ChatMessage[],
  emails: EmailResult[],
  files: FileResult[],
): Map<string, UnifiedMeetingResult> {
  const map = new Map<string, UnifiedMeetingResult>();

  const getOrCreate = (id: string): UnifiedMeetingResult => {
    if (!map.has(id)) {
      map.set(id, emptyMeetingResult(id));
    }
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    return map.get(id)!;
  };

  // Anchor: calendar events
  for (const event of deduplicateById(calendarEvents)) {
    const entry = getOrCreate(event.id);
    entry.calendarEvent = event;
  }

  // Attach transcripts
  for (const transcript of deduplicateById(transcripts)) {
    getOrCreate(transcript.meetingId).transcripts.push(transcript);
  }

  // Attach chats
  for (const chat of deduplicateById(chats)) {
    const key = chat.meetingId ?? `chat::${chat.threadId}`;
    getOrCreate(key).chats.push(chat);
  }

  // Attach emails
  for (const email of deduplicateById(emails)) {
    const key = email.meetingId ?? `email::${email.id}`;
    getOrCreate(key).emails.push(email);
  }

  // Attach files
  for (const file of deduplicateById(files)) {
    const key = file.meetingId ?? `file::${file.id}`;
    getOrCreate(key).files.push(file);
  }

  return map;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Executes a cross-domain search across all five Microsoft 365 data domains
 * in parallel and returns a unified, correlated result set.
 *
 * @param request   Unified search request.
 * @param adapters  Optional injectable domain adapters (primarily for testing).
 * @returns         A {@link SearchResponse} containing correlated results or
 *                  a diagnostic summary when no results are found.
 */
export async function search(
  request: SearchRequest,
  adapters: DomainAdapters = {},
): Promise<SearchResponse> {
  const start = Date.now();
  const domainTimeout = request.domainTimeoutMs ?? 5_000;
  const overallTimeout = request.overallTimeoutMs ?? 10_000;

  // Build domain promises, each individually time-bounded
  const calendarFetch = adapters.calendar
    ? (req: SearchRequest) => adapters.calendar!(req)
    : undefined;
  const transcriptsFetch = adapters.transcripts
    ? (req: SearchRequest) => adapters.transcripts!(req)
    : undefined;
  const chatsFetch = adapters.chats
    ? (req: SearchRequest) => adapters.chats!(req)
    : undefined;
  const emailFetch = adapters.email
    ? (req: SearchRequest) => adapters.email!(req)
    : undefined;
  const filesFetch = adapters.files
    ? (req: SearchRequest) => adapters.files!(req)
    : undefined;

  type DomainSettled<T> = PromiseSettledResult<T>;

  const [calendarResult, transcriptsResult, chatsResult, emailResult, filesResult] =
    await Promise.race([
      Promise.allSettled([
        withTimeout(searchCalendar(request, calendarFetch), domainTimeout),
        withTimeout(searchTranscripts(request, transcriptsFetch), domainTimeout),
        withTimeout(searchChats(request, chatsFetch), domainTimeout),
        withTimeout(searchEmail(request, emailFetch), domainTimeout),
        withTimeout(searchFiles(request, filesFetch), domainTimeout),
      ]) as Promise<
        [
          DomainSettled<CalendarEvent[]>,
          DomainSettled<TranscriptResult[]>,
          DomainSettled<ChatMessage[]>,
          DomainSettled<EmailResult[]>,
          DomainSettled<FileResult[]>,
        ]
      >,
      // Overall timeout: resolve with all-rejected shape
      new Promise<
        [
          DomainSettled<CalendarEvent[]>,
          DomainSettled<TranscriptResult[]>,
          DomainSettled<ChatMessage[]>,
          DomainSettled<EmailResult[]>,
          DomainSettled<FileResult[]>,
        ]
      >((resolve) =>
        setTimeout(() => {
          const err = new Error(`overall timeout after ${overallTimeout}ms`);
          resolve([
            { status: 'rejected', reason: err },
            { status: 'rejected', reason: err },
            { status: 'rejected', reason: err },
            { status: 'rejected', reason: err },
            { status: 'rejected', reason: err },
          ]);
        }, overallTimeout),
      ),
    ]);

  // Unpack results and build domain status map
  const extract = <T>(result: DomainSettled<T>, emptyDefault: T): [T, DomainStatus] => {
    if (result.status === 'fulfilled') {
      const arr = result.value as unknown as unknown[];
      return [result.value, { status: 'success', resultCount: arr.length }];
    }
    const reason = result.reason as Error;
    const isTimeout = reason?.message?.includes('timeout');
    return [
      emptyDefault,
      isTimeout ? { status: 'timeout' } : { status: 'error', error: reason?.message ?? 'unknown' },
    ];
  };

  const [calendarEvents, calendarStatus] = extract<CalendarEvent[]>(calendarResult, []);
  const [transcripts, transcriptsStatus] = extract<TranscriptResult[]>(transcriptsResult, []);
  const [chats, chatsStatus] = extract<ChatMessage[]>(chatsResult, []);
  const [emails, emailStatus] = extract<EmailResult[]>(emailResult, []);
  const [files, filesStatus] = extract<FileResult[]>(filesResult, []);

  const domainStatuses: Record<DomainName, DomainStatus> = {
    calendar: calendarStatus,
    transcripts: transcriptsStatus,
    chats: chatsStatus,
    email: emailStatus,
    files: filesStatus,
  };

  // Correlate all domain results into unified meeting results
  const correlatedMap = correlateResults(calendarEvents, transcripts, chats, emails, files);
  const results = Array.from(correlatedMap.values());

  const searchDurationMs = Date.now() - start;

  if (results.length === 0) {
    return {
      results: [],
      diagnostics: buildDiagnostics(request, domainStatuses),
      totalResults: 0,
      searchDurationMs,
    };
  }

  return {
    results,
    totalResults: results.length,
    searchDurationMs,
  };
}
