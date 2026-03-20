/**
 * Cross-domain search orchestrator.
 *
 * Fans out a single unified search request to all five domain modules
 * (calendar, transcripts, chats, email, files) concurrently, then
 * aggregates, deduplicates, and correlates the results into a unified
 * response grouped by meeting / calendar event.
 */

import {
  CalendarEvent,
  ChatMessage,
  DiagnosticSummary,
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

export interface DomainAdapters {
  calendar?: (req: SearchRequest) => Promise<CalendarEvent[]>;
  transcripts?: (req: SearchRequest) => Promise<TranscriptResult[]>;
  chats?: (req: SearchRequest) => Promise<ChatMessage[]>;
  email?: (req: SearchRequest) => Promise<EmailResult[]>;
  files?: (req: SearchRequest) => Promise<FileResult[]>;
}

function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`timeout after ${ms}ms`)), ms);
    promise.then(
      (value) => { clearTimeout(timer); resolve(value); },
      (err: unknown) => { clearTimeout(timer); reject(err); },
    );
  });
}

function emptyMeetingResult(meetingId: string): UnifiedMeetingResult {
  return { meetingId, transcripts: [], chats: [], emails: [], files: [] };
}

function deduplicateById<T extends { id: string }>(items: T[]): T[] {
  const seen = new Set<string>();
  return items.filter((item) => {
    if (seen.has(item.id)) return false;
    seen.add(item.id);
    return true;
  });
}

function correlateResults(
  calendarEvents: CalendarEvent[],
  transcripts: TranscriptResult[],
  chats: ChatMessage[],
  emails: EmailResult[],
  files: FileResult[],
): Map<string, UnifiedMeetingResult> {
  const map = new Map<string, UnifiedMeetingResult>();
  const getOrCreate = (id: string): UnifiedMeetingResult => {
    if (!map.has(id)) map.set(id, emptyMeetingResult(id));
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    return map.get(id)!;
  };

  for (const event of deduplicateById(calendarEvents)) {
    const entry = getOrCreate(event.id);
    entry.calendarEvent = event;
  }
  for (const transcript of deduplicateById(transcripts)) {
    getOrCreate(transcript.meetingId).transcripts.push(transcript);
  }
  for (const chat of deduplicateById(chats)) {
    getOrCreate(chat.meetingId ?? `chat::${chat.threadId}`).chats.push(chat);
  }
  for (const email of deduplicateById(emails)) {
    getOrCreate(email.meetingId ?? `email::${email.id}`).emails.push(email);
  }
  for (const file of deduplicateById(files)) {
    getOrCreate(file.meetingId ?? `file::${file.id}`).files.push(file);
  }
  return map;
}

function buildDiagnosticSummary(
  request: SearchRequest,
  statuses: Record<DomainName, DomainStatus>,
): DiagnosticSummary {
  const reasons: string[] = [];
  const suggestions: string[] = [];

  const timedOutDomains = (Object.entries(statuses) as [DomainName, DomainStatus][])
    .filter(([, s]) => s.status === 'timeout').map(([name]) => name);
  const erroredDomains = (Object.entries(statuses) as [DomainName, DomainStatus][])
    .filter(([, s]) => s.status === 'error').map(([name]) => name);

  if (timedOutDomains.length > 0) {
    reasons.push(`The following domains did not respond in time: ${timedOutDomains.join(', ')}.`);
    suggestions.push('Try again later, or increase the domain timeout via the domainTimeoutMs option.');
  }
  if (erroredDomains.length > 0) {
    reasons.push(`The following domains returned errors: ${erroredDomains.join(', ')}.`);
    suggestions.push('Check service connectivity and permissions for the listed domains.');
  }
  if (request.keywords && request.keywords.length > 0) {
    reasons.push(`No content matched the keywords: "${request.keywords.join('", "')}".`);
    suggestions.push('Try broader or alternative keywords.', 'Verify that the correct date range is set.');
  }
  if (request.dateRange) {
    const { start, end } = request.dateRange;
    reasons.push(`No meetings were found in the date range ${start.toISOString()} – ${end.toISOString()}.`);
    suggestions.push('Widen the date range to capture more results.');
  }
  if (request.attendees && request.attendees.length > 0) {
    reasons.push(`No meetings were found involving attendees: ${request.attendees.join(', ')}.`);
    suggestions.push('Verify the attendee email addresses are correct.');
  }
  if (request.organizer) {
    reasons.push(`No meetings were found organized by: ${request.organizer}.`);
    suggestions.push('Verify the organizer email address is correct.');
  }
  if (reasons.length === 0) {
    reasons.push('No results were found across any domain for the given criteria.');
    suggestions.push('Try relaxing the search filters.', 'Ensure you have permissions to access the relevant data sources.');
  }

  return { message: 'No results found across all search domains.', reasons, suggestions, domainStatuses: statuses };
}

export async function search(
  request: SearchRequest,
  adapters: DomainAdapters = {},
): Promise<SearchResponse> {
  const start = Date.now();
  const domainTimeout = request.domainTimeoutMs ?? 5_000;
  const overallTimeout = request.overallTimeoutMs ?? 10_000;

  const calendarFetch = adapters.calendar ? (req: SearchRequest) => adapters.calendar!(req) : undefined;
  const transcriptsFetch = adapters.transcripts ? (req: SearchRequest) => adapters.transcripts!(req) : undefined;
  const chatsFetch = adapters.chats ? (req: SearchRequest) => adapters.chats!(req) : undefined;
  const emailFetch = adapters.email ? (req: SearchRequest) => adapters.email!(req) : undefined;
  const filesFetch = adapters.files ? (req: SearchRequest) => adapters.files!(req) : undefined;

  type DomainSettled<T> = PromiseSettledResult<T>;

  const [calendarResult, transcriptsResult, chatsResult, emailResult, filesResult] =
    await Promise.race([
      Promise.allSettled([
        withTimeout(searchCalendar(request, calendarFetch), domainTimeout),
        withTimeout(searchTranscripts(request, transcriptsFetch), domainTimeout),
        withTimeout(searchChats(request, chatsFetch), domainTimeout),
        withTimeout(searchEmail(request, emailFetch), domainTimeout),
        withTimeout(searchFiles(request, filesFetch), domainTimeout),
      ]) as Promise<[DomainSettled<CalendarEvent[]>, DomainSettled<TranscriptResult[]>, DomainSettled<ChatMessage[]>, DomainSettled<EmailResult[]>, DomainSettled<FileResult[]>]>,
      new Promise<[DomainSettled<CalendarEvent[]>, DomainSettled<TranscriptResult[]>, DomainSettled<ChatMessage[]>, DomainSettled<EmailResult[]>, DomainSettled<FileResult[]>]>(
        (resolve) => setTimeout(() => {
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

  const extract = <T>(result: DomainSettled<T>, emptyDefault: T): [T, DomainStatus] => {
    if (result.status === 'fulfilled') {
      const arr = result.value as unknown as unknown[];
      return [result.value, { status: 'success', resultCount: arr.length }];
    }
    const reason = result.reason as Error;
    const isTimeout = reason?.message?.includes('timeout');
    return [emptyDefault, isTimeout ? { status: 'timeout' } : { status: 'error', error: reason?.message ?? 'unknown' }];
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

  const correlatedMap = correlateResults(calendarEvents, transcripts, chats, emails, files);
  const results = Array.from(correlatedMap.values());
  const searchDurationMs = Date.now() - start;

  if (results.length === 0) {
    return { results: [], diagnostics: buildDiagnosticSummary(request, domainStatuses), totalResults: 0, searchDurationMs };
  }

  return { results, totalResults: results.length, searchDurationMs };
}
