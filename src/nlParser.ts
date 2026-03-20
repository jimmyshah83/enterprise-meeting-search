/**
 * Natural language parser for meeting search hints.
 *
 * Converts free-form text (e.g. "mid-January, John Smith was there") into
 * structured SearchRequest parameters.
 */

import { SearchRequest } from './types';

// ---------------------------------------------------------------------------
// Date parsing
// ---------------------------------------------------------------------------

const MONTH_MAP: Record<string, number> = {
  january: 1, jan: 1,
  february: 2, feb: 2,
  march: 3, mar: 3,
  april: 4, apr: 4,
  may: 5,
  june: 6, jun: 6,
  july: 7, jul: 7,
  august: 8, aug: 8,
  september: 9, sep: 9, sept: 9,
  october: 10, oct: 10,
  november: 11, nov: 11,
  december: 12, dec: 12,
};

function parseDateHint(hint: string, referenceDate: Date): { start: Date; end: Date } | undefined {
  const lower = hint.toLowerCase();
  const year = referenceDate.getFullYear();

  // "mid-January" / "mid January" pattern
  const midMatch = lower.match(/\bmid[\s-]?(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/);
  if (midMatch) {
    const month = MONTH_MAP[midMatch[1]];
    if (month !== undefined) {
      return {
        start: new Date(year, month - 1, 10),
        end: new Date(year, month - 1, 20),
      };
    }
  }

  // "early January" pattern
  const earlyMatch = lower.match(/\bearly\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/);
  if (earlyMatch) {
    const month = MONTH_MAP[earlyMatch[1]];
    if (month !== undefined) {
      return {
        start: new Date(year, month - 1, 1),
        end: new Date(year, month - 1, 10),
      };
    }
  }

  // "late January" pattern
  const lateMatch = lower.match(/\blate\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/);
  if (lateMatch) {
    const month = MONTH_MAP[lateMatch[1]];
    if (month !== undefined) {
      return {
        start: new Date(year, month - 1, 20),
        end: new Date(year, month),
      };
    }
  }

  // Plain month name (e.g. "January 2024" or just "January")
  const monthMatch = lower.match(/\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)(?:\s+(\d{4}))?\b/);
  if (monthMatch) {
    const month = MONTH_MAP[monthMatch[1]];
    const matchYear = monthMatch[2] ? parseInt(monthMatch[2], 10) : year;
    if (month !== undefined) {
      return {
        start: new Date(matchYear, month - 1, 1),
        end: new Date(matchYear, month, 0),
      };
    }
  }

  // "last week"
  if (/\blast\s+week\b/.test(lower)) {
    const end = new Date(referenceDate);
    end.setDate(end.getDate() - end.getDay());
    const start = new Date(end);
    start.setDate(end.getDate() - 6);
    return { start, end };
  }

  // "last month"
  if (/\blast\s+month\b/.test(lower)) {
    const start = new Date(year, referenceDate.getMonth() - 1, 1);
    const end = new Date(year, referenceDate.getMonth(), 0);
    return { start, end };
  }

  return undefined;
}

// ---------------------------------------------------------------------------
// Attendee parsing
// ---------------------------------------------------------------------------

/** Extracts proper names following attendee-signal phrases. */
function parseAttendees(hint: string): string[] {
  const attendees: string[] = [];

  // Patterns: "John Smith was there", "with Jane Doe", "attended by Bob"
  const patterns = [
    /\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+(?:was there|attended|was present|joined|participated)/g,
    /\bwith\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\b/g,
    /\battended\s+by\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/g,
    /\binvolving\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/g,
  ];

  for (const pattern of patterns) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(hint)) !== null) {
      const name = match[1].trim();
      if (!attendees.includes(name)) {
        attendees.push(name);
      }
    }
  }

  return attendees;
}

// ---------------------------------------------------------------------------
// Meeting type parsing
// ---------------------------------------------------------------------------

function parseMeetingType(hint: string): string[] {
  const lower = hint.toLowerCase();
  const types: string[] = [];

  if (/\bteams?\b/.test(lower)) types.push('Teams');
  if (/\bzoom\b/.test(lower)) types.push('Zoom');
  if (/\bwebex\b/.test(lower)) types.push('Webex');
  if (/\bin[\s-]?person\b/.test(lower)) types.push('InPerson');
  if (/\bexternal\b/.test(lower)) types.push('External');
  if (/\bchannel\b/.test(lower)) types.push('Channel');

  return types;
}

// ---------------------------------------------------------------------------
// Consumed patterns (tokens extracted by other parsers)
// ---------------------------------------------------------------------------

const CONSUMED_PATTERNS: RegExp[] = [
  /\b(?:early|mid[\s-]?|late)\s+(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/gi,
  /\b(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)(?:\s+\d{4})?\b/gi,
  /\b(?:last|this)\s+(?:week|month|year)\b/gi,
  /\bwas there\b/gi,
  /\battended(?: by)?\b/gi,
  /\b(?:joined|participated|was present)\b/gi,
  /\bwith\b/gi,
  /\b(?:involving|alongside)\b/gi,
  /\b(?:microsoft\s+)?teams?\s*(?:meeting|call|chat)?\b/gi,
  /\bzoom\b/gi,
  /\bwebex\b/gi,
  /\bin[\s-]?person\b/gi,
  /\bexternal\b/gi,
  /\bchannel\b/gi,
  /\band\b/gi,
  /\bthe\b/gi,
  /\bin\b/gi,
  /\bfor\b/gi,
  /\ba\b/gi,
  /\bit\s+was\b/gi,
  /\baround\b/gi,
];

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parses a free-form natural language hint into a structured SearchRequest.
 *
 * @param hint          The raw user-supplied text.
 * @param referenceDate The date to treat as "today". Defaults to `new Date()`.
 */
export function parseNaturalLanguage(
  hint: string,
  referenceDate: Date = new Date(),
): SearchRequest {
  const dateRange = parseDateHint(hint, referenceDate);
  const attendees = parseAttendees(hint);
  const meetingType = parseMeetingType(hint);

  // Build keyword list by stripping consumed tokens
  let remaining = hint;
  for (const pattern of CONSUMED_PATTERNS) {
    pattern.lastIndex = 0;
    remaining = remaining.replace(pattern, ' ');
  }
  // Remove proper names already captured as attendees
  for (const attendee of attendees) {
    remaining = remaining.replace(attendee, ' ');
  }

  const keywords = remaining
    .split(/\s+/)
    .map((w) => w.replace(/[^a-zA-Z0-9-]/g, '').trim())
    .filter((w) => w.length > 2);

  const result: SearchRequest = {};
  if (dateRange) result.dateRange = dateRange;
  if (attendees.length > 0) result.attendees = attendees;
  if (meetingType.length > 0) result.meetingType = meetingType;
  if (keywords.length > 0) result.keywords = keywords;

  return result;
}
