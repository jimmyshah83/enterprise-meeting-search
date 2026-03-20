import { SearchParams } from './types';
import { parseDateHint } from './dateParser';
import { parseAttendees } from './attendeeParser';
import { parseMeetingType } from './meetingTypeParser';

/**
 * Words/phrases that are consumed by other parsers and should not appear as
 * generic keywords.
 */
const CONSUMED_PATTERNS: RegExp[] = [
  // date/time phrases
  /\b(?:early|mid[\s-]?|late|beginning of|end of|middle of)\s+(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/gi,
  /\b(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)(?:\s+\d{4})?\b/gi,
  /\b(?:last|this)\s+(?:week|month|year|spring|summer|fall|autumn|winter)\b/gi,
  /\blast\s+(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/gi,
  /\b(?:around\s+)?(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|a|an|couple)\s+(?:day|days|week|weeks|month|months)\s+ago\b/gi,
  /\b(?:today|yesterday)\b/gi,
  /\b(?:spring|summer|fall|autumn|winter)\b/gi,
  // attendee trigger words
  /\bwas there\b/gi,
  /\battended(?: by)?\b/gi,
  /\b(?:joined|participated|was present|is present|will be there)\b/gi,
  /\bwith\b/gi,
  /\b(?:invited|include|including|involving|alongside)\b/gi,
  // meeting type words
  /\b(?:microsoft\s+)?teams?\s*(?:meeting|call|chat|channel)?\b/gi,
  /\b(?:ms\s+)?teams\b/gi,
  /\bchannel\s*(?:meeting|call|post)?\b/gi,
  /\bexternal\s*(?:meeting|call|attendee|participant|guest)?\b/gi,
  /\b(?:outside|client|vendor)\s*(?:the\s+)?(?:company|org|meeting|call)?\b/gi,
  // filler phrases
  /\bit\s+was\b/gi,
  /\baround\b/gi,
  /\band\b/gi,
  /\bthe\b/gi,
  /\bin\b/gi,
  /\bfor\b/gi,
  /\ba\b/gi,
];

/**
 * Parse a free-form natural language hint into structured `SearchParams`.
 *
 * @param hint - The raw user-supplied text.
 * @param referenceDate - The date to treat as "today". Defaults to `new Date()`.
 * @returns Structured `SearchParams` derived from the hint.
 */
export function parseNaturalLanguage(
  hint: string,
  referenceDate: Date = new Date(),
): SearchParams {
  const dateRange = parseDateHint(hint, referenceDate);
  const attendees = parseAttendees(hint);
  const meetingType = parseMeetingType(hint);

  // Build keyword list by stripping consumed tokens from the hint
  let remaining = hint;
  for (const pattern of CONSUMED_PATTERNS) {
    pattern.lastIndex = 0;
    remaining = remaining.replace(pattern, ' ');
  }

  // Also strip extracted attendee names
  for (const name of attendees) {
    remaining = remaining.replace(new RegExp(name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), ' ');
  }

  const keywords = remaining
    .split(/\s+/)
    .map((w) => w.replace(/[^a-zA-Z0-9'-]/g, '').trim())
    .filter((w) => w.length > 2);

  return {
    date_range: dateRange,
    attendees,
    keywords,
    meeting_type: meetingType,
  };
}
