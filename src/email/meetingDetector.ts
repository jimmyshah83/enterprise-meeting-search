import { GraphMessage, MeetingEmailType } from './emailSearchTypes';

/**
 * Subject-line patterns that indicate a meeting invite, forward, reply, or
 * cancellation.  Matching is case-insensitive.
 */
const INVITE_SUBJECT_PATTERNS = [
  /\binvite\b/i,
  /\binvitation\b/i,
  /\bmeeting request\b/i,
  /\bschedule\b.*\bmeeting\b/i,
];

const FORWARD_SUBJECT_PATTERNS = [
  /^fw:/i,
  /^fwd:/i,
];

const REPLY_SUBJECT_PATTERNS = [
  /^re:/i,
];

const CANCELLATION_SUBJECT_PATTERNS = [
  /\bcanceled\b/i,
  /\bcancelled\b/i,
  /\bcancellation\b/i,
];

/**
 * Body / category patterns that indicate a meeting-related email even when
 * the subject does not match.
 */
const MEETING_BODY_PATTERNS = [
  /\bjoin\s+(microsoft\s+)?teams\s+meeting\b/i,
  /\bjoin\s+zoom\s+meeting\b/i,
  /\bjoin\s+google\s+meet\b/i,
  /\bmeeting\s+link\b/i,
  /\bdial-in\s+number\b/i,
  /\bconference\s+bridge\b/i,
  /\bagenda\b.*\bmeeting\b/i,
];

const MEETING_CATEGORIES = new Set([
  'meeting',
  'appointment',
  'calendar',
]);

/** Result of the meeting detection analysis. */
export interface MeetingDetectionResult {
  isMeetingRelated: boolean;
  meetingType?: MeetingEmailType;
}

/**
 * Analyses a raw Graph message and decides whether it is meeting-related and,
 * if so, what kind of meeting email it is.
 */
export function detectMeetingEmail(message: GraphMessage): MeetingDetectionResult {
  const subject = message.subject ?? '';
  const body = message.bodyPreview ?? '';
  const categories = (message.categories ?? []).map((c) => c.toLowerCase());

  // Check category tags set by Outlook.
  const hasMeetingCategory = categories.some((c) => MEETING_CATEGORIES.has(c));

  // Check subject patterns.
  const isCancellation = CANCELLATION_SUBJECT_PATTERNS.some((p) => p.test(subject));
  const isForward = FORWARD_SUBJECT_PATTERNS.some((p) => p.test(subject));
  const isReply = REPLY_SUBJECT_PATTERNS.some((p) => p.test(subject));
  const isInvite = INVITE_SUBJECT_PATTERNS.some((p) => p.test(subject));

  // Check body patterns.
  const hasMeetingBodyHint = MEETING_BODY_PATTERNS.some((p) => p.test(body));

  const isMeetingRelated =
    isCancellation ||
    isInvite ||
    hasMeetingCategory ||
    hasMeetingBodyHint ||
    (isForward && hasMeetingBodyHint) ||
    (isReply && hasMeetingBodyHint);

  if (!isMeetingRelated) {
    return { isMeetingRelated: false };
  }

  let meetingType: MeetingEmailType;

  if (isCancellation) {
    meetingType = 'cancellation';
  } else if (isInvite || hasMeetingCategory) {
    if (isForward) {
      meetingType = 'forward';
    } else if (isReply) {
      meetingType = 'reply';
    } else {
      meetingType = 'invite';
    }
  } else if (isForward) {
    meetingType = 'forward';
  } else {
    meetingType = 'reply';
  }

  return { isMeetingRelated: true, meetingType };
}
