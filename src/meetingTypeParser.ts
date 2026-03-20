import { MeetingType } from './types';

interface MeetingTypeRule {
  type: MeetingType;
  patterns: RegExp[];
}

const RULES: MeetingTypeRule[] = [
  // "channel" must be checked before "teams" so that "Teams channel" → channel
  {
    type: 'channel',
    patterns: [
      /\bchannel\s+(?:meeting|call|post)\b/i,
      /\bteams?\s+channel\b/i,
      /\bslack\s+channel\b/i,
      /\bchannel\b/i,
    ],
  },
  {
    type: 'teams',
    patterns: [
      /\bteams?\s+(?:meeting|call|chat)\b/i,
      /\bmicrosoft\s+teams?\b/i,
      /\bms\s+teams?\b/i,
      /\bteams\b/i,
    ],
  },
  {
    type: 'external',
    patterns: [
      /\bexternal\s+(?:meeting|call|attendee|participant|guest)\b/i,
      /\boutside\s+(?:the\s+)?(?:company|org(?:aniz(?:ation)?)?|firm)\b/i,
      /\bclient\s+(?:meeting|call)\b/i,
      /\bvendor\s+(?:meeting|call)\b/i,
      /\bexternal\b/i,
    ],
  },
];

/**
 * Identifies the meeting type from a natural language hint.
 *
 * Rules are evaluated in order; the first match wins. `channel` is checked
 * before `teams` so that "Teams channel meeting" resolves to `channel`.
 *
 * @param hint - Raw user-supplied text.
 * @returns The detected `MeetingType`, or `'unknown'` when no rule matches.
 */
export function parseMeetingType(hint: string): MeetingType {
  for (const rule of RULES) {
    for (const pattern of rule.patterns) {
      if (pattern.test(hint)) {
        return rule.type;
      }
    }
  }
  return 'unknown';
}
