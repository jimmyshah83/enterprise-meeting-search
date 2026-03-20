/**
 * Zod validation schemas for API request bodies.
 */

import { z } from 'zod';

export const DateRangeSchema = z.object({
  start: z.string().datetime({ message: 'start must be a valid ISO 8601 datetime string' }),
  end: z.string().datetime({ message: 'end must be a valid ISO 8601 datetime string' }),
}).refine((v) => new Date(v.start) <= new Date(v.end), {
  message: 'dateRange.start must be before or equal to dateRange.end',
});

export const SearchRequestSchema = z.object({
  keywords: z.array(z.string().min(1)).optional(),
  dateRange: DateRangeSchema.optional(),
  attendees: z.array(z.string().email()).optional(),
  organizer: z.string().email().optional(),
  meetingType: z.array(z.string().min(1)).optional(),
  domainTimeoutMs: z.number().int().positive().max(30_000).optional(),
  overallTimeoutMs: z.number().int().positive().max(60_000).optional(),
}).refine(
  (v) => (v.keywords?.length ?? 0) > 0 || v.dateRange || (v.attendees?.length ?? 0) > 0 || v.organizer || (v.meetingType?.length ?? 0) > 0,
  { message: 'At least one search criterion (keywords, dateRange, attendees, organizer, or meetingType) must be provided' },
);

export const NaturalSearchRequestSchema = z.object({
  hint: z.string().min(3, 'hint must be at least 3 characters').max(500, 'hint must be at most 500 characters'),
});

export type SearchRequestInput = z.infer<typeof SearchRequestSchema>;
export type NaturalSearchRequestInput = z.infer<typeof NaturalSearchRequestSchema>;
