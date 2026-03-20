import { parseNaturalLanguage } from './nlParser';

const REF = new Date(2024, 2, 15); // March 15 2024

describe('parseNaturalLanguage', () => {
  // ── Date + attendee combined ───────────────────────────────────────────────
  test('combined hint: "It was around mid-January and John Smith was there"', () => {
    const result = parseNaturalLanguage(
      'It was around mid-January and John Smith was there',
      REF,
    );
    expect(result.date_range).toBeDefined();
    expect(result.date_range!.start.getMonth()).toBe(0); // January
    expect(result.attendees).toContain('John Smith');
    expect(result.meeting_type).toBe('unknown');
  });

  // ── Date only ─────────────────────────────────────────────────────────────
  test('date only: "early Feb"', () => {
    const result = parseNaturalLanguage('early Feb', REF);
    expect(result.date_range).toBeDefined();
    expect(result.date_range!.start.getMonth()).toBe(1); // Feb
    expect(result.attendees).toHaveLength(0);
  });

  // ── Attendee only ─────────────────────────────────────────────────────────
  test('attendee only: "with Sarah Connor"', () => {
    const result = parseNaturalLanguage('with Sarah Connor', REF);
    expect(result.attendees).toContain('Sarah Connor');
    expect(result.date_range).toBeUndefined();
  });

  // ── Meeting type only ─────────────────────────────────────────────────────
  test('meeting type: "Teams meeting"', () => {
    const result = parseNaturalLanguage('Teams meeting', REF);
    expect(result.meeting_type).toBe('teams');
  });

  test('meeting type: "external"', () => {
    const result = parseNaturalLanguage('external', REF);
    expect(result.meeting_type).toBe('external');
  });

  test('meeting type: "channel meeting"', () => {
    const result = parseNaturalLanguage('channel meeting', REF);
    expect(result.meeting_type).toBe('channel');
  });

  // ── Multi-hint combinations ────────────────────────────────────────────────
  test('full combined hint: date + attendee + meeting type', () => {
    const result = parseNaturalLanguage(
      'last fall Teams meeting with Alice Johnson',
      REF,
    );
    expect(result.date_range).toBeDefined();
    expect(result.attendees).toContain('Alice Johnson');
    expect(result.meeting_type).toBe('teams');
  });

  test('multiple attendees in one hint', () => {
    const result = parseNaturalLanguage(
      'John Smith and Jane Doe attended the meeting in February',
      REF,
    );
    expect(result.attendees).toContain('John Smith');
    expect(result.attendees).toContain('Jane Doe');
    expect(result.date_range).toBeDefined();
  });

  // ── SearchParams structure ─────────────────────────────────────────────────
  test('always returns all required fields', () => {
    const result = parseNaturalLanguage('some random text', REF);
    expect(result).toHaveProperty('attendees');
    expect(result).toHaveProperty('keywords');
    expect(result).toHaveProperty('meeting_type');
    expect(Array.isArray(result.attendees)).toBe(true);
    expect(Array.isArray(result.keywords)).toBe(true);
  });

  // ── Keywords ──────────────────────────────────────────────────────────────
  test('residual keywords are extracted', () => {
    const result = parseNaturalLanguage('budget review meeting last month', REF);
    expect(result.keywords.some((k: string) => /budget|review/i.test(k))).toBe(true);
  });

  // ── Ambiguous / edge cases ─────────────────────────────────────────────────
  test('empty string returns empty/unknown results', () => {
    const result = parseNaturalLanguage('', REF);
    expect(result.date_range).toBeUndefined();
    expect(result.attendees).toHaveLength(0);
    expect(result.meeting_type).toBe('unknown');
  });

  test('"around two weeks ago" resolves a date range', () => {
    const result = parseNaturalLanguage('around two weeks ago', REF);
    expect(result.date_range).toBeDefined();
    expect(result.date_range!.start < result.date_range!.end).toBe(true);
  });
});
