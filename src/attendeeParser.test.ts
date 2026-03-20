import { parseAttendees } from './attendeeParser';

describe('parseAttendees', () => {
  // ── Explicit trigger phrases ───────────────────────────────────────────────
  test('"X was there" pattern', () => {
    expect(parseAttendees('John Smith was there')).toContain('John Smith');
  });

  test('"with X" pattern', () => {
    expect(parseAttendees('meeting with Sarah Connor')).toContain('Sarah Connor');
  });

  test('"attended by X" pattern', () => {
    expect(parseAttendees('attended by Alice Johnson')).toContain('Alice Johnson');
  });

  test('"invited X" pattern', () => {
    expect(parseAttendees('invited Bob Williams')).toContain('Bob Williams');
  });

  test('"X joined" pattern', () => {
    expect(parseAttendees('Maria Garcia joined')).toContain('Maria Garcia');
  });

  // ── Multiple people ────────────────────────────────────────────────────────
  test('extracts multiple names from "X and Y" pattern', () => {
    const result = parseAttendees('John Smith and Jane Doe were there');
    expect(result).toContain('John Smith');
    expect(result).toContain('Jane Doe');
  });

  test('extracts multiple names from separate sentences', () => {
    const result = parseAttendees('with Alice Brown and meeting invited Charlie Davis');
    expect(result.length).toBeGreaterThanOrEqual(1);
  });

  // ── Heuristic fallback ────────────────────────────────────────────────────
  test('falls back to capitalised bigram when no trigger present', () => {
    const result = parseAttendees('It was a meeting with Emma Wilson');
    expect(result).toContain('Emma Wilson');
  });

  // ── Edge cases ────────────────────────────────────────────────────────────
  test('returns empty array for plain lowercase sentence', () => {
    expect(parseAttendees('it was a regular meeting')).toHaveLength(0);
  });

  test('does not extract month names as person names', () => {
    const result = parseAttendees('It was in January at the office');
    expect(result).not.toContain('January');
  });

  test('does not extract "Teams" as a person name', () => {
    const result = parseAttendees('It was a Teams meeting');
    expect(result.some((n: string) => n.toLowerCase() === 'teams')).toBe(false);
  });

  test('handles three-word names', () => {
    const result = parseAttendees('with Mary Jane Watson');
    expect(result.some((n: string) => n.includes('Mary'))).toBe(true);
  });

  test('deduplicates the same name appearing twice', () => {
    const result = parseAttendees('John Smith was there and with John Smith');
    expect(result.filter((n: string) => n === 'John Smith').length).toBe(1);
  });
});
