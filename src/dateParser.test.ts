import { parseDateHint } from './dateParser';

/** Helper – fixed reference date: 2024-03-15 (Friday) */
const REF = new Date(2024, 2, 15); // March 15 2024

function ymd(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

describe('parseDateHint', () => {
  // ── Month portion phrases ──────────────────────────────────────────────────
  test('early Feb', () => {
    const r = parseDateHint('early Feb', REF)!;
    expect(ymd(r.start)).toBe('2024-02-01');
    expect(ymd(r.end)).toBe('2024-02-10');
  });

  test('mid-January', () => {
    const r = parseDateHint('mid-January', REF)!;
    expect(ymd(r.start)).toBe('2024-01-11');
    expect(ymd(r.end)).toBe('2024-01-20');
  });

  test('mid January (space)', () => {
    const r = parseDateHint('mid January', REF)!;
    expect(ymd(r.start)).toBe('2024-01-11');
    expect(ymd(r.end)).toBe('2024-01-20');
  });

  test('late March', () => {
    const r = parseDateHint('late March', REF)!;
    // ref is mid-March 2024; March is current month so year stays 2024
    expect(ymd(r.start)).toBe('2024-03-21');
    expect(ymd(r.end)).toBe('2024-03-31');
  });

  test('beginning of April → resolves to previous year (April is in the future)', () => {
    const r = parseDateHint('beginning of April', REF)!;
    // April > March (current month), so resolves to 2023
    expect(ymd(r.start)).toBe('2023-04-01');
    expect(ymd(r.end)).toBe('2023-04-10');
  });

  test('end of December', () => {
    const r = parseDateHint('end of December', REF)!;
    expect(ymd(r.start)).toBe('2023-12-21');
    expect(ymd(r.end)).toBe('2023-12-31');
  });

  // ── Relative "ago" expressions ─────────────────────────────────────────────
  test('two weeks ago', () => {
    const r = parseDateHint('around two weeks ago', REF)!;
    // center: March 1; padding ±3 days → Feb 27 – Mar 4
    expect(r.start).toBeDefined();
    expect(r.end).toBeDefined();
    expect(r.start < r.end).toBe(true);
    // Center should be ~14 days before REF
    const centerMs = (r.start.getTime() + r.end.getTime()) / 2;
    const centerDate = new Date(centerMs);
    const diffDays = Math.round(
      (REF.getTime() - centerDate.getTime()) / (1000 * 60 * 60 * 24),
    );
    expect(diffDays).toBeGreaterThanOrEqual(12);
    expect(diffDays).toBeLessThanOrEqual(16);
  });

  test('3 days ago', () => {
    const r = parseDateHint('3 days ago', REF)!;
    expect(r).toBeDefined();
    // center = March 12 → start March 11, end March 13
    expect(ymd(r.start)).toBe('2024-03-11');
    expect(ymd(r.end)).toBe('2024-03-13');
  });

  test('one month ago', () => {
    const r = parseDateHint('one month ago', REF)!;
    // Feb 2024
    expect(ymd(r.start)).toBe('2024-02-01');
    expect(ymd(r.end)).toBe('2024-02-29'); // 2024 is leap year
  });

  // ── Last week / month / year ───────────────────────────────────────────────
  test('last week', () => {
    const r = parseDateHint('last week', REF)!;
    // REF = Friday Mar 15 → last week Mon Mar 4 – Sun Mar 10
    expect(ymd(r.start)).toBe('2024-03-04');
    expect(ymd(r.end)).toBe('2024-03-10');
  });

  test('last month', () => {
    const r = parseDateHint('last month', REF)!;
    expect(ymd(r.start)).toBe('2024-02-01');
    expect(ymd(r.end)).toBe('2024-02-29');
  });

  test('last year', () => {
    const r = parseDateHint('last year', REF)!;
    expect(ymd(r.start)).toBe('2023-01-01');
    expect(ymd(r.end)).toBe('2023-12-31');
  });

  // ── Seasons ────────────────────────────────────────────────────────────────
  test('last fall', () => {
    const r = parseDateHint('last fall', REF)!;
    // REF is March 2024; fall starts Sep (month 8). Sep > March so it goes
    // back one more year → Sep 2023
    expect(ymd(r.start)).toBe('2023-09-01');
  });

  test('last summer', () => {
    const r = parseDateHint('last summer', REF)!;
    // Summer = June (month 5). June > March → previous year → Jun 2023
    expect(ymd(r.start)).toBe('2023-06-01');
  });

  test('plain season "winter"', () => {
    const r = parseDateHint('winter', REF)!;
    // Dec 2023 (previous occurrence)
    expect(ymd(r.start)).toBe('2023-12-01');
  });

  // ── Plain month names ──────────────────────────────────────────────────────
  test('plain month name "February"', () => {
    const r = parseDateHint('February', REF)!;
    expect(ymd(r.start)).toBe('2024-02-01');
    expect(ymd(r.end)).toBe('2024-02-29');
  });

  test('month + year "February 2023"', () => {
    const r = parseDateHint('February 2023', REF)!;
    expect(ymd(r.start)).toBe('2023-02-01');
    expect(ymd(r.end)).toBe('2023-02-28');
  });

  // ── Today / yesterday ─────────────────────────────────────────────────────
  test('today', () => {
    const r = parseDateHint('today', REF)!;
    expect(ymd(r.start)).toBe('2024-03-15');
    expect(ymd(r.end)).toBe('2024-03-15');
  });

  test('yesterday', () => {
    const r = parseDateHint('yesterday', REF)!;
    expect(ymd(r.start)).toBe('2024-03-14');
    expect(ymd(r.end)).toBe('2024-03-14');
  });

  // ── This week / month / year ───────────────────────────────────────────────
  test('this month', () => {
    const r = parseDateHint('this month', REF)!;
    expect(ymd(r.start)).toBe('2024-03-01');
    expect(ymd(r.end)).toBe('2024-03-31');
  });

  test('this year', () => {
    const r = parseDateHint('this year', REF)!;
    expect(ymd(r.start)).toBe('2024-01-01');
    expect(ymd(r.end)).toBe('2024-12-31');
  });

  // ── Edge: unrecognised input ───────────────────────────────────────────────
  test('returns undefined for unrecognised input', () => {
    expect(parseDateHint('some random text', REF)).toBeUndefined();
  });

  // ── last <month name> ──────────────────────────────────────────────────────
  test('last February', () => {
    const r = parseDateHint('last February', REF)!;
    expect(ymd(r.start)).toBe('2024-02-01');
    expect(ymd(r.end)).toBe('2024-02-29');
  });
});
