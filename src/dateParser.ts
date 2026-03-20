import { DateRange } from './types';

/** Full and abbreviated month names mapped to 0-based month index. */
const MONTH_MAP: Record<string, number> = {
  january: 0,
  jan: 0,
  february: 1,
  feb: 1,
  march: 2,
  mar: 2,
  april: 3,
  apr: 3,
  may: 4,
  june: 5,
  jun: 5,
  july: 6,
  jul: 6,
  august: 7,
  aug: 7,
  september: 8,
  sep: 8,
  sept: 8,
  october: 9,
  oct: 9,
  november: 10,
  nov: 10,
  december: 11,
  dec: 11,
};

const SEASON_MONTHS: Record<string, number> = {
  spring: 2,  // March
  summer: 5,  // June
  fall: 8,    // September
  autumn: 8,  // September
  winter: 11, // December
};

/** Number words → numeric value (for "two weeks ago", etc.) */
const NUMBER_WORDS: Record<string, number> = {
  one: 1, two: 2, three: 3, four: 4, five: 5,
  six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  eleven: 11, twelve: 12, a: 1, an: 1, couple: 2,
};

/**
 * Build a DateRange for a specific month (with optional year) covering the
 * requested portion of the month.
 */
function monthRange(
  year: number,
  month: number,
  portion: 'early' | 'mid' | 'late' | 'full',
): DateRange {
  switch (portion) {
    case 'early':
      return {
        start: new Date(year, month, 1),
        end: new Date(year, month, 10),
      };
    case 'mid':
      return {
        start: new Date(year, month, 11),
        end: new Date(year, month, 20),
      };
    case 'late':
      return {
        start: new Date(year, month, 21),
        // last day of month
        end: new Date(year, month + 1, 0),
      };
    case 'full':
      return {
        start: new Date(year, month, 1),
        end: new Date(year, month + 1, 0),
      };
  }
}

/**
 * Try to infer the most recent occurrence of a month relative to `now`.
 * If the month hasn't happened yet in the current year we go to the previous year.
 */
function resolveMonth(monthIndex: number, referenceDate: Date): number {
  if (monthIndex <= referenceDate.getMonth()) {
    return referenceDate.getFullYear();
  }
  return referenceDate.getFullYear() - 1;
}

/**
 * Resolve a season name to a DateRange covering the three-month season.
 */
function seasonRange(season: string, referenceDate: Date): DateRange {
  const startMonth = SEASON_MONTHS[season];
  const year =
    startMonth <= referenceDate.getMonth()
      ? referenceDate.getFullYear()
      : referenceDate.getFullYear() - 1;

  return {
    start: new Date(year, startMonth, 1),
    end: new Date(year, startMonth + 3, 0),
  };
}

/**
 * Parse a natural language date hint into a DateRange.
 *
 * @param hint - A natural language string fragment, e.g. "early February" or
 *               "around two weeks ago".
 * @param referenceDate - The date to treat as "today". Defaults to `new Date()`.
 * @returns A `DateRange` if the hint was understood, otherwise `undefined`.
 */
export function parseDateHint(
  hint: string,
  referenceDate: Date = new Date(),
): DateRange | undefined {
  const lower = hint.toLowerCase().trim();

  // ── "today" / "yesterday" ──────────────────────────────────────────────────
  if (/\btoday\b/.test(lower)) {
    const d = new Date(referenceDate);
    d.setHours(0, 0, 0, 0);
    const end = new Date(d);
    end.setHours(23, 59, 59, 999);
    return { start: d, end };
  }

  if (/\byesterday\b/.test(lower)) {
    const d = new Date(referenceDate);
    d.setDate(d.getDate() - 1);
    d.setHours(0, 0, 0, 0);
    const end = new Date(d);
    end.setHours(23, 59, 59, 999);
    return { start: d, end };
  }

  // ── "last week" ────────────────────────────────────────────────────────────
  if (/\blast\s+week\b/.test(lower)) {
    const d = new Date(referenceDate);
    const dayOfWeek = d.getDay(); // 0=Sun
    // Start of last week (Monday)
    const startOffset = dayOfWeek === 0 ? -13 : -(dayOfWeek + 6);
    const start = new Date(d);
    start.setDate(d.getDate() + startOffset);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return { start, end };
  }

  // ── "last month" ───────────────────────────────────────────────────────────
  if (/\blast\s+month\b/.test(lower)) {
    const year =
      referenceDate.getMonth() === 0
        ? referenceDate.getFullYear() - 1
        : referenceDate.getFullYear();
    const month =
      referenceDate.getMonth() === 0 ? 11 : referenceDate.getMonth() - 1;
    return monthRange(year, month, 'full');
  }

  // ── "last year" ────────────────────────────────────────────────────────────
  if (/\blast\s+year\b/.test(lower)) {
    const y = referenceDate.getFullYear() - 1;
    return { start: new Date(y, 0, 1), end: new Date(y, 11, 31) };
  }

  // ── "last <season>" e.g. "last fall", "last summer" ───────────────────────
  const lastSeasonMatch = lower.match(
    /\blast\s+(spring|summer|fall|autumn|winter)\b/,
  );
  if (lastSeasonMatch) {
    const season = lastSeasonMatch[1] === 'autumn' ? 'fall' : lastSeasonMatch[1];
    const startMonth = SEASON_MONTHS[season];
    // Force the previous occurrence
    const year =
      startMonth < referenceDate.getMonth()
        ? referenceDate.getFullYear()
        : referenceDate.getFullYear() - 1;
    return {
      start: new Date(year, startMonth, 1),
      end: new Date(year, startMonth + 3, 0),
    };
  }

  // ── "last <month>" e.g. "last February" ──────────────────────────────────
  const lastMonthNameMatch = lower.match(
    /\blast\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/,
  );
  if (lastMonthNameMatch) {
    const monthIndex = MONTH_MAP[lastMonthNameMatch[1]];
    // Force the previous occurrence
    const year =
      monthIndex < referenceDate.getMonth()
        ? referenceDate.getFullYear()
        : referenceDate.getFullYear() - 1;
    return monthRange(year, monthIndex, 'full');
  }

  // ── "N days/weeks/months ago" (with numeric or word counts) ───────────────
  const agoMatch = lower.match(
    /\b(?:around\s+)?(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|a|an|couple)\s+(day|days|week|weeks|month|months)\s+ago\b/,
  );
  if (agoMatch) {
    const rawCount = agoMatch[1];
    const count =
      /^\d+$/.test(rawCount) ? parseInt(rawCount, 10) : NUMBER_WORDS[rawCount] ?? 1;
    const unit = agoMatch[2].replace(/s$/, ''); // normalise plural

    const center = new Date(referenceDate);
    if (unit === 'day') {
      center.setDate(center.getDate() - count);
      const padding = 1;
      const start = new Date(center);
      start.setDate(center.getDate() - padding);
      start.setHours(0, 0, 0, 0);
      const end = new Date(center);
      end.setDate(center.getDate() + padding);
      end.setHours(23, 59, 59, 999);
      return { start, end };
    } else if (unit === 'week') {
      center.setDate(center.getDate() - count * 7);
      const start = new Date(center);
      start.setDate(center.getDate() - 3);
      start.setHours(0, 0, 0, 0);
      const end = new Date(center);
      end.setDate(center.getDate() + 3);
      end.setHours(23, 59, 59, 999);
      return { start, end };
    } else {
      // month
      const targetMonth = referenceDate.getMonth() - count;
      const year =
        referenceDate.getFullYear() + Math.floor(targetMonth / 12);
      const month = ((targetMonth % 12) + 12) % 12;
      return monthRange(year, month, 'full');
    }
  }

  // ── "this week" / "this month" / "this year" ──────────────────────────────
  if (/\bthis\s+week\b/.test(lower)) {
    const d = new Date(referenceDate);
    const dayOfWeek = d.getDay();
    const startOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const start = new Date(d);
    start.setDate(d.getDate() + startOffset);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return { start, end };
  }

  if (/\bthis\s+month\b/.test(lower)) {
    return monthRange(
      referenceDate.getFullYear(),
      referenceDate.getMonth(),
      'full',
    );
  }

  if (/\bthis\s+year\b/.test(lower)) {
    const y = referenceDate.getFullYear();
    return { start: new Date(y, 0, 1), end: new Date(y, 11, 31) };
  }

  // ── Season names (without "last") e.g. "fall" / "last fall" already handled
  const seasonMatch = lower.match(
    /\b(spring|summer|fall|autumn|winter)\b/,
  );
  if (seasonMatch) {
    return seasonRange(
      seasonMatch[1] === 'autumn' ? 'fall' : seasonMatch[1],
      referenceDate,
    );
  }

  // ── "early/mid/late <month>" or "beginning/end of <month>" ────────────────
  const portionMonthMatch = lower.match(
    /\b(early|beginning of|mid[\s-]?|middle of|late|end of)\s*(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/,
  );
  if (portionMonthMatch) {
    const rawPortion = portionMonthMatch[1].replace(/\s+of$/, '').trim();
    const monthIndex = MONTH_MAP[portionMonthMatch[2]];
    const year = resolveMonth(monthIndex, referenceDate);

    let portion: 'early' | 'mid' | 'late' | 'full';
    if (/early|beginning/.test(rawPortion)) {
      portion = 'early';
    } else if (/mid|middle/.test(rawPortion)) {
      portion = 'mid';
    } else {
      portion = 'late';
    }
    return monthRange(year, monthIndex, portion);
  }

  // ── "<month> <year>" e.g. "February 2024" ─────────────────────────────────
  const monthYearMatch = lower.match(
    /\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b/,
  );
  if (monthYearMatch) {
    const monthIndex = MONTH_MAP[monthYearMatch[1]];
    const year = parseInt(monthYearMatch[2], 10);
    return monthRange(year, monthIndex, 'full');
  }

  // ── Plain month name e.g. "February", "in March" ─────────────────────────
  const plainMonthMatch = lower.match(
    /\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/,
  );
  if (plainMonthMatch) {
    const monthIndex = MONTH_MAP[plainMonthMatch[1]];
    const year = resolveMonth(monthIndex, referenceDate);
    return monthRange(year, monthIndex, 'full');
  }

  return undefined;
}
