/**
 * Unit tests for the diagnostic and suggestion engine.
 *
 * Covers:
 *  - Structured output shape (probableCauses, retrySuggestions, domainStatuses)
 *  - Probable causes ranked by likelihood
 *  - Retry suggestions with informationNeeded references
 *  - Various empty-result scenarios: keyword-only, date-range, attendees,
 *    organizer, combined, no filters, domain timeout, domain error
 *  - Human-readable text formatting via formatDiagnosticText
 *  - Backwards-compatible flat reasons/suggestions arrays
 */

import { buildDiagnostics, formatDiagnosticText } from '../diagnostics';
import {
  DiagnosticSummary,
  DomainName,
  DomainStatus,
  SearchRequest,
} from '../types';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const ALL_SUCCESS: Record<DomainName, DomainStatus> = {
  calendar: { status: 'success', resultCount: 0 },
  transcripts: { status: 'success', resultCount: 0 },
  chats: { status: 'success', resultCount: 0 },
  email: { status: 'success', resultCount: 0 },
  files: { status: 'success', resultCount: 0 },
};

function buildDiag(request: SearchRequest, overrides?: Partial<Record<DomainName, DomainStatus>>): DiagnosticSummary {
  const statuses = { ...ALL_SUCCESS, ...overrides };
  return buildDiagnostics(request, statuses);
}

// ---------------------------------------------------------------------------
// Output structure
// ---------------------------------------------------------------------------

describe('output structure', () => {
  it('returns a DiagnosticSummary with all required fields', () => {
    const diag = buildDiag({ keywords: ['Q4'] });

    expect(typeof diag.message).toBe('string');
    expect(diag.message.length).toBeGreaterThan(0);
    expect(Array.isArray(diag.probableCauses)).toBe(true);
    expect(Array.isArray(diag.retrySuggestions)).toBe(true);
    expect(typeof diag.domainStatuses).toBe('object');
    expect(Array.isArray(diag.reasons)).toBe(true);
    expect(Array.isArray(diag.suggestions)).toBe(true);
  });

  it('message contains "No results found"', () => {
    const diag = buildDiag({});
    expect(diag.message).toContain('No results found');
  });

  it('probableCauses and retrySuggestions are non-empty', () => {
    const diag = buildDiag({ keywords: ['meeting'] });
    expect(diag.probableCauses.length).toBeGreaterThan(0);
    expect(diag.retrySuggestions.length).toBeGreaterThan(0);
  });

  it('each probable cause has code, description, and likelihood', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    for (const cause of diag.probableCauses) {
      expect(typeof cause.code).toBe('string');
      expect(typeof cause.description).toBe('string');
      expect(['high', 'medium', 'low']).toContain(cause.likelihood);
    }
  });

  it('each retry suggestion has code, action, and likelihood', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    for (const sug of diag.retrySuggestions) {
      expect(typeof sug.code).toBe('string');
      expect(typeof sug.action).toBe('string');
      expect(['high', 'medium', 'low']).toContain(sug.likelihood);
    }
  });

  it('domainStatuses includes all five domains', () => {
    const diag = buildDiag({});
    const domains: DomainName[] = ['calendar', 'transcripts', 'chats', 'email', 'files'];
    for (const domain of domains) {
      expect(diag.domainStatuses[domain]).toBeDefined();
    }
  });

  it('reasons array mirrors probableCauses descriptions', () => {
    const diag = buildDiag({ keywords: ['PTU'] });
    const causeDescriptions = diag.probableCauses.map((c) => c.description);
    expect(diag.reasons).toEqual(causeDescriptions);
  });

  it('suggestions array mirrors retrySuggestions actions', () => {
    const diag = buildDiag({ keywords: ['PTU'] });
    const suggestionActions = diag.retrySuggestions.map((s) => s.action);
    expect(diag.suggestions).toEqual(suggestionActions);
  });
});

// ---------------------------------------------------------------------------
// Likelihood ranking
// ---------------------------------------------------------------------------

describe('likelihood ranking', () => {
  it('probableCauses are sorted high → medium → low', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const weights = { high: 3, medium: 2, low: 1 } as const;
    for (let i = 0; i < diag.probableCauses.length - 1; i++) {
      expect(weights[diag.probableCauses[i].likelihood]).toBeGreaterThanOrEqual(
        weights[diag.probableCauses[i + 1].likelihood],
      );
    }
  });

  it('retrySuggestions are sorted high → medium → low', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const weights = { high: 3, medium: 2, low: 1 } as const;
    for (let i = 0; i < diag.retrySuggestions.length - 1; i++) {
      expect(weights[diag.retrySuggestions[i].likelihood]).toBeGreaterThanOrEqual(
        weights[diag.retrySuggestions[i + 1].likelihood],
      );
    }
  });

  it('domain timeout causes are ranked high', () => {
    const diag = buildDiag(
      { keywords: ['Q4'] },
      { calendar: { status: 'timeout' } },
    );
    const timeoutCause = diag.probableCauses.find((c) => c.code === 'DOMAIN_TIMEOUT');
    expect(timeoutCause).toBeDefined();
    expect(timeoutCause!.likelihood).toBe('high');
  });

  it('domain error causes are ranked high', () => {
    const diag = buildDiag(
      {},
      { transcripts: { status: 'error', error: 'API unavailable' } },
    );
    const errorCause = diag.probableCauses.find((c) => c.code === 'DOMAIN_ERROR');
    expect(errorCause).toBeDefined();
    expect(errorCause!.likelihood).toBe('high');
  });

  it('keyword mismatch cause is ranked high when keywords provided', () => {
    const diag = buildDiag({ keywords: ['synergy'] });
    const cause = diag.probableCauses.find((c) => c.code === 'KEYWORD_MISMATCH');
    expect(cause).toBeDefined();
    expect(cause!.likelihood).toBe('high');
  });
});

// ---------------------------------------------------------------------------
// Keyword scenarios
// ---------------------------------------------------------------------------

describe('keyword scenarios', () => {
  it('includes keyword mismatch cause with the searched keywords', () => {
    const diag = buildDiag({ keywords: ['synergy', 'roadmap'] });
    const cause = diag.probableCauses.find((c) => c.code === 'KEYWORD_MISMATCH');
    expect(cause).toBeDefined();
    expect(cause!.description).toContain('synergy');
    expect(cause!.description).toContain('roadmap');
  });

  it('includes TRY_ALTERNATE_KEYWORDS suggestion with searched terms', () => {
    const diag = buildDiag({ keywords: ['WSIB', 'PTU'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'TRY_ALTERNATE_KEYWORDS');
    expect(sug).toBeDefined();
    expect(sug!.action).toContain('WSIB');
    expect(sug!.action).toContain('PTU');
  });

  it('alternate keyword suggestion references "alternate title" as informationNeeded', () => {
    const diag = buildDiag({ keywords: ['project-x'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'TRY_ALTERNATE_KEYWORDS');
    expect(sug?.informationNeeded).toBe('alternate title');
  });

  it('includes REDUCE_KEYWORD_COUNT suggestion when multiple keywords given', () => {
    const diag = buildDiag({ keywords: ['alpha', 'beta', 'gamma'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'REDUCE_KEYWORD_COUNT');
    expect(sug).toBeDefined();
  });

  it('includes PROVIDE_KEYWORDS suggestion when no keywords given', () => {
    const diag = buildDiag({ dateRange: { start: new Date('2024-01-01'), end: new Date('2024-01-31') } });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_KEYWORDS');
    expect(sug).toBeDefined();
  });

  it('does not include PROVIDE_KEYWORDS suggestion when keywords are given', () => {
    const diag = buildDiag({ keywords: ['meeting'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_KEYWORDS');
    expect(sug).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Date range scenarios
// ---------------------------------------------------------------------------

describe('date range scenarios', () => {
  it('includes PROVIDE_DATE_RANGE suggestion when no date range given', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_DATE_RANGE');
    expect(sug).toBeDefined();
    expect(sug!.informationNeeded).toBe('date');
  });

  it('includes DATE_RANGE_TOO_NARROW cause for sub-day range', () => {
    const start = new Date('2024-02-14T09:00:00Z');
    const end = new Date('2024-02-14T17:00:00Z');
    const diag = buildDiag({ dateRange: { start, end } });
    const cause = diag.probableCauses.find((c) => c.code === 'DATE_RANGE_TOO_NARROW');
    expect(cause).toBeDefined();
    expect(cause!.likelihood).toBe('high');
  });

  it('includes DATE_RANGE_NARROW cause for ≤7 day range', () => {
    const start = new Date('2024-02-12T00:00:00Z');
    const end = new Date('2024-02-14T23:59:59Z');
    const diag = buildDiag({ dateRange: { start, end } });
    const cause = diag.probableCauses.find((c) => c.code === 'DATE_RANGE_NARROW');
    expect(cause).toBeDefined();
    expect(cause!.likelihood).toBe('medium');
  });

  it('includes WIDEN_DATE_RANGE suggestion for narrow range', () => {
    const start = new Date('2024-03-01T00:00:00Z');
    const end = new Date('2024-03-03T23:59:59Z');
    const diag = buildDiag({ dateRange: { start, end } });
    const sug = diag.retrySuggestions.find((s) => s.code === 'WIDEN_DATE_RANGE');
    expect(sug).toBeDefined();
    expect(sug!.action).toContain('2024-03-01');
    expect(sug!.action).toContain('2024-03-03');
  });

  it('date range suggestion references "date" as informationNeeded', () => {
    const start = new Date('2024-03-01T00:00:00Z');
    const end = new Date('2024-03-02T23:59:59Z');
    const diag = buildDiag({ dateRange: { start, end } });
    const sug = diag.retrySuggestions.find(
      (s) => s.code === 'WIDEN_DATE_RANGE' || s.code === 'PROVIDE_DATE_RANGE',
    );
    expect(sug?.informationNeeded).toBe('date');
  });

  it('does not include DATE_RANGE_TOO_NARROW cause for wide date range', () => {
    const start = new Date('2024-01-01T00:00:00Z');
    const end = new Date('2024-03-31T23:59:59Z');
    const diag = buildDiag({ dateRange: { start, end } });
    const cause = diag.probableCauses.find(
      (c) => c.code === 'DATE_RANGE_TOO_NARROW' || c.code === 'DATE_RANGE_NARROW',
    );
    expect(cause).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Organizer scenarios
// ---------------------------------------------------------------------------

describe('organizer scenarios', () => {
  it('includes PROVIDE_ORGANIZER suggestion when no organizer given', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_ORGANIZER');
    expect(sug).toBeDefined();
    expect(sug!.informationNeeded).toBe('organizer');
  });

  it('does not include PROVIDE_ORGANIZER suggestion when organizer is given', () => {
    const diag = buildDiag({ organizer: 'dave@example.com' });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_ORGANIZER');
    expect(sug).toBeUndefined();
  });

  it('external meeting cause has lower likelihood when organizer is provided', () => {
    const diagWith = buildDiag({ organizer: 'host@external.com' });
    const diagWithout = buildDiag({});
    const causeWith = diagWith.probableCauses.find((c) => c.code === 'EXTERNAL_MEETING');
    const causeWithout = diagWithout.probableCauses.find((c) => c.code === 'EXTERNAL_MEETING');
    expect(causeWith).toBeDefined();
    expect(causeWithout).toBeDefined();
    // Providing an organizer lowers the likelihood of EXTERNAL_MEETING
    const weights = { high: 3, medium: 2, low: 1 } as const;
    expect(weights[causeWith!.likelihood]).toBeLessThan(weights[causeWithout!.likelihood]);
  });
});

// ---------------------------------------------------------------------------
// Attendee scenarios
// ---------------------------------------------------------------------------

describe('attendee scenarios', () => {
  it('includes PROVIDE_ATTENDEES suggestion when no attendees given', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_ATTENDEES');
    expect(sug).toBeDefined();
  });

  it('does not include PROVIDE_ATTENDEES suggestion when attendees are given', () => {
    const diag = buildDiag({ attendees: ['carol@example.com'] });
    const sug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_ATTENDEES');
    expect(sug).toBeUndefined();
  });

  it('includes INSUFFICIENT_FILTERS cause when neither attendees nor organizer given', () => {
    const diag = buildDiag({ keywords: ['budget'] });
    const cause = diag.probableCauses.find((c) => c.code === 'INSUFFICIENT_FILTERS');
    expect(cause).toBeDefined();
  });

  it('does not include INSUFFICIENT_FILTERS cause when attendees are given', () => {
    const diag = buildDiag({ attendees: ['carol@example.com'] });
    const cause = diag.probableCauses.find((c) => c.code === 'INSUFFICIENT_FILTERS');
    expect(cause).toBeUndefined();
  });

  it('does not include INSUFFICIENT_FILTERS cause when organizer is given', () => {
    const diag = buildDiag({ organizer: 'dave@example.com' });
    const cause = diag.probableCauses.find((c) => c.code === 'INSUFFICIENT_FILTERS');
    expect(cause).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Infrastructure failure scenarios
// ---------------------------------------------------------------------------

describe('infrastructure failure scenarios', () => {
  it('includes DOMAIN_TIMEOUT cause when a domain times out', () => {
    const diag = buildDiag(
      { keywords: ['Q4'] },
      { calendar: { status: 'timeout' } },
    );
    const cause = diag.probableCauses.find((c) => c.code === 'DOMAIN_TIMEOUT');
    expect(cause).toBeDefined();
    expect(cause!.description).toContain('calendar');
  });

  it('includes DOMAIN_ERROR cause when a domain errors', () => {
    const diag = buildDiag(
      {},
      { email: { status: 'error', error: 'Forbidden' } },
    );
    const cause = diag.probableCauses.find((c) => c.code === 'DOMAIN_ERROR');
    expect(cause).toBeDefined();
    expect(cause!.description).toContain('email');
  });

  it('includes INCREASE_TIMEOUT suggestion when a domain times out', () => {
    const diag = buildDiag(
      {},
      { files: { status: 'timeout' } },
    );
    const sug = diag.retrySuggestions.find((s) => s.code === 'INCREASE_TIMEOUT');
    expect(sug).toBeDefined();
  });

  it('includes CHECK_PERMISSIONS suggestion when a domain errors', () => {
    const diag = buildDiag(
      {},
      { chats: { status: 'error', error: 'Unauthorized' } },
    );
    const sug = diag.retrySuggestions.find((s) => s.code === 'CHECK_PERMISSIONS');
    expect(sug).toBeDefined();
  });

  it('lists multiple timed-out domains in the cause description', () => {
    const diag = buildDiag(
      {},
      {
        calendar: { status: 'timeout' },
        transcripts: { status: 'timeout' },
      },
    );
    const cause = diag.probableCauses.find((c) => c.code === 'DOMAIN_TIMEOUT');
    expect(cause!.description).toContain('calendar');
    expect(cause!.description).toContain('transcripts');
  });
});

// ---------------------------------------------------------------------------
// Meeting nature causes
// ---------------------------------------------------------------------------

describe('meeting nature causes', () => {
  it('always includes EXTERNAL_MEETING cause', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const cause = diag.probableCauses.find((c) => c.code === 'EXTERNAL_MEETING');
    expect(cause).toBeDefined();
  });

  it('always includes NO_RECORDING_OR_TRANSCRIPT cause', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const cause = diag.probableCauses.find((c) => c.code === 'NO_RECORDING_OR_TRANSCRIPT');
    expect(cause).toBeDefined();
  });

  it('always includes NOTES_OUTSIDE_M365 cause', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const cause = diag.probableCauses.find((c) => c.code === 'NOTES_OUTSIDE_M365');
    expect(cause).toBeDefined();
  });

  it('CHECK_EXTERNAL_INVITE suggestion is always present', () => {
    const diag = buildDiag({});
    const sug = diag.retrySuggestions.find((s) => s.code === 'CHECK_EXTERNAL_INVITE');
    expect(sug).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// Combined parameter scenarios
// ---------------------------------------------------------------------------

describe('combined parameter scenarios', () => {
  it('handles request with all parameters specified', () => {
    const diag = buildDiag({
      keywords: ['budget', 'Q4'],
      dateRange: {
        start: new Date('2024-01-15T00:00:00Z'),
        end: new Date('2024-01-15T23:59:59Z'),
      },
      attendees: ['alice@example.com'],
      organizer: 'bob@example.com',
      meetingType: ['Teams'],
    });

    expect(diag.probableCauses.length).toBeGreaterThan(0);
    expect(diag.retrySuggestions.length).toBeGreaterThan(0);
    // KEYWORD_MISMATCH should be present since keywords were provided
    const cause = diag.probableCauses.find((c) => c.code === 'KEYWORD_MISMATCH');
    expect(cause).toBeDefined();
    // PROVIDE_ORGANIZER should NOT be present since organizer was provided
    const orgSug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_ORGANIZER');
    expect(orgSug).toBeUndefined();
  });

  it('handles completely empty request (no filters at all)', () => {
    const diag = buildDiag({});
    expect(diag.probableCauses.length).toBeGreaterThan(0);
    expect(diag.retrySuggestions.length).toBeGreaterThan(0);
    // With no date range, suggest providing one
    const dateSug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_DATE_RANGE');
    expect(dateSug).toBeDefined();
    // With no keywords, suggest providing some
    const kwSug = diag.retrySuggestions.find((s) => s.code === 'PROVIDE_KEYWORDS');
    expect(kwSug).toBeDefined();
  });

  it('handles mixed domain statuses (some success, some timeout, some error)', () => {
    const diag = buildDiag(
      { keywords: ['alpha'] },
      {
        calendar: { status: 'success', resultCount: 0 },
        transcripts: { status: 'timeout' },
        chats: { status: 'error', error: 'Forbidden' },
        email: { status: 'success', resultCount: 0 },
        files: { status: 'success', resultCount: 0 },
      },
    );

    const timeoutCause = diag.probableCauses.find((c) => c.code === 'DOMAIN_TIMEOUT');
    const errorCause = diag.probableCauses.find((c) => c.code === 'DOMAIN_ERROR');
    expect(timeoutCause).toBeDefined();
    expect(errorCause).toBeDefined();
    expect(timeoutCause!.description).toContain('transcripts');
    expect(errorCause!.description).toContain('chats');
  });
});

// ---------------------------------------------------------------------------
// Human-readable text formatting
// ---------------------------------------------------------------------------

describe('formatDiagnosticText', () => {
  it('returns a non-empty string', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(typeof text).toBe('string');
    expect(text.length).toBeGreaterThan(0);
  });

  it('includes the summary message', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toContain(diag.message);
  });

  it('includes "Probable causes" section header', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toContain('Probable causes');
  });

  it('includes "Suggested next steps" section header', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toContain('Suggested next steps');
  });

  it('includes domain search status section', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toContain('Domain search status');
  });

  it('shows timed-out domains as "timed out"', () => {
    const diag = buildDiag(
      {},
      { calendar: { status: 'timeout' } },
    );
    const text = formatDiagnosticText(diag);
    expect(text).toContain('calendar: timed out');
  });

  it('shows errored domains with the error message', () => {
    const diag = buildDiag(
      {},
      { email: { status: 'error', error: 'Forbidden' } },
    );
    const text = formatDiagnosticText(diag);
    expect(text).toContain('email: error — Forbidden');
  });

  it('shows successful domains with result counts', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toContain('calendar: searched OK (0 result(s))');
  });

  it('includes likelihood labels in output', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    expect(text).toMatch(/\[HIGH\]|\[MEDIUM\]|\[LOW\]/);
  });

  it('includes informationNeeded in suggestions when present', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    const text = formatDiagnosticText(diag);
    // PROVIDE_DATE_RANGE has informationNeeded: 'date'
    expect(text).toContain('information needed: date');
  });

  it('produces output suitable for API JSON serialization', () => {
    const diag = buildDiag({ keywords: ['Q4'] });
    // Should be directly JSON-serializable (no circular refs, no functions)
    const json = JSON.stringify(diag);
    const parsed = JSON.parse(json);
    expect(parsed.message).toBe(diag.message);
    expect(Array.isArray(parsed.probableCauses)).toBe(true);
    expect(Array.isArray(parsed.retrySuggestions)).toBe(true);
  });
});
