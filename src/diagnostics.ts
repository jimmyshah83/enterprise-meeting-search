/**
 * Diagnostic and suggestion engine for empty search results.
 *
 * When all search domains return no results, this module analyses the
 * original search parameters and the per-domain execution statuses to
 * produce:
 *
 *  1. A list of **probable causes** ranked by likelihood – explaining *why*
 *     the search came up empty (e.g. keyword mismatch, meeting is external,
 *     no recording was made, notes stored outside M365).
 *
 *  2. A list of **retry suggestions** ranked by expected impact – giving
 *     the user concrete, actionable steps to find the meeting (e.g.
 *     "provide a date range", "specify the organiser", "try alternate
 *     keywords").
 *
 * The output is a {@link DiagnosticSummary} that is serialisable as JSON
 * (suitable for REST API responses) and can also be rendered as a
 * human-readable string via {@link formatDiagnosticText}.
 */

import {
  DiagnosticSummary,
  DomainName,
  DomainStatus,
  Likelihood,
  ProbableCause,
  RetrySuggestion,
  SearchRequest,
} from './types';

// ---------------------------------------------------------------------------
// Likelihood helpers
// ---------------------------------------------------------------------------

/** Numeric weight used to sort by likelihood (high → low). */
const LIKELIHOOD_WEIGHT: Record<Likelihood, number> = {
  high: 3,
  medium: 2,
  low: 1,
};

function sortByLikelihood<T extends { likelihood: Likelihood }>(items: T[]): T[] {
  return [...items].sort(
    (a, b) => LIKELIHOOD_WEIGHT[b.likelihood] - LIKELIHOOD_WEIGHT[a.likelihood],
  );
}

// ---------------------------------------------------------------------------
// Cause builders
// ---------------------------------------------------------------------------

/**
 * Analyses domain-level execution statuses and returns causes attributable
 * to infrastructure problems (timeouts, errors).
 */
function buildInfrastructureCauses(
  statuses: Record<DomainName, DomainStatus>,
): ProbableCause[] {
  const causes: ProbableCause[] = [];

  const timedOut = (Object.entries(statuses) as [DomainName, DomainStatus][])
    .filter(([, s]) => s.status === 'timeout')
    .map(([name]) => name);

  const errored = (Object.entries(statuses) as [DomainName, DomainStatus][])
    .filter(([, s]) => s.status === 'error')
    .map(([name]) => name);

  if (timedOut.length > 0) {
    causes.push({
      code: 'DOMAIN_TIMEOUT',
      description:
        `One or more data sources did not respond in time (${timedOut.join(', ')}). ` +
        'Results may be incomplete.',
      likelihood: 'high',
    });
  }

  if (errored.length > 0) {
    causes.push({
      code: 'DOMAIN_ERROR',
      description:
        `One or more data sources returned an error (${errored.join(', ')}). ` +
        'Access permissions or service availability may be the cause.',
      likelihood: 'high',
    });
  }

  return causes;
}

/**
 * Analyses the search keywords and returns causes related to the keyword
 * matching strategy.
 */
function buildKeywordCauses(request: SearchRequest): ProbableCause[] {
  const causes: ProbableCause[] = [];
  const keywords = request.keywords ?? [];

  if (keywords.length === 0) {
    return causes;
  }

  causes.push({
    code: 'KEYWORD_MISMATCH',
    description:
      `The meeting may have been recorded or documented under a different title. ` +
      `The keywords searched were: "${keywords.join('", "')}".`,
    likelihood: 'high',
  });

  causes.push({
    code: 'ABBREVIATION_OR_ALIAS',
    description:
      'The meeting title or project name might be known by an abbreviation, ' +
      'acronym, or alias that differs from the keywords provided.',
    likelihood: 'medium',
  });

  return causes;
}

/**
 * Analyses the date range filter and returns causes related to the temporal
 * search window.
 */
function buildDateRangeCauses(request: SearchRequest): ProbableCause[] {
  const causes: ProbableCause[] = [];

  if (!request.dateRange) {
    causes.push({
      code: 'NO_DATE_RANGE',
      description:
        'No date range was specified. Without a date constraint the search ' +
        'spans all available history, making it harder to isolate results — ' +
        'but also means the meeting date may differ from expectations.',
      likelihood: 'medium',
    });
    return causes;
  }

  const { start, end } = request.dateRange;
  const rangeMs = end.getTime() - start.getTime();
  const rangeDays = rangeMs / (1000 * 60 * 60 * 24);

  // Always emit a cause that mentions the specific dates so callers can
  // surface them in diagnostics text.
  if (rangeDays < 1) {
    causes.push({
      code: 'DATE_RANGE_TOO_NARROW',
      description:
        `The date range is very narrow (${start.toISOString().slice(0, 10)} – ` +
        `${end.toISOString().slice(0, 10)}). The meeting may have occurred on a ` +
        'different day.',
      likelihood: 'high',
    });
  } else if (rangeDays <= 7) {
    causes.push({
      code: 'DATE_RANGE_NARROW',
      description:
        `The date range covers only ${Math.round(rangeDays)} day(s) ` +
        `(${start.toISOString().slice(0, 10)} – ${end.toISOString().slice(0, 10)}). ` +
        'The meeting may have taken place just outside this window.',
      likelihood: 'medium',
    });
  } else {
    causes.push({
      code: 'DATE_RANGE_NO_MATCH',
      description:
        `No meetings were found in the date range ${start.toISOString().slice(0, 10)} – ` +
        `${end.toISOString().slice(0, 10)}. The meeting may have occurred ` +
        'outside this window or under a different time zone.',
      likelihood: 'medium',
    });
  }

  return causes;
}

/**
 * Returns causes related to the nature of the meeting itself (external
 * participants, no recording, notes stored outside M365).
 */
function buildMeetingNatureCauses(request: SearchRequest): ProbableCause[] {
  const causes: ProbableCause[] = [];

  causes.push({
    code: 'EXTERNAL_MEETING',
    description:
      'The meeting may have been organised by an external party (outside your ' +
      'organisation) and therefore not captured in your Microsoft 365 calendar.',
    likelihood: request.organizer ? 'low' : 'medium',
  });

  causes.push({
    code: 'NO_RECORDING_OR_TRANSCRIPT',
    description:
      'The meeting may not have been recorded, or the recording / transcript ' +
      'may not have been saved to Microsoft Stream or OneDrive.',
    likelihood: 'medium',
  });

  causes.push({
    code: 'NOTES_OUTSIDE_M365',
    description:
      'Meeting notes, agendas, or minutes may have been stored in a tool ' +
      'outside Microsoft 365 (e.g. Confluence, Notion, Google Drive) and are ' +
      'therefore not searchable by this system.',
    likelihood: 'low',
  });

  if (request.attendees && request.attendees.length > 0) {
    causes.push({
      code: 'ATTENDEE_NOT_FOUND',
      description:
        `No meetings were found involving attendees: ${request.attendees.join(', ')}. ` +
        'Verify the email addresses are correct, or the attendee may have been ' +
        'invited under a different address.',
      likelihood: 'medium',
    });
  }

  if (request.organizer) {
    causes.push({
      code: 'ORGANIZER_NOT_FOUND',
      description:
        `No meetings were found organised by: ${request.organizer}. ` +
        'Verify the organiser email address is correct.',
      likelihood: 'medium',
    });
  }

  if (!request.attendees && !request.organizer) {
    causes.push({
      code: 'INSUFFICIENT_FILTERS',
      description:
        'No attendee or organiser filter was provided. The meeting may exist ' +
        'in a calendar or chat that is not associated with your account.',
      likelihood: 'medium',
    });
  }

  return causes;
}

// ---------------------------------------------------------------------------
// Suggestion builders
// ---------------------------------------------------------------------------

/**
 * Returns suggestions to address infrastructure failures (timeouts / errors).
 */
function buildInfrastructureSuggestions(
  statuses: Record<DomainName, DomainStatus>,
): RetrySuggestion[] {
  const suggestions: RetrySuggestion[] = [];

  const hasTimeout = Object.values(statuses).some((s) => s.status === 'timeout');
  const hasError = Object.values(statuses).some((s) => s.status === 'error');

  if (hasTimeout) {
    suggestions.push({
      code: 'INCREASE_TIMEOUT',
      action:
        'Retry the search with a higher `domainTimeoutMs` value, or try again ' +
        'when the service load is lower.',
      likelihood: 'high',
    });
  }

  if (hasError) {
    suggestions.push({
      code: 'CHECK_PERMISSIONS',
      action:
        'Verify that the application has the required Microsoft 365 permissions ' +
        '(Calendars.Read, Chat.Read, Mail.Read, Files.Read, OnlineMeetings.Read).',
      likelihood: 'high',
    });
  }

  return suggestions;
}

/**
 * Returns suggestions related to providing or broadening date information.
 */
function buildDateSuggestions(request: SearchRequest): RetrySuggestion[] {
  const suggestions: RetrySuggestion[] = [];

  if (!request.dateRange) {
    suggestions.push({
      code: 'PROVIDE_DATE_RANGE',
      action:
        'Provide an approximate date or date range for the meeting ' +
        '(e.g. "early February", "last Monday", or a specific YYYY-MM-DD range). ' +
        'This significantly reduces the search space and improves accuracy.',
      informationNeeded: 'date',
      likelihood: 'high',
    });
    return suggestions;
  }

  const { start, end } = request.dateRange;
  const rangeDays =
    (end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24);

  if (rangeDays <= 7) {
    suggestions.push({
      code: 'WIDEN_DATE_RANGE',
      action:
        `The current date range is only ${Math.round(rangeDays)} day(s) ` +
        `(${start.toISOString().slice(0, 10)} – ${end.toISOString().slice(0, 10)}). ` +
        'Try widening the range by ±1 week to catch meetings that may have ' +
        'been rescheduled.',
      informationNeeded: 'date',
      likelihood: 'high',
    });
  }

  return suggestions;
}

/**
 * Returns suggestions related to providing organiser information.
 */
function buildOrganizerSuggestions(request: SearchRequest): RetrySuggestion[] {
  if (request.organizer) return [];

  return [
    {
      code: 'PROVIDE_ORGANIZER',
      action:
        'If you know who organised the meeting, provide their email address ' +
        'via the `organizer` filter. This narrows the search to their calendar ' +
        'and is one of the most effective ways to locate a specific meeting.',
      informationNeeded: 'organizer',
      likelihood: 'high',
    },
  ];
}

/**
 * Returns suggestions related to providing attendee information.
 */
function buildAttendeeSuggestions(request: SearchRequest): RetrySuggestion[] {
  if (request.attendees && request.attendees.length > 0) return [];

  return [
    {
      code: 'PROVIDE_ATTENDEES',
      action:
        'Provide one or more attendee email addresses via the `attendees` ' +
        'filter. Even a single known attendee can dramatically reduce the ' +
        'candidate meeting set.',
      informationNeeded: 'attendees',
      likelihood: 'medium',
    },
  ];
}

/**
 * Returns suggestions related to keyword strategy.
 */
function buildKeywordSuggestions(request: SearchRequest): RetrySuggestion[] {
  const suggestions: RetrySuggestion[] = [];
  const keywords = request.keywords ?? [];

  if (keywords.length === 0) {
    suggestions.push({
      code: 'PROVIDE_KEYWORDS',
      action:
        'Try adding keywords related to the meeting topic, project name, ' +
        'or a phrase you remember being used during the meeting.',
      informationNeeded: 'alternate title',
      likelihood: 'medium',
    });
    return suggestions;
  }

  suggestions.push({
    code: 'TRY_ALTERNATE_KEYWORDS',
    action:
      `The keywords "${keywords.join('", "')}" returned no matches. ` +
      'Try alternate terms, abbreviations, or the full project/initiative name. ' +
      'For example, if you searched for an acronym, try the full phrase, ' +
      'or vice versa.',
    informationNeeded: 'alternate title',
    likelihood: 'high',
  });

  suggestions.push({
    code: 'REDUCE_KEYWORD_COUNT',
    action:
      `You used ${keywords.length} keyword(s). Try searching with only the ` +
      'most distinctive single keyword to cast a wider net, then refine.',
    informationNeeded: 'alternate title',
    likelihood: 'medium',
  });

  return suggestions;
}

/**
 * Returns suggestions when the meeting might be external or undocumented.
 */
function buildMeetingNatureSuggestions(request: SearchRequest): RetrySuggestion[] {
  const suggestions: RetrySuggestion[] = [];

  suggestions.push({
    code: 'CHECK_EXTERNAL_INVITE',
    action:
      'The meeting may have been organised by someone outside your ' +
      'organisation. Check your email inbox for a forwarded invite or ' +
      'look for the meeting in the organiser\'s calendar directly.',
    likelihood: 'medium',
  });

  if (!request.meetingType) {
    suggestions.push({
      code: 'SPECIFY_MEETING_TYPE',
      action:
        'Specify the meeting type (e.g. "Teams", "InPerson", "Zoom") to ' +
        'restrict the search to the relevant platform and avoid false negatives.',
      informationNeeded: 'meeting type',
      likelihood: 'low',
    });
  }

  return suggestions;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Analyses a failed (empty-result) search and produces a structured
 * {@link DiagnosticSummary} with probable causes and actionable retry
 * suggestions, both ranked by likelihood.
 *
 * @param request        The original search request that yielded no results.
 * @param domainStatuses Per-domain execution results from the orchestrator.
 * @returns              A {@link DiagnosticSummary} ready for API or UI consumption.
 */
export function buildDiagnostics(
  request: SearchRequest,
  domainStatuses: Record<DomainName, DomainStatus>,
): DiagnosticSummary {
  // --- Collect probable causes ---
  const rawCauses: ProbableCause[] = [
    ...buildInfrastructureCauses(domainStatuses),
    ...buildKeywordCauses(request),
    ...buildDateRangeCauses(request),
    ...buildMeetingNatureCauses(request),
  ];
  const probableCauses = sortByLikelihood(rawCauses);

  // --- Collect retry suggestions ---
  const rawSuggestions: RetrySuggestion[] = [
    ...buildInfrastructureSuggestions(domainStatuses),
    ...buildDateSuggestions(request),
    ...buildOrganizerSuggestions(request),
    ...buildKeywordSuggestions(request),
    ...buildAttendeeSuggestions(request),
    ...buildMeetingNatureSuggestions(request),
  ];
  const retrySuggestions = sortByLikelihood(rawSuggestions);

  // --- Backwards-compatible flat arrays ---
  const reasons = probableCauses.map((c) => c.description);
  const suggestions = retrySuggestions.map((s) => s.action);

  return {
    message: 'No results found across all search domains.',
    probableCauses,
    retrySuggestions,
    domainStatuses,
    reasons,
    suggestions,
  };
}

/**
 * Converts a {@link DiagnosticSummary} to a human-readable multi-line string
 * suitable for display in a conversational UI (e.g. a chatbot or CLI).
 *
 * @param diagnostic  The diagnostic summary produced by {@link buildDiagnostics}.
 * @returns           Formatted plain-text explanation.
 */
export function formatDiagnosticText(diagnostic: DiagnosticSummary): string {
  const lines: string[] = [];

  lines.push(diagnostic.message);
  lines.push('');

  if (diagnostic.probableCauses.length > 0) {
    lines.push('Probable causes:');
    diagnostic.probableCauses.forEach((cause, i) => {
      lines.push(`  ${i + 1}. [${cause.likelihood.toUpperCase()}] ${cause.description}`);
    });
    lines.push('');
  }

  if (diagnostic.retrySuggestions.length > 0) {
    lines.push('Suggested next steps:');
    diagnostic.retrySuggestions.forEach((suggestion, i) => {
      const info = suggestion.informationNeeded
        ? ` (information needed: ${suggestion.informationNeeded})`
        : '';
      lines.push(
        `  ${i + 1}. [${suggestion.likelihood.toUpperCase()}] ${suggestion.action}${info}`,
      );
    });
    lines.push('');
  }

  const domainLines = (
    Object.entries(diagnostic.domainStatuses) as [string, DomainStatus][]
  ).map(([name, status]) => {
    switch (status.status) {
      case 'success':
        return `  ${name}: searched OK (${status.resultCount} result(s))`;
      case 'timeout':
        return `  ${name}: timed out`;
      case 'error':
        return `  ${name}: error — ${status.error}`;
    }
  });

  if (domainLines.length > 0) {
    lines.push('Domain search status:');
    lines.push(...domainLines);
  }

  return lines.join('\n');
}
