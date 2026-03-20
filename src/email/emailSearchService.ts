import {
  EmailAttachment,
  EmailResult,
  EmailSearchParams,
  EmailSearchResponse,
  GraphAttachment,
  GraphMessage,
} from './emailSearchTypes';
import { detectMeetingEmail } from './meetingDetector';
import { GraphClientWrapper } from '../graph/graphClient';

/** Default maximum number of results per search. */
const DEFAULT_MAX_RESULTS = 25;

/**
 * Validates that a string is a well-formed email address.
 * Throws if the value does not match a basic email pattern.
 */
function validateEmailAddress(value: string): void {
  const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!EMAIL_PATTERN.test(value)) {
    throw new Error(`Invalid email address supplied for filter: "${value}"`);
  }
}

/**
 * Sanitizes a single search keyword so it is safe to embed inside an OData
 * `$search` quoted string.  Only alphanumeric characters, spaces, hyphens, and
 * underscores are kept; all other characters are stripped.
 */
function sanitizeKeyword(keyword: string): string {
  return keyword.replace(/[^a-zA-Z0-9 \-_]/g, '');
}

/**
 * Builds an OData `$filter` expression from the supplied search parameters.
 * Only non-empty filter clauses are included.
 */
function buildFilterExpression(params: EmailSearchParams): string {
  const clauses: string[] = [];

  if (params.dateFrom) {
    clauses.push(`receivedDateTime ge ${params.dateFrom.toISOString()}`);
  }

  if (params.dateTo) {
    clauses.push(`receivedDateTime le ${params.dateTo.toISOString()}`);
  }

  if (params.sender) {
    // Validate the sender is a well-formed email address before embedding it
    // in the OData filter to guard against injection.
    validateEmailAddress(params.sender);
    clauses.push(
      `from/emailAddress/address eq '${params.sender}'`,
    );
  }

  return clauses.join(' and ');
}

/**
 * Builds an OData `$search` expression from the supplied keywords.
 * The search targets subject and body fields.
 *
 * @see https://learn.microsoft.com/en-us/graph/search-query-parameter
 */
function buildSearchExpression(keywords: string[]): string {
  // Each keyword is sanitized and wrapped in double-quotes to search for the
  // exact token while preventing OData injection.
  const terms = keywords
    .map((k) => sanitizeKeyword(k))
    .filter((k) => k.length > 0)
    .map((k) => `"${k}"`)
    .join(' OR ');
  return terms;
}

/**
 * Converts a raw Graph attachment object into the typed EmailAttachment shape.
 */
function mapAttachment(raw: GraphAttachment): EmailAttachment {
  return {
    id: raw.id,
    name: raw.name ?? '',
    contentType: raw.contentType ?? '',
    size: raw.size ?? 0,
  };
}

/**
 * Converts a raw Graph message into a typed EmailResult, running meeting
 * detection along the way.
 */
function mapMessageToEmailResult(raw: GraphMessage): EmailResult {
  const sender =
    raw.from?.emailAddress?.address ??
    raw.from?.emailAddress?.name ??
    '';

  const recipients = [
    ...(raw.toRecipients ?? []),
    ...(raw.ccRecipients ?? []),
  ]
    .map((r) => r.emailAddress?.address ?? r.emailAddress?.name ?? '')
    .filter(Boolean);

  const attachments: EmailAttachment[] = (raw.attachments ?? []).map(mapAttachment);

  const { isMeetingRelated, meetingType } = detectMeetingEmail(raw);

  const bodyPreview = raw.bodyPreview ?? '';

  return {
    id: raw.id,
    subject: raw.subject ?? '',
    sender,
    receivedDateTime: raw.receivedDateTime ?? '',
    snippet: bodyPreview.slice(0, 200),
    isMeetingRelated,
    ...(meetingType !== undefined ? { meetingType } : {}),
    attachments,
    recipients,
    bodyPreview,
  };
}

/**
 * Service for searching Outlook emails via the Microsoft Graph API.
 *
 * All filtering, keyword searching, and meeting-detection logic is
 * encapsulated here; callers only need to provide a configured
 * {@link GraphClientWrapper}.
 */
export class EmailSearchService {
  constructor(private readonly graphClient: GraphClientWrapper) {}

  /**
   * Searches emails according to the supplied parameters and returns a
   * structured {@link EmailSearchResponse}.
   *
   * @param params  Search parameters (keywords, date range, sender, etc.).
   */
  async search(params: EmailSearchParams): Promise<EmailSearchResponse> {
    const queryParams: Record<string, string | number> = {
      $top: params.maxResults ?? DEFAULT_MAX_RESULTS,
    };

    // Build the $search parameter from keywords.
    if (params.keywords && params.keywords.length > 0) {
      queryParams['$search'] = buildSearchExpression(params.keywords);
    }

    // Build the $filter parameter from date range and sender.
    const filterExpression = buildFilterExpression(params);
    if (filterExpression) {
      queryParams['$filter'] = filterExpression;
    }

    // Expand attachments when the caller has asked for them.
    if (params.includeAttachments) {
      queryParams['$expand'] = 'attachments';
    }

    // Request relevant fields to reduce response size.
    queryParams['$select'] =
      'id,subject,bodyPreview,receivedDateTime,from,toRecipients,ccRecipients,hasAttachments,categories';

    const response = await this.graphClient.getMessages(queryParams);
    const rawMessages = (response.value ?? []) as GraphMessage[];

    // Post-filter by recipient if requested (Graph $filter doesn't support
    // searching inside recipient collections without advanced query support).
    let filteredMessages = rawMessages;
    if (params.recipients && params.recipients.length > 0) {
      const lowerRecipients = params.recipients.map((r) => r.toLowerCase());
      filteredMessages = rawMessages.filter((msg) => {
        const allRecipients = [
          ...(msg.toRecipients ?? []),
          ...(msg.ccRecipients ?? []),
        ];
        return allRecipients.some((r) => {
          const addr = (r.emailAddress?.address ?? '').toLowerCase();
          return lowerRecipients.includes(addr);
        });
      });
    }

    const results = filteredMessages.map(mapMessageToEmailResult);

    return {
      results,
      totalCount: results.length,
      searchParams: params,
    };
  }
}
