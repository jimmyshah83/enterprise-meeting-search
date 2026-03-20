/**
 * Types and interfaces for the email search module.
 */

/** Parameters for searching emails. */
export interface EmailSearchParams {
  /** Keywords to search across subject, body, and attachment names. */
  keywords?: string[];
  /** Start of the date range filter (inclusive). */
  dateFrom?: Date;
  /** End of the date range filter (inclusive). */
  dateTo?: Date;
  /** Filter by sender email address. */
  sender?: string;
  /** Filter by recipient email addresses (any match). */
  recipients?: string[];
  /** Whether to include attachment metadata in results. Defaults to false. */
  includeAttachments?: boolean;
  /** Maximum number of results to return. Defaults to 25. */
  maxResults?: number;
}

/** Metadata for an email attachment. */
export interface EmailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
}

/** The type of meeting correlation detected in an email. */
export type MeetingEmailType = 'invite' | 'forward' | 'reply' | 'cancellation';

/** A single email result returned by the search. */
export interface EmailResult {
  id: string;
  subject: string;
  sender: string;
  receivedDateTime: string;
  /** Short text excerpt from the email body. */
  snippet: string;
  /** Whether this email was detected as meeting-related. */
  isMeetingRelated: boolean;
  /** The kind of meeting-related email, if detected. */
  meetingType?: MeetingEmailType;
  attachments: EmailAttachment[];
  recipients: string[];
  bodyPreview: string;
}

/** The structured response returned by the email search service. */
export interface EmailSearchResponse {
  results: EmailResult[];
  totalCount: number;
  searchParams: EmailSearchParams;
}

/** Raw message shape returned by Microsoft Graph API /me/messages. */
export interface GraphMessage {
  id: string;
  subject?: string;
  bodyPreview?: string;
  receivedDateTime?: string;
  from?: {
    emailAddress?: {
      address?: string;
      name?: string;
    };
  };
  toRecipients?: Array<{
    emailAddress?: {
      address?: string;
      name?: string;
    };
  }>;
  ccRecipients?: Array<{
    emailAddress?: {
      address?: string;
      name?: string;
    };
  }>;
  hasAttachments?: boolean;
  attachments?: GraphAttachment[];
  internetMessageHeaders?: Array<{ name: string; value: string }>;
  categories?: string[];
  importance?: string;
  inferenceClassification?: string;
}

/** Raw attachment shape returned by Microsoft Graph API. */
export interface GraphAttachment {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
}
