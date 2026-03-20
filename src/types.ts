/**
 * Supported document file types for OneDrive / SharePoint search.
 */
export const SUPPORTED_FILE_TYPES = ['docx', 'pptx', 'xlsx', 'pdf', 'txt', 'md'] as const;
export type SupportedFileType = (typeof SUPPORTED_FILE_TYPES)[number];

/**
 * Source of a file result.
 */
export type FileSource = 'onedrive' | 'sharepoint';

/**
 * Confidence level for meeting correlation.
 */
export type CorrelationConfidence = 'high' | 'medium' | 'low';

/**
 * Represents a match between a file and a meeting based on naming conventions.
 */
export interface MeetingCorrelation {
  /** The matched meeting title / keyword that triggered the correlation. */
  matchedMeetingTitle: string;
  /** Confidence level of the correlation. */
  confidence: CorrelationConfidence;
}

/**
 * Represents a single file result returned from a search.
 */
export interface FileSearchResult {
  /** Display name of the file. */
  name: string;
  /** Full path / breadcrumb of the file location (e.g. "Documents > Meetings"). */
  path: string;
  /** Web URL to open the file in a browser. */
  webUrl: string;
  /** ISO-8601 date-time string of the last modification. */
  lastModified: string;
  /** Display name of the author / last-modifier. */
  author: string;
  /** A short preview snippet of relevant file content (may be empty). */
  contentSnippet: string;
  /** File extension without the leading dot (e.g. "docx"). */
  fileType: string;
  /** File size in bytes. */
  size: number;
  /** Graph API drive item ID. */
  driveItemId: string;
  /** Graph API site ID (present for SharePoint results, absent for OneDrive). */
  siteId?: string;
  /** Whether the file came from OneDrive or SharePoint. */
  source: FileSource;
  /** Present when the file can be correlated to a meeting. */
  meetingCorrelation?: MeetingCorrelation;
}

/**
 * Options accepted by {@link OnedriveSearchClient.searchFiles}.
 */
export interface FileSearchOptions {
  /** Free-text keyword query. Supports AND / OR / phrase operators. */
  query: string;
  /**
   * Restrict results to these file extensions (without the leading dot).
   * Defaults to all supported types when omitted.
   */
  fileTypes?: SupportedFileType[];
  /** Inclusive lower bound on the last-modified date. */
  dateFrom?: Date;
  /** Inclusive upper bound on the last-modified date. */
  dateTo?: Date;
  /**
   * Which storage locations to search.
   * Defaults to both 'onedrive' and 'sharepoint' when omitted.
   */
  sources?: FileSource[];
  /** Maximum number of results to return (default: 25, max: 100). */
  maxResults?: number;
  /**
   * When provided, the search result names / paths are checked against these
   * meeting title keywords so that files can be correlated to meetings.
   */
  meetingTitles?: string[];
}

/**
 * Aggregated response from a file search operation.
 */
export interface FileSearchResponse {
  /** All file results from OneDrive and/or SharePoint. */
  results: FileSearchResult[];
  /** Total number of matching items reported by the Graph API. */
  totalCount: number;
  /** Whether there are more results beyond the current page. */
  hasMore: boolean;
}

// ---------------------------------------------------------------------------
// Raw Microsoft Graph API response shapes (internal)
// ---------------------------------------------------------------------------

/** @internal */
export interface GraphSearchHit {
  hitId: string;
  rank: number;
  summary?: string;
  resource: GraphDriveItemResource;
}

/** @internal */
export interface GraphDriveItemResource {
  '@odata.type': string;
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  lastModifiedDateTime?: string;
  fileSystemInfo?: { lastModifiedDateTime?: string };
  parentReference?: {
    path?: string;
    siteId?: string;
    driveId?: string;
  };
  createdBy?: { user?: { displayName?: string } };
  lastModifiedBy?: { user?: { displayName?: string } };
  file?: { mimeType?: string };
}

/** @internal */
export interface GraphSearchResponse {
  value: Array<{
    searchTerms: string[];
    hitsContainers: Array<{
      hits?: GraphSearchHit[];
      total?: number;
      moreResultsAvailable?: boolean;
    }>;
  }>;
}
