import { GraphClient } from './graph-client';
import {
  FileSearchOptions,
  FileSearchResponse,
  FileSearchResult,
  GraphDriveItemResource,
  GraphSearchHit,
  GraphSearchResponse,
  MeetingCorrelation,
  SUPPORTED_FILE_TYPES,
  SupportedFileType,
} from './types';

const DEFAULT_MAX_RESULTS = 25;
const MAX_ALLOWED_RESULTS = 100;

/**
 * Naming-convention patterns that suggest a file is associated with a meeting.
 * The regex tests against the lower-cased file name.
 */
const MEETING_FILE_PATTERNS = [
  /\bmeeting[\s_-]?notes?\b/,
  /\bagenda\b/,
  /\bminutes?\b/,
  /\baction[\s_-]?items?\b/,
  /\bdeck\b/,
  /\bpresentation\b/,
  /\brecap\b/,
  /\bfollow[\s_-]?up\b/,
];

/**
 * Client for searching files across a user's OneDrive and accessible SharePoint
 * sites using the Microsoft Graph `/search/query` API.
 *
 * @example
 * ```typescript
 * const client = new OnedriveSearchClient(accessToken);
 * const results = await client.searchFiles({
 *   query: 'project kickoff',
 *   fileTypes: ['docx', 'pptx'],
 *   dateFrom: new Date('2024-01-01'),
 *   sources: ['onedrive', 'sharepoint'],
 * });
 * console.log(results.results);
 * ```
 */
export class OnedriveSearchClient {
  private readonly graphClient: GraphClient;

  constructor(accessToken: string) {
    this.graphClient = new GraphClient(accessToken);
  }

  /**
   * Search for files across OneDrive and/or SharePoint.
   *
   * @param options - Search parameters (query, filters, etc.).
   * @returns Aggregated file search results.
   */
  async searchFiles(options: FileSearchOptions): Promise<FileSearchResponse> {
    const {
      query,
      fileTypes = [...SUPPORTED_FILE_TYPES],
      dateFrom,
      dateTo,
      maxResults = DEFAULT_MAX_RESULTS,
      meetingTitles = [],
    } = options;

    const effectiveMaxResults = Math.min(maxResults, MAX_ALLOWED_RESULTS);

    // Build the KQL query string including optional file-type and date filters.
    const kqlQuery = buildKqlQuery(query, fileTypes, dateFrom, dateTo);

    const requestBody = {
      requests: [
        {
          entityTypes: ['driveItem'],
          query: { queryString: kqlQuery },
          fields: [
            'id',
            'name',
            'webUrl',
            'size',
            'lastModifiedDateTime',
            'createdBy',
            'lastModifiedBy',
            'parentReference',
            'file',
            'fileSystemInfo',
          ],
          from: 0,
          size: effectiveMaxResults,
        },
      ],
    };

    const response = await this.graphClient.post<GraphSearchResponse>('/search/query', requestBody);

    return parseSearchResponse(response, meetingTitles);
  }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Build a KQL (Keyword Query Language) string for the Graph search API.
 *
 * Graph search supports KQL operators such as:
 *   - `FileType:docx`
 *   - `LastModifiedTime>=2024-01-01`
 */
function buildKqlQuery(
  query: string,
  fileTypes: SupportedFileType[],
  dateFrom?: Date,
  dateTo?: Date,
): string {
  const parts: string[] = [query.trim()];

  if (fileTypes.length > 0 && fileTypes.length < SUPPORTED_FILE_TYPES.length) {
    // Build a disjunction only when a subset of types is requested.
    const typeClause = fileTypes.map((t) => `FileType:${t}`).join(' OR ');
    parts.push(`(${typeClause})`);
  }

  if (dateFrom) {
    parts.push(`LastModifiedTime>=${formatDate(dateFrom)}`);
  }

  if (dateTo) {
    parts.push(`LastModifiedTime<=${formatDate(dateTo)}`);
  }

  return parts.join(' AND ');
}

/** Format a Date as `YYYY-MM-DD` for KQL queries. */
function formatDate(date: Date): string {
  return date.toISOString().slice(0, 10);
}

/**
 * Parse a raw Graph search response into a {@link FileSearchResponse}.
 */
function parseSearchResponse(
  raw: GraphSearchResponse,
  meetingTitles: string[],
): FileSearchResponse {
  let totalCount = 0;
  let hasMore = false;
  const results: FileSearchResult[] = [];

  for (const responseItem of raw.value) {
    for (const container of responseItem.hitsContainers) {
      totalCount += container.total ?? 0;
      hasMore = hasMore || (container.moreResultsAvailable ?? false);

      for (const hit of container.hits ?? []) {
        const result = hitToFileResult(hit, meetingTitles);
        if (result) {
          results.push(result);
        }
      }
    }
  }

  return { results, totalCount, hasMore };
}

/**
 * Convert a single Graph search hit into a {@link FileSearchResult}.
 * Returns `null` if the hit cannot be mapped (e.g. unsupported entity type).
 */
function hitToFileResult(
  hit: GraphSearchHit,
  meetingTitles: string[],
): FileSearchResult | null {
  const resource: GraphDriveItemResource = hit.resource;

  if (!resource['@odata.type']?.includes('driveItem')) {
    return null;
  }

  const name = resource.name ?? '';
  const webUrl = resource.webUrl ?? '';
  const siteId = resource.parentReference?.siteId;
  const source = siteId ? 'sharepoint' : 'onedrive';
  const lastModified =
    resource.lastModifiedDateTime ??
    resource.fileSystemInfo?.lastModifiedDateTime ??
    '';
  const author =
    resource.lastModifiedBy?.user?.displayName ??
    resource.createdBy?.user?.displayName ??
    '';
  const path = buildPath(resource);
  const fileType = extractFileType(name);
  const contentSnippet = hit.summary ?? '';
  const size = resource.size ?? 0;

  const meetingCorrelation = correlateMeeting(name, path, meetingTitles);

  return {
    name,
    path,
    webUrl,
    lastModified,
    author,
    contentSnippet,
    fileType,
    size,
    driveItemId: resource.id,
    ...(siteId ? { siteId } : {}),
    source,
    ...(meetingCorrelation ? { meetingCorrelation } : {}),
  };
}

/**
 * Build a human-readable path from the parentReference of a drive item.
 * Graph returns paths like `/drives/<id>/root:/Documents/Meetings` — we
 * extract the portion after `root:`.
 */
function buildPath(resource: GraphDriveItemResource): string {
  const raw = resource.parentReference?.path ?? '';
  const rootIndex = raw.indexOf('root:');
  if (rootIndex !== -1) {
    return raw.slice(rootIndex + 5) || '/';
  }
  return raw || '/';
}

/** Extract the lower-cased file extension from a file name. */
function extractFileType(name: string): string {
  const lastDot = name.lastIndexOf('.');
  if (lastDot === -1) return '';
  return name.slice(lastDot + 1).toLowerCase();
}

/**
 * Attempt to correlate a file to a meeting by:
 *   1. Checking whether the file name matches known meeting-file naming patterns.
 *   2. Checking whether any of the provided meeting titles appear in the file name.
 *
 * Returns a {@link MeetingCorrelation} when a match is found, otherwise `undefined`.
 */
function correlateMeeting(
  name: string,
  path: string,
  meetingTitles: string[],
): MeetingCorrelation | undefined {
  const lowerName = name.toLowerCase();
  const lowerPath = path.toLowerCase();

  // Check explicit meeting-title matches first (highest confidence).
  for (const title of meetingTitles) {
    if (!title.trim()) continue;
    const lowerTitle = title.toLowerCase();
    if (lowerName.includes(lowerTitle) || lowerPath.includes(lowerTitle)) {
      return { matchedMeetingTitle: title, confidence: 'high' };
    }
    // Partial word match → medium confidence.
    const words = lowerTitle.split(/\s+/).filter(Boolean);
    if (words.length > 0 && words.every((w) => lowerName.includes(w))) {
      return { matchedMeetingTitle: title, confidence: 'medium' };
    }
  }

  // Check naming-convention patterns → low confidence.
  for (const pattern of MEETING_FILE_PATTERNS) {
    if (pattern.test(lowerName)) {
      return { matchedMeetingTitle: name, confidence: 'low' };
    }
  }

  return undefined;
}
