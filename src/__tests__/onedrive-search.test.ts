import axios from 'axios';
import { OnedriveSearchClient } from '../onedrive-search';
import {
  FileSearchResponse,
  GraphSearchResponse,
} from '../types';

// ---------------------------------------------------------------------------
// Mock axios so no real HTTP calls are made.
// ---------------------------------------------------------------------------
jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

/** Helper that builds a minimal mock axios instance. */
function buildMockAxiosInstance(postImpl: jest.Mock) {
  return {
    post: postImpl,
    get: jest.fn(),
    defaults: { headers: {} },
    interceptors: { request: { use: jest.fn() }, response: { use: jest.fn() } },
  };
}

// ---------------------------------------------------------------------------
// Fixture factories
// ---------------------------------------------------------------------------

function makeOneDriveHit(overrides: Partial<{
  id: string;
  name: string;
  webUrl: string;
  size: number;
  lastModifiedDateTime: string;
  authorName: string;
  path: string;
  summary: string;
}> = {}) {
  const {
    id = 'item-001',
    name = 'Meeting Notes Q1.docx',
    webUrl = 'https://onedrive.live.com/item-001',
    size = 12345,
    lastModifiedDateTime = '2024-02-10T14:30:00Z',
    authorName = 'Alice Smith',
    path = '/drives/driveId123/root:/Documents/Meetings',
    summary = 'Discussed project timeline and deliverables.',
  } = overrides;

  return {
    hitId: id,
    rank: 1,
    summary,
    resource: {
      '@odata.type': '#microsoft.graph.driveItem',
      id,
      name,
      webUrl,
      size,
      lastModifiedDateTime,
      lastModifiedBy: { user: { displayName: authorName } },
      parentReference: {
        path,
        driveId: 'driveId123',
      },
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    },
  };
}

function makeSharePointHit(overrides: Partial<{
  id: string;
  name: string;
  webUrl: string;
  size: number;
  lastModifiedDateTime: string;
  authorName: string;
  path: string;
  siteId: string;
  summary: string;
}> = {}) {
  const {
    id = 'sp-item-001',
    name = 'Project Kickoff Agenda.pptx',
    webUrl = 'https://contoso.sharepoint.com/sites/eng/Shared Documents/Project Kickoff Agenda.pptx',
    size = 98765,
    lastModifiedDateTime = '2024-03-15T09:00:00Z',
    authorName = 'Bob Johnson',
    path = '/drives/spDriveId/root:/Shared Documents',
    siteId = 'contoso.sharepoint.com,site-001,web-001',
    summary = 'Kickoff agenda for the new quarter.',
  } = overrides;

  return {
    hitId: id,
    rank: 2,
    summary,
    resource: {
      '@odata.type': '#microsoft.graph.driveItem',
      id,
      name,
      webUrl,
      size,
      lastModifiedDateTime,
      lastModifiedBy: { user: { displayName: authorName } },
      parentReference: {
        path,
        siteId,
        driveId: 'spDriveId',
      },
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
    },
  };
}

function makeGraphResponse(hits: ReturnType<typeof makeOneDriveHit>[], total = hits.length, moreResultsAvailable = false): GraphSearchResponse {
  return {
    value: [
      {
        searchTerms: ['test'],
        hitsContainers: [
          {
            hits,
            total,
            moreResultsAvailable,
          },
        ],
      },
    ],
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('OnedriveSearchClient', () => {
  const ACCESS_TOKEN = 'mock-access-token';
  let postMock: jest.Mock;

  beforeEach(() => {
    jest.clearAllMocks();
    postMock = jest.fn();
    mockedAxios.create.mockReturnValue(buildMockAxiosInstance(postMock) as never);
  });

  // -------------------------------------------------------------------------
  // Basic search
  // -------------------------------------------------------------------------

  describe('searchFiles – basic search', () => {
    it('returns mapped OneDrive results', async () => {
      const hit = makeOneDriveHit();
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response: FileSearchResponse = await client.searchFiles({ query: 'meeting notes' });

      expect(response.results).toHaveLength(1);
      const [result] = response.results;
      expect(result.name).toBe('Meeting Notes Q1.docx');
      expect(result.fileType).toBe('docx');
      expect(result.source).toBe('onedrive');
      expect(result.author).toBe('Alice Smith');
      expect(result.lastModified).toBe('2024-02-10T14:30:00Z');
      expect(result.contentSnippet).toBe('Discussed project timeline and deliverables.');
      expect(result.webUrl).toBe('https://onedrive.live.com/item-001');
      expect(result.size).toBe(12345);
    });

    it('returns mapped SharePoint results with siteId', async () => {
      const hit = makeSharePointHit();
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'kickoff agenda' });

      expect(response.results).toHaveLength(1);
      const [result] = response.results;
      expect(result.source).toBe('sharepoint');
      expect(result.siteId).toBe('contoso.sharepoint.com,site-001,web-001');
      expect(result.fileType).toBe('pptx');
    });

    it('returns mixed OneDrive and SharePoint results', async () => {
      const hits = [makeOneDriveHit(), makeSharePointHit()];
      postMock.mockResolvedValueOnce({ data: makeGraphResponse(hits, 2) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'project' });

      expect(response.results).toHaveLength(2);
      expect(response.results.map((r) => r.source)).toEqual(
        expect.arrayContaining(['onedrive', 'sharepoint']),
      );
    });

    it('returns empty results when no hits', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([], 0) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'nonexistent' });

      expect(response.results).toHaveLength(0);
      expect(response.totalCount).toBe(0);
      expect(response.hasMore).toBe(false);
    });

    it('reports totalCount and hasMore from the Graph response', async () => {
      const hit = makeOneDriveHit();
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit], 150, true) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'report' });

      expect(response.totalCount).toBe(150);
      expect(response.hasMore).toBe(true);
    });
  });

  // -------------------------------------------------------------------------
  // File-type filtering
  // -------------------------------------------------------------------------

  describe('searchFiles – file-type filtering', () => {
    it('includes FileType clause in KQL when fileTypes is a subset', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'notes', fileTypes: ['docx', 'pdf'] });

      const requestBody = postMock.mock.calls[0][1];
      const kql: string = requestBody.requests[0].query.queryString;
      expect(kql).toContain('FileType:docx');
      expect(kql).toContain('FileType:pdf');
    });

    it('omits FileType clause when all supported types are requested', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      // Default call uses all supported types.
      await client.searchFiles({ query: 'notes' });

      const requestBody = postMock.mock.calls[0][1];
      const kql: string = requestBody.requests[0].query.queryString;
      expect(kql).not.toContain('FileType:');
    });

    it('supports all declared file types', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({
        query: 'report',
        fileTypes: ['docx', 'pptx', 'xlsx', 'pdf', 'txt', 'md'],
      });

      // When ALL types are listed explicitly the clause should be omitted
      // (same as omitting the fileTypes option).
      const requestBody = postMock.mock.calls[0][1];
      const kql: string = requestBody.requests[0].query.queryString;
      expect(kql).not.toContain('FileType:');
    });
  });

  // -------------------------------------------------------------------------
  // Date range filtering
  // -------------------------------------------------------------------------

  describe('searchFiles – date range filtering', () => {
    it('includes LastModifiedTime lower bound in KQL', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'notes', dateFrom: new Date('2024-01-01') });

      const kql: string = postMock.mock.calls[0][1].requests[0].query.queryString;
      expect(kql).toContain('LastModifiedTime>=2024-01-01');
    });

    it('includes LastModifiedTime upper bound in KQL', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'notes', dateTo: new Date('2024-12-31') });

      const kql: string = postMock.mock.calls[0][1].requests[0].query.queryString;
      expect(kql).toContain('LastModifiedTime<=2024-12-31');
    });

    it('includes both date bounds in KQL', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({
        query: 'quarterly review',
        dateFrom: new Date('2024-01-01'),
        dateTo: new Date('2024-03-31'),
      });

      const kql: string = postMock.mock.calls[0][1].requests[0].query.queryString;
      expect(kql).toContain('LastModifiedTime>=2024-01-01');
      expect(kql).toContain('LastModifiedTime<=2024-03-31');
    });
  });

  // -------------------------------------------------------------------------
  // maxResults option
  // -------------------------------------------------------------------------

  describe('searchFiles – maxResults', () => {
    it('defaults to 25 results', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'doc' });

      const size: number = postMock.mock.calls[0][1].requests[0].size;
      expect(size).toBe(25);
    });

    it('respects a custom maxResults value', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'doc', maxResults: 10 });

      const size: number = postMock.mock.calls[0][1].requests[0].size;
      expect(size).toBe(10);
    });

    it('caps maxResults at 100', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'doc', maxResults: 999 });

      const size: number = postMock.mock.calls[0][1].requests[0].size;
      expect(size).toBe(100);
    });
  });

  // -------------------------------------------------------------------------
  // Path extraction
  // -------------------------------------------------------------------------

  describe('searchFiles – path extraction', () => {
    it('extracts the path segment after root:', async () => {
      const hit = makeOneDriveHit({ path: '/drives/abc/root:/Documents/Meetings' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'meeting' });

      expect(response.results[0].path).toBe('/Documents/Meetings');
    });

    it('falls back to "/" when parentReference is missing', async () => {
      const hit = makeOneDriveHit();
      // Remove parentReference from the hit
      (hit.resource as Record<string, unknown>).parentReference = undefined;
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'meeting' });

      expect(response.results[0].path).toBe('/');
    });

    it('returns "/" when the path is exactly at root', async () => {
      const hit = makeOneDriveHit({ path: '/drives/abc/root:' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'file' });

      expect(response.results[0].path).toBe('/');
    });
  });

  // -------------------------------------------------------------------------
  // Meeting correlation
  // -------------------------------------------------------------------------

  describe('searchFiles – meeting correlation', () => {
    it('correlates file with high confidence when name contains meeting title', async () => {
      const hit = makeOneDriveHit({ name: 'Q1 Project Kickoff Notes.docx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({
        query: 'kickoff',
        meetingTitles: ['Project Kickoff'],
      });

      const correlation = response.results[0].meetingCorrelation;
      expect(correlation).toBeDefined();
      expect(correlation!.confidence).toBe('high');
      expect(correlation!.matchedMeetingTitle).toBe('Project Kickoff');
    });

    it('correlates file with medium confidence on partial word match', async () => {
      // "Quarterly Budget Review" does NOT appear as a contiguous substring in
      // the file name, but each individual word does — triggers medium confidence.
      const hit = makeOneDriveHit({ name: 'quarterly summary for budget review.xlsx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({
        query: 'quarterly',
        meetingTitles: ['Quarterly Budget Review'],
      });

      const correlation = response.results[0].meetingCorrelation;
      expect(correlation).toBeDefined();
      expect(correlation!.confidence).toBe('medium');
    });

    it('correlates file with low confidence via naming-convention patterns', async () => {
      const hit = makeOneDriveHit({ name: 'team meeting notes.docx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'team' });

      const correlation = response.results[0].meetingCorrelation;
      expect(correlation).toBeDefined();
      expect(correlation!.confidence).toBe('low');
    });

    it('detects agenda files as low-confidence meeting files', async () => {
      const hit = makeOneDriveHit({ name: 'Sprint Planning Agenda.docx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'sprint' });

      expect(response.results[0].meetingCorrelation?.confidence).toBe('low');
    });

    it('detects minutes files as low-confidence meeting files', async () => {
      const hit = makeOneDriveHit({ name: '2024-01-15 Board Minutes.pdf' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'board' });

      expect(response.results[0].meetingCorrelation?.confidence).toBe('low');
    });

    it('does not correlate an unrelated file', async () => {
      const hit = makeOneDriveHit({ name: 'Budget Forecast 2024.xlsx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'budget' });

      expect(response.results[0].meetingCorrelation).toBeUndefined();
    });

    it('ignores empty meeting title strings', async () => {
      const hit = makeOneDriveHit({ name: 'Budget Forecast 2024.xlsx' });
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({
        query: 'budget',
        meetingTitles: ['', '   '],
      });

      expect(response.results[0].meetingCorrelation).toBeUndefined();
    });
  });

  // -------------------------------------------------------------------------
  // Author fallback
  // -------------------------------------------------------------------------

  describe('searchFiles – author resolution', () => {
    it('falls back to createdBy when lastModifiedBy is absent', async () => {
      const hit = makeOneDriveHit({ authorName: '' });
      (hit.resource as Record<string, unknown>).lastModifiedBy = undefined;
      (hit.resource as Record<string, unknown>).createdBy = { user: { displayName: 'Carol White' } };
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'notes' });

      expect(response.results[0].author).toBe('Carol White');
    });

    it('returns empty string when neither author field is present', async () => {
      const hit = makeOneDriveHit({ authorName: '' });
      (hit.resource as Record<string, unknown>).lastModifiedBy = undefined;
      (hit.resource as Record<string, unknown>).createdBy = undefined;
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'notes' });

      expect(response.results[0].author).toBe('');
    });
  });

  // -------------------------------------------------------------------------
  // Supported file types
  // -------------------------------------------------------------------------

  describe('searchFiles – supported file types', () => {
    const fileTypes = ['docx', 'pptx', 'xlsx', 'pdf', 'txt', 'md'] as const;

    fileTypes.forEach((ext) => {
      it(`correctly identifies .${ext} file type`, async () => {
        const hit = makeOneDriveHit({ name: `document.${ext}`, id: `item-${ext}` });
        postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

        const client = new OnedriveSearchClient(ACCESS_TOKEN);
        const response = await client.searchFiles({ query: 'document', fileTypes: [ext] });

        expect(response.results[0].fileType).toBe(ext);
      });
    });
  });

  // -------------------------------------------------------------------------
  // Error propagation
  // -------------------------------------------------------------------------

  describe('searchFiles – error handling', () => {
    it('propagates non-retryable HTTP errors (e.g. 403)', async () => {
      const error = Object.assign(new Error('Forbidden'), {
        response: { status: 403, headers: {} },
        isAxiosError: true,
      });
      postMock.mockRejectedValueOnce(error);

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await expect(client.searchFiles({ query: 'secret' })).rejects.toThrow('Forbidden');
    });

    it('propagates non-retryable HTTP errors (e.g. 401)', async () => {
      const error = Object.assign(new Error('Unauthorized'), {
        response: { status: 401, headers: {} },
        isAxiosError: true,
      });
      postMock.mockRejectedValueOnce(error);

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await expect(client.searchFiles({ query: 'doc' })).rejects.toThrow('Unauthorized');
    });

    it('retries on HTTP 429 and eventually succeeds', async () => {
      jest.useFakeTimers();

      const rateLimitError = Object.assign(new Error('Too Many Requests'), {
        response: { status: 429, headers: { 'retry-after': '0' } },
        isAxiosError: true,
      });
      const successData = makeGraphResponse([makeOneDriveHit()]);

      postMock
        .mockRejectedValueOnce(rateLimitError)
        .mockResolvedValueOnce({ data: successData });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const searchPromise = client.searchFiles({ query: 'doc' });

      // Let the retry delay resolve.
      await jest.runAllTimersAsync();

      const response = await searchPromise;
      expect(response.results).toHaveLength(1);
      expect(postMock).toHaveBeenCalledTimes(2);

      jest.useRealTimers();
    });

    it('retries on HTTP 500 and eventually succeeds', async () => {
      jest.useFakeTimers();

      const serverError = Object.assign(new Error('Internal Server Error'), {
        response: { status: 500, headers: {} },
        isAxiosError: true,
      });
      const successData = makeGraphResponse([makeOneDriveHit()]);

      postMock
        .mockRejectedValueOnce(serverError)
        .mockResolvedValueOnce({ data: successData });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const searchPromise = client.searchFiles({ query: 'doc' });

      await jest.runAllTimersAsync();

      const response = await searchPromise;
      expect(response.results).toHaveLength(1);

      jest.useRealTimers();
    });

    it('throws after exhausting retries on persistent 500 errors', async () => {
      jest.useFakeTimers();

      const serverError = Object.assign(new Error('Internal Server Error'), {
        response: { status: 500, headers: {} },
        isAxiosError: true,
      });

      postMock.mockRejectedValue(serverError);

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      // Attach the rejection handler immediately to avoid unhandled-rejection warnings.
      const rejectionPromise = expect(
        client.searchFiles({ query: 'doc' }),
      ).rejects.toThrow('Internal Server Error');

      // Advance all retry delays.
      await jest.runAllTimersAsync();

      await rejectionPromise;
      // 1 original + 3 retries = 4 total attempts
      expect(postMock).toHaveBeenCalledTimes(4);

      jest.useRealTimers();
    });
  });

  // -------------------------------------------------------------------------
  // Graph request structure
  // -------------------------------------------------------------------------

  describe('searchFiles – Graph API request structure', () => {
    it('sends a POST to /search/query with the correct body shape', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'kickoff' });

      expect(postMock).toHaveBeenCalledWith(
        '/search/query',
        expect.objectContaining({
          requests: expect.arrayContaining([
            expect.objectContaining({
              entityTypes: ['driveItem'],
              query: expect.objectContaining({ queryString: expect.any(String) }),
            }),
          ]),
        }),
      );
    });

    it('includes required fields in Graph request', async () => {
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      await client.searchFiles({ query: 'doc' });

      const fields: string[] = postMock.mock.calls[0][1].requests[0].fields;
      expect(fields).toContain('name');
      expect(fields).toContain('webUrl');
      expect(fields).toContain('lastModifiedDateTime');
      expect(fields).toContain('createdBy');
      expect(fields).toContain('parentReference');
    });

    it('creates axios instance with correct base URL and auth header', () => {
      new OnedriveSearchClient(ACCESS_TOKEN);

      expect(mockedAxios.create).toHaveBeenCalledWith(
        expect.objectContaining({
          baseURL: 'https://graph.microsoft.com/v1.0',
          headers: expect.objectContaining({
            Authorization: `Bearer ${ACCESS_TOKEN}`,
          }),
        }),
      );
    });
  });

  // -------------------------------------------------------------------------
  // Handles missing optional fields in the Graph response
  // -------------------------------------------------------------------------

  describe('searchFiles – graceful handling of missing fields', () => {
    it('handles missing summary (empty contentSnippet)', async () => {
      const hit = makeOneDriveHit({ summary: '' });
      (hit as Record<string, unknown>).summary = undefined;
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'notes' });

      expect(response.results[0].contentSnippet).toBe('');
    });

    it('handles missing webUrl (empty string)', async () => {
      const hit = makeOneDriveHit({ webUrl: '' });
      (hit.resource as Record<string, unknown>).webUrl = undefined;
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'notes' });

      expect(response.results[0].webUrl).toBe('');
    });

    it('falls back to fileSystemInfo for lastModified when lastModifiedDateTime is absent', async () => {
      const hit = makeOneDriveHit();
      (hit.resource as Record<string, unknown>).lastModifiedDateTime = undefined;
      (hit.resource as Record<string, unknown>).fileSystemInfo = { lastModifiedDateTime: '2024-05-01T00:00:00Z' };
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([hit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'notes' });

      expect(response.results[0].lastModified).toBe('2024-05-01T00:00:00Z');
    });

    it('skips hits that are not driveItem resources', async () => {
      const nonDriveHit = {
        hitId: 'event-001',
        rank: 1,
        summary: '',
        resource: {
          '@odata.type': '#microsoft.graph.event',
          id: 'event-001',
          name: 'Calendar Event',
        },
      };
      // @ts-expect-error: intentionally using a non-driveItem type for testing
      postMock.mockResolvedValueOnce({ data: makeGraphResponse([nonDriveHit]) });

      const client = new OnedriveSearchClient(ACCESS_TOKEN);
      const response = await client.searchFiles({ query: 'event' });

      expect(response.results).toHaveLength(0);
    });
  });
});
