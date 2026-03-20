/**
 * OneDrive / SharePoint file domain search module.
 *
 * In a production system this would call the Microsoft Graph API
 * (search / drive items endpoint).
 */

import { FileResult, SearchRequest } from '../types';

/**
 * Search the file domain (OneDrive / SharePoint).
 *
 * @param request  Unified search criteria.
 * @param fetch    Optional injectable fetch function.
 * @returns        Matching file results.
 */
export async function searchFiles(
  request: SearchRequest,
  fetch: (req: SearchRequest) => Promise<FileResult[]> = _defaultFetch,
): Promise<FileResult[]> {
  return fetch(request);
}

/** Default stub – returns no results until wired to a real data source. */
async function _defaultFetch(_request: SearchRequest): Promise<FileResult[]> {
  return [];
}
