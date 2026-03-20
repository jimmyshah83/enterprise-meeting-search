/**
 * Transcript domain search module.
 *
 * In a production system this would call the Microsoft Graph API
 * (onlineMeetings/transcripts) or a dedicated search index for meeting
 * transcripts.
 */

import { TranscriptResult, SearchRequest } from '../types';

/**
 * Search the meeting-transcript domain.
 *
 * @param request  Unified search criteria.
 * @param fetch    Optional injectable fetch function.
 * @returns        Matching transcript results.
 */
export async function searchTranscripts(
  request: SearchRequest,
  fetch: (req: SearchRequest) => Promise<TranscriptResult[]> = _defaultFetch,
): Promise<TranscriptResult[]> {
  return fetch(request);
}

/** Default stub – returns no results until wired to a real data source. */
async function _defaultFetch(_request: SearchRequest): Promise<TranscriptResult[]> {
  return [];
}
