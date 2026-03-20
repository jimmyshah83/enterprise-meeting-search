import { TranscriptResult, SearchRequest } from '../types';

/**
 * Searches the transcripts domain for results matching the request criteria.
 * In production this would call the Microsoft Graph transcripts API.
 */
export async function searchTranscripts(
  request: SearchRequest,
  adapter?: (req: SearchRequest) => Promise<TranscriptResult[]>,
): Promise<TranscriptResult[]> {
  if (adapter) {
    return adapter(request);
  }
  return [];
}
