import { FileResult, SearchRequest } from '../types';

/**
 * Searches the files domain for documents matching the request criteria.
 * In production this would call the Microsoft Graph OneDrive/SharePoint API.
 */
export async function searchFiles(
  request: SearchRequest,
  adapter?: (req: SearchRequest) => Promise<FileResult[]>,
): Promise<FileResult[]> {
  if (adapter) {
    return adapter(request);
  }
  return [];
}
