import { EmailResult, SearchRequest } from '../types';

/**
 * Searches the email domain for messages matching the request criteria.
 * In production this would call the Microsoft Graph Mail API.
 */
export async function searchEmail(
  request: SearchRequest,
  adapter?: (req: SearchRequest) => Promise<EmailResult[]>,
): Promise<EmailResult[]> {
  if (adapter) {
    return adapter(request);
  }
  return [];
}
