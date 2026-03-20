/**
 * Email domain search module.
 *
 * In a production system this would call the Microsoft Graph API
 * (messages / mail endpoint).
 */

import { EmailResult, SearchRequest } from '../types';

/**
 * Search the email domain.
 *
 * @param request  Unified search criteria.
 * @param fetch    Optional injectable fetch function.
 * @returns        Matching email results.
 */
export async function searchEmail(
  request: SearchRequest,
  fetch: (req: SearchRequest) => Promise<EmailResult[]> = _defaultFetch,
): Promise<EmailResult[]> {
  return fetch(request);
}

/** Default stub – returns no results until wired to a real data source. */
async function _defaultFetch(_request: SearchRequest): Promise<EmailResult[]> {
  return [];
}
