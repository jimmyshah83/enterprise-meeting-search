/**
 * Teams chat domain search module.
 *
 * In a production system this would call the Microsoft Graph API
 * (chats / messages endpoint).
 */

import { ChatMessage, SearchRequest } from '../types';

/**
 * Search the Teams chat domain.
 *
 * @param request  Unified search criteria.
 * @param fetch    Optional injectable fetch function.
 * @returns        Matching chat messages.
 */
export async function searchChats(
  request: SearchRequest,
  fetch: (req: SearchRequest) => Promise<ChatMessage[]> = _defaultFetch,
): Promise<ChatMessage[]> {
  return fetch(request);
}

/** Default stub – returns no results until wired to a real data source. */
async function _defaultFetch(_request: SearchRequest): Promise<ChatMessage[]> {
  return [];
}
