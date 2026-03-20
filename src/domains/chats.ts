import { ChatMessage, SearchRequest } from '../types';

/**
 * Searches the chats domain for messages matching the request criteria.
 * In production this would call the Microsoft Graph Teams chats API.
 */
export async function searchChats(
  request: SearchRequest,
  adapter?: (req: SearchRequest) => Promise<ChatMessage[]>,
): Promise<ChatMessage[]> {
  if (adapter) {
    return adapter(request);
  }
  return [];
}
