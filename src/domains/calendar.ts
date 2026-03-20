import { CalendarEvent, SearchRequest } from '../types';

/**
 * Searches the calendar domain for events matching the request criteria.
 * In production this would call the Microsoft Graph Calendar API.
 */
export async function searchCalendar(
  request: SearchRequest,
  adapter?: (req: SearchRequest) => Promise<CalendarEvent[]>,
): Promise<CalendarEvent[]> {
  if (adapter) {
    return adapter(request);
  }
  return [];
}
