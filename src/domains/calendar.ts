/**
 * Calendar domain search module.
 *
 * In a production system this would call the Microsoft Graph API (calendar
 * events endpoint).  Here we expose a thin adapter interface so the
 * orchestrator can call it uniformly alongside the other domain modules.
 */

import { CalendarEvent, SearchRequest } from '../types';

/**
 * Search the calendar / meeting-metadata domain.
 *
 * @param request  Unified search criteria.
 * @param fetch    Optional injectable fetch function (defaults to a no-op
 *                 stub that returns an empty array; replace with a real
 *                 Graph API client in production).
 * @returns        Matching calendar events.
 */
export async function searchCalendar(
  request: SearchRequest,
  fetch: (req: SearchRequest) => Promise<CalendarEvent[]> = _defaultFetch,
): Promise<CalendarEvent[]> {
  return fetch(request);
}

/** Default stub – returns no results until wired to a real data source. */
async function _defaultFetch(_request: SearchRequest): Promise<CalendarEvent[]> {
  return [];
}
