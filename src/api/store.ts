/**
 * In-memory stores for async search and reconstruction operations.
 * In production these would be backed by a database or cache (e.g. Redis).
 */

import { AsyncSearchRecord, AsyncReconstructionRecord, UnifiedMeetingResult } from '../types';

class SearchStore {
  private readonly records = new Map<string, AsyncSearchRecord>();

  set(record: AsyncSearchRecord): void {
    this.records.set(record.searchId, record);
  }

  get(searchId: string): AsyncSearchRecord | undefined {
    return this.records.get(searchId);
  }

  update(searchId: string, updates: Partial<AsyncSearchRecord>): void {
    const existing = this.records.get(searchId);
    if (existing) {
      this.records.set(searchId, { ...existing, ...updates });
    }
  }

  /** Finds the first UnifiedMeetingResult with the given meetingId across all completed searches. */
  findMeetingResult(meetingId: string): UnifiedMeetingResult | undefined {
    for (const record of this.records.values()) {
      if (record.status === 'completed' && record.response) {
        const match = record.response.results.find((r) => r.meetingId === meetingId);
        if (match) return match;
      }
    }
    return undefined;
  }
}

class ReconstructionStore {
  private readonly records = new Map<string, AsyncReconstructionRecord>();

  set(record: AsyncReconstructionRecord): void {
    this.records.set(record.reconstructionId, record);
  }

  get(reconstructionId: string): AsyncReconstructionRecord | undefined {
    return this.records.get(reconstructionId);
  }

  update(reconstructionId: string, updates: Partial<AsyncReconstructionRecord>): void {
    const existing = this.records.get(reconstructionId);
    if (existing) {
      this.records.set(reconstructionId, { ...existing, ...updates });
    }
  }

  getByMeetingId(meetingId: string): AsyncReconstructionRecord | undefined {
    for (const record of this.records.values()) {
      if (record.meetingId === meetingId) return record;
    }
    return undefined;
  }
}

export const searchStore = new SearchStore();
export const reconstructionStore = new ReconstructionStore();
