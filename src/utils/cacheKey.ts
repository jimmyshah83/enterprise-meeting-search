import * as crypto from 'crypto';
import { DomainType, SearchParams } from '../types';

export interface NormalizedSearchParams {
  keywords: string[];
  startDate?: string;
  endDate?: string;
  user?: string;
  attendees?: string[];
  meetingType?: string;
}

export function normalizeSearchParams(params: SearchParams): NormalizedSearchParams {
  return {
    keywords: [...params.keywords].map(k => k.toLowerCase().trim()).sort(),
    startDate: params.dateRange?.start.toISOString(),
    endDate: params.dateRange?.end.toISOString(),
    user: params.user?.toLowerCase().trim(),
    attendees: params.attendees ? [...params.attendees].map(a => a.toLowerCase().trim()).sort() : undefined,
    meetingType: params.meetingType,
  };
}

export function generateCacheKey(params: SearchParams, domain: DomainType, keyPrefix: string): string {
  const normalized = normalizeSearchParams(params);
  const hash = crypto.createHash('sha256').update(JSON.stringify(normalized)).digest('hex').substring(0, 16);
  return `${keyPrefix}:${domain}:${hash}`;
}

export function generateSearchPatternKey(params: SearchParams, keyPrefix: string): string {
  const normalized = normalizeSearchParams(params);
  const hash = crypto.createHash('sha256').update(JSON.stringify(normalized)).digest('hex').substring(0, 16);
  return `${keyPrefix}:*:${hash}`;
}
