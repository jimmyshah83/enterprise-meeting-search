export type DomainType = 'calendar' | 'transcript' | 'chat' | 'email' | 'file';

export interface DateRange {
  start: Date;
  end: Date;
}

export interface SearchParams {
  keywords: string[];
  dateRange?: DateRange;
  user?: string;
  attendees?: string[];
  meetingType?: 'teams' | 'external' | 'channel';
}

export interface MeetingResult {
  id: string;
  title: string;
  dateTime: Date;
  organizer?: string;
  attendees?: string[];
  domain: DomainType;
  snippet?: string;
  metadata?: Record<string, unknown>;
}

export interface DomainSearchResult {
  domain: DomainType;
  results: MeetingResult[];
  cached: boolean;
  searchedAt: Date;
  error?: string;
}

export interface AggregatedSearchResult {
  domainResults: DomainSearchResult[];
  totalResults: number;
  searchParams: SearchParams;
  searchedAt: Date;
  durationMs: number;
}

export interface CacheOptions {
  ttlSeconds: number;
  keyPrefix: string;
}

export interface CacheMetrics {
  hits: number;
  misses: number;
  hitRate: number;
  invalidations: number;
}

export interface DomainAdapter {
  readonly domain: DomainType;
  search(params: SearchParams): Promise<MeetingResult[]>;
}
