import { SearchOrchestrator } from '../src/services/searchOrchestrator';
import { CacheService, RedisClient } from '../src/services/cacheService';
import { DomainAdapter, MeetingResult, SearchParams } from '../src/types';

function createMockRedis(): jest.Mocked<RedisClient> & { store: Map<string, { value: string; expiresAt: number }> } {
  const store = new Map<string, { value: string; expiresAt: number }>();

  return {
    store,
    get: jest.fn(async (key: string) => {
      const entry = store.get(key);
      if (!entry) return null;
      if (Date.now() > entry.expiresAt) {
        store.delete(key);
        return null;
      }
      return entry.value;
    }),
    set: jest.fn(async (key: string, value: string, _expiryMode: string, time: number) => {
      store.set(key, { value, expiresAt: Date.now() + time * 1000 });
      return 'OK';
    }),
    del: jest.fn(async (...keys: string[]) => {
      let count = 0;
      for (const key of keys) {
        if (store.delete(key)) count++;
      }
      return count;
    }),
    keys: jest.fn(async (pattern: string) => {
      const regexStr = pattern.replace(/\*/g, '.*').replace(/\?/g, '[^:]');
      const regex = new RegExp(`^${regexStr}$`);
      return Array.from(store.keys()).filter(k => regex.test(k));
    }),
  };
}

function createMockAdapter(domain: 'calendar' | 'chat' | 'email' | 'transcript' | 'file', results: MeetingResult[]): DomainAdapter {
  return {
    domain,
    search: jest.fn(async () => results),
  };
}

describe('SearchOrchestrator', () => {
  const searchParams: SearchParams = {
    keywords: ['PTU'],
    dateRange: {
      start: new Date('2024-02-01T00:00:00Z'),
      end: new Date('2024-02-28T23:59:59Z'),
    },
    user: 'user@example.com',
  };

  const calendarMeeting: MeetingResult = {
    id: 'cal-1',
    title: 'PTU Review',
    dateTime: new Date('2024-02-15T10:00:00Z'),
    domain: 'calendar',
  };

  const chatMessage: MeetingResult = {
    id: 'chat-1',
    title: 'PTU Chat',
    dateTime: new Date('2024-02-15T10:30:00Z'),
    domain: 'chat',
  };

  describe('search without cache', () => {
    it('searches all domains and aggregates results', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const chatAdapter = createMockAdapter('chat', [chatMessage]);
      const orchestrator = new SearchOrchestrator([calendarAdapter, chatAdapter]);

      const result = await orchestrator.search(searchParams);

      expect(result.domainResults).toHaveLength(2);
      expect(result.totalResults).toBe(2);
      expect(calendarAdapter.search).toHaveBeenCalledWith(searchParams);
      expect(chatAdapter.search).toHaveBeenCalledWith(searchParams);
    });

    it('returns correct aggregated search metadata', async () => {
      const adapter = createMockAdapter('calendar', [calendarMeeting]);
      const orchestrator = new SearchOrchestrator([adapter]);

      const result = await orchestrator.search(searchParams);

      expect(result.searchParams).toBe(searchParams);
      expect(result.searchedAt).toBeInstanceOf(Date);
      expect(result.durationMs).toBeGreaterThanOrEqual(0);
    });

    it('handles domain errors gracefully', async () => {
      const failingAdapter: DomainAdapter = {
        domain: 'calendar',
        search: jest.fn(async () => { throw new Error('Graph API error'); }),
      };
      const orchestrator = new SearchOrchestrator([failingAdapter]);

      const result = await orchestrator.search(searchParams);

      expect(result.domainResults[0].error).toBe('Graph API error');
      expect(result.domainResults[0].results).toHaveLength(0);
      expect(result.totalResults).toBe(0);
    });

    it('marks all results as not cached when no cache service', async () => {
      const adapter = createMockAdapter('calendar', [calendarMeeting]);
      const orchestrator = new SearchOrchestrator([adapter]);

      const result = await orchestrator.search(searchParams);

      expect(result.domainResults[0].cached).toBe(false);
    });
  });

  describe('search with cache', () => {
    it('caches domain results after first search', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });
      const orchestrator = new SearchOrchestrator([calendarAdapter], cacheService);

      await orchestrator.search(searchParams);

      expect(redis.set).toHaveBeenCalledWith(
        expect.stringContaining('test:calendar:'),
        expect.any(String),
        'EX',
        900
      );
    });

    it('returns cached results on second search without calling adapter', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });
      const orchestrator = new SearchOrchestrator([calendarAdapter], cacheService);

      await orchestrator.search(searchParams);  // First call - populates cache
      await orchestrator.search(searchParams);  // Second call - should use cache

      expect(calendarAdapter.search).toHaveBeenCalledTimes(1);
    });

    it('marks cached domain results with cached=true', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });
      const orchestrator = new SearchOrchestrator([calendarAdapter], cacheService);

      await orchestrator.search(searchParams);
      const secondResult = await orchestrator.search(searchParams);

      expect(secondResult.domainResults[0].cached).toBe(true);
    });

    it('performs partial cache hit when only some domains are cached', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const chatAdapter = createMockAdapter('chat', [chatMessage]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });

      // Only cache calendar results manually
      await cacheService.setCachedDomainResult(searchParams, 'calendar', {
        domain: 'calendar',
        results: [calendarMeeting],
        cached: false,
        searchedAt: new Date(),
      });

      const orchestrator = new SearchOrchestrator([calendarAdapter, chatAdapter], cacheService);
      const result = await orchestrator.search(searchParams);

      expect(calendarAdapter.search).not.toHaveBeenCalled(); // calendar was cached
      expect(chatAdapter.search).toHaveBeenCalledTimes(1);   // chat was not cached
      expect(result.domainResults[0].cached).toBe(true);
      expect(result.domainResults[1].cached).toBe(false);
    });

    it('tracks cache hit/miss metrics correctly', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });
      const orchestrator = new SearchOrchestrator([calendarAdapter], cacheService);

      await orchestrator.search(searchParams); // miss
      await orchestrator.search(searchParams); // hit

      const metrics = cacheService.getMetrics();
      expect(metrics.hits).toBe(1);
      expect(metrics.misses).toBe(1);
      expect(metrics.hitRate).toBeCloseTo(0.5);
    });

    it('allows cache invalidation between searches', async () => {
      const calendarAdapter = createMockAdapter('calendar', [calendarMeeting]);
      const redis = createMockRedis();
      const cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test' });
      const orchestrator = new SearchOrchestrator([calendarAdapter], cacheService);

      await orchestrator.search(searchParams);   // First call - caches
      await cacheService.invalidateSearch(searchParams); // Invalidate
      await orchestrator.search(searchParams);   // Should call adapter again

      expect(calendarAdapter.search).toHaveBeenCalledTimes(2);
    });
  });
});
