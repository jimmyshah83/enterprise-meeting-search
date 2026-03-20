import { CacheService, RedisClient } from '../src/services/cacheService';
import { DomainSearchResult, SearchParams } from '../src/types';

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

describe('CacheService', () => {
  let redis: ReturnType<typeof createMockRedis>;
  let cacheService: CacheService;

  const searchParams: SearchParams = {
    keywords: ['PTU', 'WSIB'],
    dateRange: {
      start: new Date('2024-02-01T00:00:00Z'),
      end: new Date('2024-02-28T23:59:59Z'),
    },
    user: 'user@example.com',
  };

  const mockDomainResult: DomainSearchResult = {
    domain: 'calendar',
    results: [
      {
        id: 'meeting-123',
        title: 'PTU WSIB Review',
        dateTime: new Date('2024-02-15T10:00:00Z'),
        organizer: 'organizer@example.com',
        domain: 'calendar',
      },
    ],
    cached: false,
    searchedAt: new Date('2024-02-15T09:00:00Z'),
  };

  beforeEach(() => {
    redis = createMockRedis();
    cacheService = new CacheService(redis, { ttlSeconds: 900, keyPrefix: 'test-cache' });
  });

  describe('getCachedDomainResult', () => {
    it('returns null on cache miss', async () => {
      const result = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      expect(result).toBeNull();
    });

    it('returns cached result on cache hit', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      const result = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      expect(result).not.toBeNull();
      expect(result!.domain).toBe('calendar');
      expect(result!.results).toHaveLength(1);
      expect(result!.results[0].title).toBe('PTU WSIB Review');
    });

    it('marks cached results with cached=true', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      const result = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      expect(result!.cached).toBe(true);
    });

    it('deserializes date fields correctly', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      const result = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      expect(result!.searchedAt).toBeInstanceOf(Date);
      expect(result!.results[0].dateTime).toBeInstanceOf(Date);
    });

    it('returns null for expired entries', async () => {
      const shortTtlService = new CacheService(redis, { ttlSeconds: 1, keyPrefix: 'test-cache' });
      await shortTtlService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);

      // Manually expire the entry
      const entries = Array.from(redis.store.entries());
      for (const [key, entry] of entries) {
        redis.store.set(key, { ...entry, expiresAt: Date.now() - 1 });
      }

      const result = await shortTtlService.getCachedDomainResult(searchParams, 'calendar');
      expect(result).toBeNull();
    });

    it('caches different domains separately', async () => {
      const chatResult: DomainSearchResult = { ...mockDomainResult, domain: 'chat' };
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.setCachedDomainResult(searchParams, 'chat', chatResult);

      const calResult = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      const chatCached = await cacheService.getCachedDomainResult(searchParams, 'chat');

      expect(calResult!.domain).toBe('calendar');
      expect(chatCached!.domain).toBe('chat');
    });
  });

  describe('setCachedDomainResult', () => {
    it('stores result in Redis with correct TTL', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      expect(redis.set).toHaveBeenCalledWith(
        expect.stringContaining('test-cache:calendar:'),
        expect.any(String),
        'EX',
        900
      );
    });

    it('stores result with cached=false in Redis', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      const call = (redis.set as jest.Mock).mock.calls[0];
      const stored = JSON.parse(call[1]);
      expect(stored.cached).toBe(false);
    });
  });

  describe('invalidateSearch', () => {
    it('removes all domain cache entries for the given search params', async () => {
      const chatResult: DomainSearchResult = { ...mockDomainResult, domain: 'chat' };
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.setCachedDomainResult(searchParams, 'chat', chatResult);

      await cacheService.invalidateSearch(searchParams);

      const calResult = await cacheService.getCachedDomainResult(searchParams, 'calendar');
      const chatCached = await cacheService.getCachedDomainResult(searchParams, 'chat');
      expect(calResult).toBeNull();
      expect(chatCached).toBeNull();
    });

    it('does not remove cache entries for different search params', async () => {
      const otherParams: SearchParams = { keywords: ['different'] };
      const otherResult: DomainSearchResult = { ...mockDomainResult };
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.setCachedDomainResult(otherParams, 'calendar', otherResult);

      await cacheService.invalidateSearch(searchParams);

      const otherCached = await cacheService.getCachedDomainResult(otherParams, 'calendar');
      expect(otherCached).not.toBeNull();
    });
  });

  describe('invalidateAll', () => {
    it('removes all cache entries', async () => {
      const otherParams: SearchParams = { keywords: ['different'] };
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.setCachedDomainResult(otherParams, 'chat', { ...mockDomainResult, domain: 'chat' });

      await cacheService.invalidateAll();

      expect(redis.store.size).toBe(0);
    });
  });

  describe('metrics', () => {
    it('starts with zero metrics', () => {
      const metrics = cacheService.getMetrics();
      expect(metrics.hits).toBe(0);
      expect(metrics.misses).toBe(0);
      expect(metrics.hitRate).toBe(0);
      expect(metrics.invalidations).toBe(0);
    });

    it('increments misses on cache miss', async () => {
      await cacheService.getCachedDomainResult(searchParams, 'calendar');
      const metrics = cacheService.getMetrics();
      expect(metrics.misses).toBe(1);
      expect(metrics.hits).toBe(0);
    });

    it('increments hits on cache hit', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.getCachedDomainResult(searchParams, 'calendar');
      const metrics = cacheService.getMetrics();
      expect(metrics.hits).toBe(1);
      expect(metrics.misses).toBe(0);
    });

    it('calculates hit rate correctly', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.getCachedDomainResult(searchParams, 'calendar'); // hit
      await cacheService.getCachedDomainResult(searchParams, 'chat');    // miss
      await cacheService.getCachedDomainResult(searchParams, 'calendar'); // hit
      const metrics = cacheService.getMetrics();
      expect(metrics.hits).toBe(2);
      expect(metrics.misses).toBe(1);
      expect(metrics.hitRate).toBeCloseTo(2 / 3);
    });

    it('increments invalidations on invalidateSearch', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.invalidateSearch(searchParams);
      const metrics = cacheService.getMetrics();
      expect(metrics.invalidations).toBe(1);
    });

    it('increments invalidations on invalidateAll', async () => {
      await cacheService.setCachedDomainResult(searchParams, 'calendar', mockDomainResult);
      await cacheService.invalidateAll();
      const metrics = cacheService.getMetrics();
      expect(metrics.invalidations).toBe(1);
    });

    it('can be reset', () => {
      // Simulate some metrics
      cacheService.resetMetrics();
      const metrics = cacheService.getMetrics();
      expect(metrics.hits).toBe(0);
      expect(metrics.misses).toBe(0);
      expect(metrics.invalidations).toBe(0);
    });
  });
});
