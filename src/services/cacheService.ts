import { CacheMetrics, CacheOptions, DomainSearchResult, DomainType, SearchParams } from '../types';
import { generateCacheKey, generateSearchPatternKey } from '../utils/cacheKey';

export interface RedisClient {
  get(key: string): Promise<string | null>;
  set(key: string, value: string, expiryMode: string, time: number): Promise<unknown>;
  del(...keys: string[]): Promise<number>;
  keys(pattern: string): Promise<string[]>;
}

const DEFAULT_CACHE_OPTIONS: CacheOptions = {
  ttlSeconds: 900, // 15 minutes
  keyPrefix: 'enterprise-meeting-search:cache',
};

export class CacheService {
  private readonly redis: RedisClient;
  private readonly options: CacheOptions;
  private metrics: CacheMetrics = { hits: 0, misses: 0, hitRate: 0, invalidations: 0 };

  constructor(redis: RedisClient, options: Partial<CacheOptions> = {}) {
    this.redis = redis;
    this.options = { ...DEFAULT_CACHE_OPTIONS, ...options };
  }

  async getCachedDomainResult(params: SearchParams, domain: DomainType): Promise<DomainSearchResult | null> {
    const key = generateCacheKey(params, domain, this.options.keyPrefix);
    const cached = await this.redis.get(key);

    if (cached !== null) {
      this.metrics.hits++;
      this.updateHitRate();
      console.log(`[CacheService] Cache HIT for domain=${domain}, key=${key}`);
      const result = JSON.parse(cached) as DomainSearchResult;
      result.cached = true;
      result.searchedAt = new Date(result.searchedAt);
      result.results = result.results.map(r => ({ ...r, dateTime: new Date(r.dateTime) }));
      return result;
    }

    this.metrics.misses++;
    this.updateHitRate();
    console.log(`[CacheService] Cache MISS for domain=${domain}, key=${key}`);
    return null;
  }

  async setCachedDomainResult(params: SearchParams, domain: DomainType, result: DomainSearchResult): Promise<void> {
    const key = generateCacheKey(params, domain, this.options.keyPrefix);
    const value = JSON.stringify({ ...result, cached: false });
    await this.redis.set(key, value, 'EX', this.options.ttlSeconds);
    console.log(`[CacheService] Cached domain=${domain}, key=${key}, ttl=${this.options.ttlSeconds}s`);
  }

  async invalidateSearch(params: SearchParams): Promise<void> {
    const pattern = generateSearchPatternKey(params, this.options.keyPrefix);
    const keys = await this.redis.keys(pattern);
    if (keys.length > 0) {
      await this.redis.del(...keys);
      this.metrics.invalidations++;
      console.log(`[CacheService] Invalidated ${keys.length} cache entries for pattern=${pattern}`);
    }
  }

  async invalidateAll(): Promise<void> {
    const pattern = `${this.options.keyPrefix}:*`;
    const keys = await this.redis.keys(pattern);
    if (keys.length > 0) {
      await this.redis.del(...keys);
      this.metrics.invalidations++;
      console.log(`[CacheService] Invalidated all ${keys.length} cache entries`);
    }
  }

  getMetrics(): CacheMetrics {
    return { ...this.metrics };
  }

  resetMetrics(): void {
    this.metrics = { hits: 0, misses: 0, hitRate: 0, invalidations: 0 };
  }

  private updateHitRate(): void {
    const total = this.metrics.hits + this.metrics.misses;
    this.metrics.hitRate = total > 0 ? this.metrics.hits / total : 0;
  }
}
