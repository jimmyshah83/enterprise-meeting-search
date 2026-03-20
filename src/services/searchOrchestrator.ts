import { AggregatedSearchResult, DomainAdapter, DomainSearchResult, SearchParams } from '../types';
import { CacheService } from './cacheService';

export interface OrchestratorOptions {
  domainTimeoutMs: number;
  overallTimeoutMs: number;
}

const DEFAULT_OPTIONS: OrchestratorOptions = {
  domainTimeoutMs: 10000,
  overallTimeoutMs: 30000,
};

export class SearchOrchestrator {
  private readonly domains: DomainAdapter[];
  private readonly cacheService?: CacheService;
  private readonly options: OrchestratorOptions;

  constructor(domains: DomainAdapter[], cacheService?: CacheService, options: Partial<OrchestratorOptions> = {}) {
    this.domains = domains;
    this.cacheService = cacheService;
    this.options = { ...DEFAULT_OPTIONS, ...options };
  }

  async search(params: SearchParams): Promise<AggregatedSearchResult> {
    const startTime = Date.now();

    const domainSearches = this.domains.map(adapter => this.searchDomain(adapter, params));
    const domainResults = await Promise.all(domainSearches);

    const totalResults = domainResults.reduce((sum, r) => sum + r.results.length, 0);

    return {
      domainResults,
      totalResults,
      searchParams: params,
      searchedAt: new Date(),
      durationMs: Date.now() - startTime,
    };
  }

  private async searchDomain(adapter: DomainAdapter, params: SearchParams): Promise<DomainSearchResult> {
    if (this.cacheService) {
      const cached = await this.cacheService.getCachedDomainResult(params, adapter.domain);
      if (cached) {
        return cached;
      }
    }

    try {
      const results = await this.withTimeout(adapter.search(params), this.options.domainTimeoutMs, adapter.domain);
      const domainResult: DomainSearchResult = {
        domain: adapter.domain,
        results,
        cached: false,
        searchedAt: new Date(),
      };

      if (this.cacheService) {
        await this.cacheService.setCachedDomainResult(params, adapter.domain, domainResult);
      }

      return domainResult;
    } catch (error) {
      console.error(`[SearchOrchestrator] Error searching domain=${adapter.domain}:`, error);
      return {
        domain: adapter.domain,
        results: [],
        cached: false,
        searchedAt: new Date(),
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  private async withTimeout<T>(promise: Promise<T>, ms: number, domain: string): Promise<T> {
    const timeout = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error(`Domain ${domain} timed out after ${ms}ms`)), ms)
    );
    return Promise.race([promise, timeout]);
  }
}
