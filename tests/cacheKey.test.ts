import { generateCacheKey, generateSearchPatternKey, normalizeSearchParams } from '../src/utils/cacheKey';
import { SearchParams } from '../src/types';

describe('cacheKey utilities', () => {
  const baseParams: SearchParams = {
    keywords: ['PTU', 'WSIB'],
    dateRange: {
      start: new Date('2024-02-01T00:00:00Z'),
      end: new Date('2024-02-28T23:59:59Z'),
    },
    user: 'user@example.com',
  };

  describe('normalizeSearchParams', () => {
    it('sorts keywords alphabetically', () => {
      const params: SearchParams = { keywords: ['Zebra', 'apple', 'mango'] };
      const normalized = normalizeSearchParams(params);
      expect(normalized.keywords).toEqual(['apple', 'mango', 'zebra']);
    });

    it('lowercases and trims keywords', () => {
      const params: SearchParams = { keywords: ['  PTU  ', 'WSIB'] };
      const normalized = normalizeSearchParams(params);
      expect(normalized.keywords).toEqual(['ptu', 'wsib']);
    });

    it('converts dateRange to ISO strings', () => {
      const normalized = normalizeSearchParams(baseParams);
      expect(normalized.startDate).toBe('2024-02-01T00:00:00.000Z');
      expect(normalized.endDate).toBe('2024-02-28T23:59:59.000Z');
    });

    it('lowercases user email', () => {
      const params: SearchParams = { keywords: [], user: 'User@Example.COM' };
      const normalized = normalizeSearchParams(params);
      expect(normalized.user).toBe('user@example.com');
    });

    it('sorts and lowercases attendees', () => {
      const params: SearchParams = {
        keywords: [],
        attendees: ['Bob@Example.COM', 'alice@example.com'],
      };
      const normalized = normalizeSearchParams(params);
      expect(normalized.attendees).toEqual(['alice@example.com', 'bob@example.com']);
    });

    it('produces identical normalized output for equivalent params regardless of input order', () => {
      const params1: SearchParams = { keywords: ['WSIB', 'PTU'] };
      const params2: SearchParams = { keywords: ['ptu', 'wsib'] };
      const n1 = normalizeSearchParams(params1);
      const n2 = normalizeSearchParams(params2);
      expect(n1).toEqual(n2);
    });
  });

  describe('generateCacheKey', () => {
    it('generates a key with the expected format', () => {
      const key = generateCacheKey(baseParams, 'calendar', 'test-prefix');
      expect(key).toMatch(/^test-prefix:calendar:[a-f0-9]{16}$/);
    });

    it('generates different keys for different domains', () => {
      const calKey = generateCacheKey(baseParams, 'calendar', 'prefix');
      const chatKey = generateCacheKey(baseParams, 'chat', 'prefix');
      expect(calKey).not.toBe(chatKey);
    });

    it('generates the same key for equivalent params in different order', () => {
      const params1: SearchParams = { keywords: ['WSIB', 'PTU'] };
      const params2: SearchParams = { keywords: ['ptu', 'wsib'] };
      const key1 = generateCacheKey(params1, 'calendar', 'prefix');
      const key2 = generateCacheKey(params2, 'calendar', 'prefix');
      expect(key1).toBe(key2);
    });

    it('generates different keys for different search params', () => {
      const params1: SearchParams = { keywords: ['PTU'] };
      const params2: SearchParams = { keywords: ['WSIB'] };
      const key1 = generateCacheKey(params1, 'calendar', 'prefix');
      const key2 = generateCacheKey(params2, 'calendar', 'prefix');
      expect(key1).not.toBe(key2);
    });

    it('generates different keys for different key prefixes', () => {
      const key1 = generateCacheKey(baseParams, 'calendar', 'prefix1');
      const key2 = generateCacheKey(baseParams, 'calendar', 'prefix2');
      expect(key1).not.toBe(key2);
    });
  });

  describe('generateSearchPatternKey', () => {
    it('generates a wildcard pattern key', () => {
      const pattern = generateSearchPatternKey(baseParams, 'test-prefix');
      expect(pattern).toMatch(/^test-prefix:\*:[a-f0-9]{16}$/);
    });

    it('pattern hash matches single domain key hash', () => {
      const pattern = generateSearchPatternKey(baseParams, 'prefix');
      const calKey = generateCacheKey(baseParams, 'calendar', 'prefix');
      const patternHash = pattern.split(':').pop();
      const calHash = calKey.split(':').pop();
      expect(patternHash).toBe(calHash);
    });
  });
});
