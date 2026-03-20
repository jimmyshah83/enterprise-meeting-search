/**
 * Integration tests for the Enterprise Meeting Search REST API.
 *
 * Tests all 4 endpoints:
 *  1. POST /api/v1/search
 *  2. POST /api/v1/search/natural
 *  3. GET  /api/v1/search/:searchId/results
 *  4. POST /api/v1/reconstruct/:meetingId
 */

import request from 'supertest';
import { createApp } from '../api/app';
import { searchStore } from '../api/store';
import { v4 as uuidv4 } from 'uuid';
import { AsyncSearchRecord, SearchResponse } from '../types';

const app = createApp();

// ---------------------------------------------------------------------------
// Health check
// ---------------------------------------------------------------------------

describe('GET /health', () => {
  it('returns 200 with ok status', async () => {
    const res = await request(app).get('/health');
    expect(res.status).toBe(200);
    expect(res.body.status).toBe('ok');
    expect(res.body.timestamp).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// POST /api/v1/search
// ---------------------------------------------------------------------------

describe('POST /api/v1/search', () => {
  it('returns 202 with a searchId for valid keyword search', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({ keywords: ['budget', 'Q4'] });

    expect(res.status).toBe(202);
    expect(res.body.searchId).toBeDefined();
    expect(res.body.status).toBe('running');
    expect(res.body.resultsUrl).toMatch(/\/api\/v1\/search\/.+\/results/);
  });

  it('returns 202 for a date range search', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({
        dateRange: {
          start: '2024-01-01T00:00:00Z',
          end: '2024-03-31T23:59:59Z',
        },
      });

    expect(res.status).toBe(202);
    expect(res.body.searchId).toBeDefined();
  });

  it('returns 202 for attendee + meetingType search', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({
        attendees: ['john.smith@example.com'],
        meetingType: ['Teams'],
      });

    expect(res.status).toBe(202);
    expect(res.body.searchId).toBeDefined();
  });

  it('returns 422 when no search criterion is provided', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({});

    expect(res.status).toBe(422);
    expect(res.body.error).toBeDefined();
  });

  it('returns 422 for invalid attendee email', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({ attendees: ['not-an-email'] });

    expect(res.status).toBe(422);
    expect(res.body.error).toBeDefined();
  });

  it('returns 422 when dateRange.start is after dateRange.end', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send({
        dateRange: {
          start: '2024-12-31T00:00:00Z',
          end: '2024-01-01T00:00:00Z',
        },
      });

    expect(res.status).toBe(422);
  });

  it('returns 422 for invalid JSON body (non-object)', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .send([]);

    expect(res.status).toBe(422);
  });
});

// ---------------------------------------------------------------------------
// POST /api/v1/search/natural
// ---------------------------------------------------------------------------

describe('POST /api/v1/search/natural', () => {
  it('returns 202 with searchId and parsedRequest for a valid hint', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({ hint: 'mid-January, John Smith was there' });

    expect(res.status).toBe(202);
    expect(res.body.searchId).toBeDefined();
    expect(res.body.hint).toBe('mid-January, John Smith was there');
    expect(res.body.parsedRequest).toBeDefined();
    expect(res.body.resultsUrl).toMatch(/\/api\/v1\/search\/.+\/results/);
  });

  it('parses meeting type from hint', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({ hint: 'Teams meeting last week about roadmap' });

    expect(res.status).toBe(202);
    expect(res.body.parsedRequest.meetingType).toContain('Teams');
  });

  it('parses attendee name from hint', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({ hint: 'John Smith was there for the budget review' });

    expect(res.status).toBe(202);
    expect(res.body.parsedRequest.attendees).toContain('John Smith');
  });

  it('returns 422 when hint is missing', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({});

    expect(res.status).toBe(422);
    expect(res.body.error).toBeDefined();
  });

  it('returns 422 when hint is too short', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({ hint: 'ab' });

    expect(res.status).toBe(422);
  });

  it('returns 422 when hint is too long', async () => {
    const res = await request(app)
      .post('/api/v1/search/natural')
      .send({ hint: 'a'.repeat(501) });

    expect(res.status).toBe(422);
  });
});

// ---------------------------------------------------------------------------
// GET /api/v1/search/:searchId/results
// ---------------------------------------------------------------------------

describe('GET /api/v1/search/:searchId/results', () => {
  it('returns 404 for an unknown searchId', async () => {
    const res = await request(app)
      .get(`/api/v1/search/${uuidv4()}/results`);

    expect(res.status).toBe(404);
    expect(res.body.error).toBeDefined();
  });

  it('returns 200 with status=running for an in-progress search', async () => {
    // Seed a running record
    const searchId = uuidv4();
    searchStore.set({
      searchId,
      status: 'running',
      request: { keywords: ['test'] },
      createdAt: new Date(),
    });

    const res = await request(app)
      .get(`/api/v1/search/${searchId}/results`);

    expect(res.status).toBe(200);
    expect(res.body.status).toBe('running');
    expect(res.body.searchId).toBe(searchId);
  });

  it('returns 200 with status=completed and results for a finished search', async () => {
    const searchId = uuidv4();
    const mockResponse: SearchResponse = {
      results: [],
      totalResults: 0,
      searchDurationMs: 10,
      diagnostics: {
        message: 'No results found across all search domains.',
        reasons: ['No results found'],
        suggestions: ['Try relaxing the search filters.'],
        domainStatuses: {
          calendar: { status: 'success', resultCount: 0 },
          transcripts: { status: 'success', resultCount: 0 },
          chats: { status: 'success', resultCount: 0 },
          email: { status: 'success', resultCount: 0 },
          files: { status: 'success', resultCount: 0 },
        },
      },
    };

    searchStore.set({
      searchId,
      status: 'completed',
      request: { keywords: ['test'] },
      response: mockResponse,
      createdAt: new Date(),
      completedAt: new Date(),
    });

    const res = await request(app)
      .get(`/api/v1/search/${searchId}/results`);

    expect(res.status).toBe(200);
    expect(res.body.status).toBe('completed');
    expect(res.body.results).toBeDefined();
    expect(res.body.completedAt).toBeDefined();
  });

  it('returns 200 with status=failed for a failed search', async () => {
    const searchId = uuidv4();
    searchStore.set({
      searchId,
      status: 'failed',
      request: { keywords: ['test'] },
      error: 'Connection refused',
      createdAt: new Date(),
      completedAt: new Date(),
    } as AsyncSearchRecord);

    const res = await request(app)
      .get(`/api/v1/search/${searchId}/results`);

    expect(res.status).toBe(200);
    expect(res.body.status).toBe('failed');
    expect(res.body.error).toBe('Connection refused');
  });

  it('returns the searchId in the response', async () => {
    // First create a search
    const postRes = await request(app)
      .post('/api/v1/search')
      .send({ keywords: ['integration-test'] });

    expect(postRes.status).toBe(202);
    const { searchId } = postRes.body;

    const getRes = await request(app)
      .get(`/api/v1/search/${searchId}/results`);

    expect(getRes.status).toBe(200);
    expect(getRes.body.searchId).toBe(searchId);
  });
});

// ---------------------------------------------------------------------------
// POST /api/v1/reconstruct/:meetingId
// ---------------------------------------------------------------------------

describe('POST /api/v1/reconstruct/:meetingId', () => {
  it('returns 404 when meeting has no search results', async () => {
    const res = await request(app)
      .post('/api/v1/reconstruct/nonexistent-meeting-id');

    expect(res.status).toBe(404);
    expect(res.body.error).toMatch(/nonexistent-meeting-id/);
  });

  it('returns 200 or 202 with reconstruction for a known meeting', async () => {
    // Seed a completed search with a meeting result
    const meetingId = `meeting-${uuidv4()}`;
    const searchId = uuidv4();
    searchStore.set({
      searchId,
      status: 'completed',
      request: { keywords: ['test'] },
      response: {
        results: [
          {
            meetingId,
            calendarEvent: {
              id: meetingId,
              title: 'Q4 Budget Review',
              startTime: new Date('2024-01-15T09:00:00Z'),
              endTime: new Date('2024-01-15T10:00:00Z'),
              organizer: 'alice@example.com',
              attendees: ['alice@example.com', 'bob@example.com'],
              meetingType: 'Teams',
              source: 'calendar',
            },
            transcripts: [
              {
                id: 'transcript-1',
                meetingId,
                content: 'We decided to proceed with the budget. Action item: Bob will review the Q4 report.',
                speakers: ['Alice', 'Bob'],
                timestamp: new Date('2024-01-15T10:01:00Z'),
                source: 'transcript',
              },
            ],
            chats: [],
            emails: [],
            files: [],
          },
        ],
        totalResults: 1,
        searchDurationMs: 5,
      },
      createdAt: new Date(),
      completedAt: new Date(),
    });

    const res = await request(app)
      .post(`/api/v1/reconstruct/${meetingId}`);

    expect([200, 202]).toContain(res.status);
    expect(res.body.meetingId).toBe(meetingId);
    expect(res.body.reconstructionId).toBeDefined();

    if (res.status === 200) {
      expect(res.body.reconstruction).toBeDefined();
      expect(res.body.reconstruction.overallConfidence).toBeDefined();
    }
  });
});

// ---------------------------------------------------------------------------
// Error handling
// ---------------------------------------------------------------------------

describe('Error handling', () => {
  it('returns 404 for unknown routes', async () => {
    const res = await request(app).get('/api/v1/unknown-route');
    expect(res.status).toBe(404);
    expect(res.body.error).toBe('Route not found');
  });

  it('returns 400 for malformed JSON', async () => {
    const res = await request(app)
      .post('/api/v1/search')
      .set('Content-Type', 'application/json')
      .send('{ invalid json }');

    expect(res.status).toBe(400);
  });
});

// ---------------------------------------------------------------------------
// OpenAPI docs
// ---------------------------------------------------------------------------

describe('GET /api-docs.json', () => {
  it('returns the OpenAPI spec', async () => {
    const res = await request(app).get('/api-docs.json');
    expect(res.status).toBe(200);
    expect(res.body.openapi).toBe('3.0.3');
    expect(res.body.info.title).toBe('Enterprise Meeting Search API');
    expect(res.body.paths['/api/v1/search']).toBeDefined();
    expect(res.body.paths['/api/v1/search/natural']).toBeDefined();
    expect(res.body.paths['/api/v1/search/{searchId}/results']).toBeDefined();
    expect(res.body.paths['/api/v1/reconstruct/{meetingId}']).toBeDefined();
  });
});
