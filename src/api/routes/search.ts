/**
 * Search API routes.
 *
 * POST   /api/v1/search              — structured search with filters
 * POST   /api/v1/search/natural      — natural language search
 * GET    /api/v1/search/:searchId/results — retrieve async search results
 */

import { Router, Request, Response, NextFunction } from 'express';
import { v4 as uuidv4 } from 'uuid';
import { SearchRequestSchema, NaturalSearchRequestSchema } from '../schemas';
import { search } from '../../orchestrator';
import { parseNaturalLanguage } from '../../nlParser';
import { searchStore } from '../store';
import { searchRateLimiter } from '../middleware/rateLimiter';
import { NotFoundError } from '../middleware/errorHandler';
import { SearchRequest } from '../../types';

const router = Router();

/**
 * @openapi
 * /api/v1/search:
 *   post:
 *     summary: Execute a structured cross-domain meeting search
 *     tags: [Search]
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             $ref: '#/components/schemas/SearchRequest'
 *     responses:
 *       202:
 *         description: Search accepted and queued; use the returned searchId to poll for results
 *       422:
 *         description: Validation error
 *       429:
 *         description: Rate limit exceeded
 *       500:
 *         description: Internal server error
 */
router.post('/', searchRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const parsed = SearchRequestSchema.parse(req.body);
    const searchRequest: SearchRequest = {
      ...parsed,
      dateRange: parsed.dateRange
        ? { start: new Date(parsed.dateRange.start), end: new Date(parsed.dateRange.end) }
        : undefined,
    };

    const searchId = uuidv4();
    searchStore.set({
      searchId,
      status: 'running',
      request: searchRequest,
      createdAt: new Date(),
    });

    // Execute search asynchronously
    search(searchRequest)
      .then((response) => {
        searchStore.update(searchId, {
          status: 'completed',
          response,
          completedAt: new Date(),
        });
      })
      .catch((err: unknown) => {
        searchStore.update(searchId, {
          status: 'failed',
          error: err instanceof Error ? err.message : 'Unknown error',
          completedAt: new Date(),
        });
      });

    res.status(202).json({
      searchId,
      status: 'running',
      message: 'Search queued. Use GET /api/v1/search/:searchId/results to retrieve results.',
      resultsUrl: `/api/v1/search/${searchId}/results`,
    });
  } catch (err) {
    next(err);
  }
});

/**
 * @openapi
 * /api/v1/search/natural:
 *   post:
 *     summary: Execute a natural language meeting search
 *     tags: [Search]
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             $ref: '#/components/schemas/NaturalSearchRequest'
 *     responses:
 *       202:
 *         description: Search accepted and queued
 *       422:
 *         description: Validation error
 *       429:
 *         description: Rate limit exceeded
 *       500:
 *         description: Internal server error
 */
router.post('/natural', searchRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { hint } = NaturalSearchRequestSchema.parse(req.body);
    const searchRequest = parseNaturalLanguage(hint);

    const searchId = uuidv4();
    searchStore.set({
      searchId,
      status: 'running',
      request: searchRequest,
      createdAt: new Date(),
    });

    search(searchRequest)
      .then((response) => {
        searchStore.update(searchId, {
          status: 'completed',
          response,
          completedAt: new Date(),
        });
      })
      .catch((err: unknown) => {
        searchStore.update(searchId, {
          status: 'failed',
          error: err instanceof Error ? err.message : 'Unknown error',
          completedAt: new Date(),
        });
      });

    res.status(202).json({
      searchId,
      status: 'running',
      hint,
      parsedRequest: searchRequest,
      message: 'Search queued. Use GET /api/v1/search/:searchId/results to retrieve results.',
      resultsUrl: `/api/v1/search/${searchId}/results`,
    });
  } catch (err) {
    next(err);
  }
});

/**
 * @openapi
 * /api/v1/search/{searchId}/results:
 *   get:
 *     summary: Retrieve the results of an async search
 *     tags: [Search]
 *     parameters:
 *       - in: path
 *         name: searchId
 *         required: true
 *         schema:
 *           type: string
 *           format: uuid
 *     responses:
 *       200:
 *         description: Search results (may still be running)
 *       404:
 *         description: Search ID not found
 *       500:
 *         description: Internal server error
 */
router.get('/:searchId/results', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { searchId } = req.params;
    const record = searchStore.get(searchId);

    if (!record) {
      return next(new NotFoundError(`Search with ID '${searchId}' not found`));
    }

    if (record.status === 'running' || record.status === 'pending') {
      return res.status(200).json({
        searchId,
        status: record.status,
        message: 'Search is still in progress. Please try again shortly.',
        createdAt: record.createdAt,
      });
    }

    if (record.status === 'failed') {
      return res.status(200).json({
        searchId,
        status: 'failed',
        error: record.error,
        createdAt: record.createdAt,
        completedAt: record.completedAt,
      });
    }

    return res.status(200).json({
      searchId,
      status: 'completed',
      results: record.response,
      createdAt: record.createdAt,
      completedAt: record.completedAt,
    });
  } catch (err) {
    next(err);
  }
});

export default router;
