/**
 * Rate limiting configuration for search endpoints.
 */

import rateLimit from 'express-rate-limit';

/** 100 requests per 15 minutes per IP on search endpoints. */
export const searchRateLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many search requests, please try again later.' },
});

/** 30 reconstruction requests per 15 minutes per IP. */
export const reconstructionRateLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 30,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many reconstruction requests, please try again later.' },
});
