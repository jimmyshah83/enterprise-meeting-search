/**
 * Express error handling and validation middleware.
 */

import { Request, Response, NextFunction } from 'express';
import { ZodError } from 'zod';

export interface ApiError {
  error: string;
  details?: unknown;
  statusCode: number;
}

export class NotFoundError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'NotFoundError';
  }
}

export class ValidationError extends Error {
  constructor(
    message: string,
    public readonly details: unknown,
  ) {
    super(message);
    this.name = 'ValidationError';
  }
}

/** Centralized error-handling middleware. */
export function errorHandler(
  err: unknown,
  _req: Request,
  res: Response,
  _next: NextFunction,
): void {
  // Body-parser JSON parse errors
  if (typeof err === 'object' && err !== null && 'statusCode' in err && (err as { statusCode: number }).statusCode === 400 && 'type' in err) {
    res.status(400).json({ error: 'Invalid JSON body' });
    return;
  }

  if (err instanceof ZodError) {
    res.status(422).json({
      error: 'Validation error',
      details: err.errors.map((e) => ({ path: e.path.join('.'), message: e.message })),
    });
    return;
  }

  if (err instanceof ValidationError) {
    res.status(422).json({ error: err.message, details: err.details });
    return;
  }

  if (err instanceof NotFoundError) {
    res.status(404).json({ error: err.message });
    return;
  }

  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
}

/** 404 handler for unknown routes. */
export function notFoundHandler(_req: Request, res: Response): void {
  res.status(404).json({ error: 'Route not found' });
}
