/**
 * Express application factory.
 * Exported separately from server.ts so tests can import it without starting the HTTP server.
 */

import express, { Application } from 'express';
import swaggerUi from 'swagger-ui-express';
import { errorHandler, notFoundHandler } from './middleware/errorHandler';
import searchRouter from './routes/search';
import reconstructRouter from './routes/reconstruct';
import { openApiSpec } from './openapi';

export function createApp(): Application {
  const app = express();

  // Parse JSON bodies
  app.use(express.json());

  // Health check
  app.get('/health', (_req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
  });

  // OpenAPI / Swagger UI
  app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(openApiSpec));
  app.get('/api-docs.json', (_req, res) => res.json(openApiSpec));

  // API routes
  app.use('/api/v1/search', searchRouter);
  app.use('/api/v1/reconstruct', reconstructRouter);

  // 404 and error handlers (must be last)
  app.use(notFoundHandler);
  app.use(errorHandler);

  return app;
}
