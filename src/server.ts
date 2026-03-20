/**
 * HTTP server entry point.
 */

import { createApp } from './api/app';

const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const app = createApp();

app.listen(PORT, () => {
  console.log(`Enterprise Meeting Search API listening on port ${PORT}`);
  console.log(`OpenAPI docs: http://localhost:${PORT}/api-docs`);
});
