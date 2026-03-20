import axios, { AxiosInstance, AxiosError } from 'axios';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/** Maximum number of retries on rate-limit (HTTP 429) responses. */
const MAX_RETRIES = 3;
/** Base delay in milliseconds for exponential back-off. */
const RETRY_BASE_DELAY_MS = 1000;

/**
 * Lightweight Microsoft Graph API client.
 *
 * Handles:
 *   - Bearer-token injection
 *   - Automatic retry with exponential back-off on HTTP 429 (rate-limit) and
 *     transient 5xx errors
 *
 * The caller is responsible for acquiring and refreshing the access token
 * before it expires.
 */
export class GraphClient {
  private readonly http: AxiosInstance;

  constructor(accessToken: string) {
    this.http = axios.create({
      baseURL: GRAPH_BASE_URL,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });
  }

  /**
   * Perform a POST request against a Graph API endpoint.
   *
   * @param path   - Path relative to the Graph base URL (e.g. `/search/query`).
   * @param body   - JSON request body.
   * @returns Parsed JSON response.
   */
  async post<T>(path: string, body: unknown): Promise<T> {
    return this.withRetry(() => this.http.post<T>(path, body).then((r) => r.data));
  }

  /**
   * Perform a GET request against a Graph API endpoint.
   *
   * @param path - Path relative to the Graph base URL.
   * @returns Parsed JSON response.
   */
  async get<T>(path: string): Promise<T> {
    return this.withRetry(() => this.http.get<T>(path).then((r) => r.data));
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  private async withRetry<T>(fn: () => Promise<T>, attempt = 0): Promise<T> {
    try {
      return await fn();
    } catch (err) {
      const axiosErr = err as AxiosError;
      const status = axiosErr.response?.status;

      const isRetryable = status === 429 || (status !== undefined && status >= 500);
      if (isRetryable && attempt < MAX_RETRIES) {
        const retryAfterHeader = axiosErr.response?.headers?.['retry-after'];
        const retryAfterMs = retryAfterHeader
          ? parseInt(retryAfterHeader as string, 10) * 1000
          : RETRY_BASE_DELAY_MS * Math.pow(2, attempt);

        await delay(retryAfterMs);
        return this.withRetry(fn, attempt + 1);
      }

      throw err;
    }
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
