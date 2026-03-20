import { Client } from '@microsoft/microsoft-graph-client';

/**
 * Thin wrapper around the Microsoft Graph client that makes it easy to swap
 * in a mock implementation during tests.
 */
export class GraphClientWrapper {
  private client: Client;

  constructor(client: Client) {
    this.client = client;
  }

  /**
   * Fetches a page of messages matching the supplied OData query parameters.
   *
   * @param queryParams  Key-value map of OData query parameters such as
   *                     `$filter`, `$search`, `$top`, and `$expand`.
   */
  async getMessages(
    queryParams: Record<string, string | number>,
  ): Promise<{ value: unknown[] }> {
    const request = this.client.api('/me/messages');

    for (const [key, value] of Object.entries(queryParams)) {
      request.query({ [key]: value });
    }

    return request.get() as Promise<{ value: unknown[] }>;
  }
}

/**
 * Creates a GraphClientWrapper backed by a real MSAL access-token provider.
 *
 * @param accessTokenProvider  A zero-argument async function that resolves to a
 *                             valid Bearer token for the Microsoft Graph API.
 */
export function createGraphClientWrapper(
  accessTokenProvider: () => Promise<string>,
): GraphClientWrapper {
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: accessTokenProvider,
    },
  });
  return new GraphClientWrapper(client);
}
