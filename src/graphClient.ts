import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import 'isomorphic-fetch';
import { GraphClientConfig } from './types';

/**
 * Builds an authenticated Microsoft Graph client using the client-credentials
 * (app-only) OAuth flow.
 *
 * The returned client can call Graph endpoints on behalf of the application
 * (e.g. reading a specific user's calendar when the app has been granted
 * `Calendars.Read` application permission).
 */
export function createGraphClient(config: GraphClientConfig): Client {
  const credential = new ClientSecretCredential(
    config.tenantId,
    config.clientId,
    config.clientSecret,
  );

  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken(
          'https://graph.microsoft.com/.default',
        );
        return token.token;
      },
    },
  });
}
