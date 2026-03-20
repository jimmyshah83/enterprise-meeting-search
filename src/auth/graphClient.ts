import {
  Client,
  AuthenticationProvider,
  AuthenticationProviderOptions,
} from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

/**
 * Options required to authenticate against Microsoft Graph using a service
 * principal (client credentials flow).
 */
export interface GraphClientOptions {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

/**
 * Lightweight {@link AuthenticationProvider} that delegates token acquisition
 * to an `@azure/identity` `ClientSecretCredential`.
 */
class AzureIdentityAuthProvider implements AuthenticationProvider {
  private readonly credential: ClientSecretCredential;
  private static readonly GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

  constructor(credential: ClientSecretCredential) {
    this.credential = credential;
  }

  async getAccessToken(
    _options?: AuthenticationProviderOptions,
  ): Promise<string> {
    const token = await this.credential.getToken(
      AzureIdentityAuthProvider.GRAPH_SCOPE,
    );
    return token.token;
  }
}

/**
 * Creates and returns a Microsoft Graph {@link Client} authenticated with the
 * provided service-principal credentials.
 *
 * @param options - Tenant, client-ID and client-secret for the app registration.
 * @returns An authenticated Graph client ready to make API calls.
 */
export function createGraphClient(options: GraphClientOptions): Client {
  const credential = new ClientSecretCredential(
    options.tenantId,
    options.clientId,
    options.clientSecret,
  );

  const authProvider = new AzureIdentityAuthProvider(credential);

  return Client.initWithMiddleware({ authProvider });
}
