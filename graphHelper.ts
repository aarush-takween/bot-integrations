import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { AppSettings } from './appSettings';

let _settings: AppSettings | undefined = undefined;
let _clientSecretCredential: ClientSecretCredential | undefined = undefined;
let _appClient: Client | undefined = undefined;

export function initializeGraphForAppOnlyAuth(settings: AppSettings) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  if (!_clientSecretCredential) {
    _clientSecretCredential = new ClientSecretCredential(
      _settings.tenantId,
      _settings.clientId,
      _settings.clientSecret
    );
  }

  if (!_appClient) {
    const authProvider = new TokenCredentialAuthenticationProvider(_clientSecretCredential, {
      scopes: [ 'https://graph.microsoft.com/.default' ]
    });

    _appClient = Client.initWithMiddleware({
      authProvider: authProvider
    });
  }
}

export async function getUserAsync(userAdId: string): Promise<PageCollection> {
    // Ensure client isn't undefined
    if (!_appClient) {
      throw new Error('Graph has not been initialized for app-only auth');
    }
  
    return _appClient?.api('/users/' + userAdId)
      .select(['displayName', 'id', 'mail'])
      .top(25)
      .orderby('displayName')
      .get();
  }