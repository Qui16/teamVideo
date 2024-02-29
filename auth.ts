import { Client } from "@microsoft/microsoft-graph-client";
import { AuthorizationCodeCredential, ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";


const clientId = '61af1d10-38df-4b9b-b4b4-3963cb53c547';
const clientSecret = 'j758Q~T_AceWIbNTGZvv7RqHoaK8F.6Xo9~4hcZ8';
const tenantId = '5c0d2688-aa31-4ac3-bd1c-f31c96bca508';

// @azure/identity
const credential = new ClientSecretCredential(
    tenantId,
    clientId,
    clientSecret,
);

// @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    // The client credentials flow requires that you request the
    // /.default scope, and pre-configure your permissions on the
    // app registration in Azure. An administrator must grant consent
    // to those permissions beforehand.
    scopes: ['https://graph.microsoft.com/.default'],
});

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });