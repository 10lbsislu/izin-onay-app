/**
 * shared/graphClient.js
 * Microsoft Graph API client.
 * İstemci kimlik bilgileriyle (service principal) çalışır.
 * Kullanıcı token'ı OBO flow ile değiştirilebilir.
 */

const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require(
  "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials"
);

let _clientInstance = null;

/**
 * Servis hesabı Graph client'ı (uygulama izinleri)
 */
function getAppGraphClient() {
  if (_clientInstance) return _clientInstance;

  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  _clientInstance = Client.initWithMiddleware({ authProvider });
  return _clientInstance;
}

/**
 * Gelen Bearer token'dan delegated Graph client oluşturur.
 * Teams SSO OBO exchange burada yapılır.
 */
async function getUserGraphClient(bearerToken) {
  // OBO (On-Behalf-Of) flow ile kullanıcı adına token al
  const params = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    assertion: bearerToken,
    requested_token_use: "on_behalf_of",
    scope: "https://graph.microsoft.com/User.ReadBasic.All User.Read",
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    }
  );

  const tokenData = await response.json();

  if (!response.ok) {
    throw new Error(`OBO exchange failed: ${tokenData.error_description}`);
  }

  return Client.init({
    authProvider: (done) => done(null, tokenData.access_token),
  });
}

module.exports = { getAppGraphClient, getUserGraphClient };
