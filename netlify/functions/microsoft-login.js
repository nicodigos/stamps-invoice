const { randomUUID } = require("node:crypto");

function buildOrigin(event) {
  const proto = event.headers["x-forwarded-proto"] || "https";
  const host = event.headers.host;
  return `${proto}://${host}`;
}

exports.handler = async function handler(event) {
  const tenantId = process.env.TENANT_ID || "";
  const clientId = process.env.CLIENT_ID || "";
  if (!tenantId || !clientId) {
    return { statusCode: 500, body: "Missing TENANT_ID or CLIENT_ID." };
  }

  const state = randomUUID();
  const redirectUri = process.env.REDIRECT_URI || `${buildOrigin(event)}/`;
  const authUrl = new URL(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set("client_id", clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", "User.Read Files.ReadWrite.All");
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("prompt", "select_account");

  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ authUrl: authUrl.toString(), state }),
  };
};
