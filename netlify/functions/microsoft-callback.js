exports.handler = async function handler(event) {
  const code = event.queryStringParameters?.code || "";
  const redirectUri = event.queryStringParameters?.redirect_uri || process.env.REDIRECT_URI || "";
  if (!code) {
    return { statusCode: 400, body: "Missing authorization code." };
  }

  const tenantId = process.env.TENANT_ID || "";
  const clientId = process.env.CLIENT_ID || "";
  const clientSecret = process.env.CLIENT_SECRET || "";
  if (!tenantId || !clientId || !clientSecret || !redirectUri) {
    return { statusCode: 500, body: "Incomplete Microsoft configuration." };
  }

  const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      code,
      redirect_uri: redirectUri,
      grant_type: "authorization_code",
      scope: "User.Read Files.ReadWrite.All",
    }),
  });

  const payload = await tokenResponse.json();
  if (!tokenResponse.ok || !payload.access_token) {
    return { statusCode: 500, body: JSON.stringify(payload) };
  }

  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ accessToken: payload.access_token }),
  };
};
