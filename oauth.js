// OAuth2 PKCE flow using Thunderbird's registered public client.
// Reuses Mozilla's existing Microsoft app registration (client_id 9e5f94bc-...)
// — NO Azure app registration required. Produces 90-day rolling refresh tokens
// (public client, not SPA), so the user logs in once and never needs to re-login
// (as long as Thunderbird runs at least once every 90 days).
var TbOAuth = {
  // Thunderbird's registered Microsoft public client ID. This is the same
  // client_id Thunderbird uses for IMAP/EWS OAuth2 — we reuse it, no
  // registration needed on our end.
  CLIENT_ID: "9e5f94bc-e8a4-4e73-b8be-63364c29d753",

  // Microsoft OAuth2 v2.0 endpoints (common tenant — works for any M365 account).
  AUTH_ENDPOINT: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
  TOKEN_ENDPOINT: "https://login.microsoftonline.com/common/oauth2/v2.0/token",

  // Redirect URI: http://localhost is registered for public clients.
  // We intercept this redirect via webRequest before it hits the network.
  REDIRECT_URI: "http://localhost",

  // Exchange Online scopes for calendar + contacts (same audience as OWA service.svc).
  // offline_access ensures we get a refresh token.
  SCOPES: "https://outlook.office.com/Calendars.ReadWrite https://outlook.office.com/Contacts.ReadWrite offline_access",

  // Base64url-encode an ArrayBuffer (for PKCE + crypto random values).
  _base64Url(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    let str = "";
    for (const b of bytes) str += String.fromCharCode(b);
    return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  },

  // Generate a random PKCE code_verifier (43+ chars, base64url of 32 random bytes).
  _generateVerifier() {
    const bytes = new Uint8Array(32);
    crypto.getRandomValues(bytes);
    return this._base64Url(bytes.buffer);
  },

  // Compute the PKCE code_challenge (S256 method: base64url(SHA-256(verifier))).
  async _computeChallenge(verifier) {
    const data = new TextEncoder().encode(verifier);
    const hash = await crypto.subtle.digest("SHA-256", data);
    return this._base64Url(hash);
  },

  // Generate a random state parameter (CSRF protection for the redirect).
  _generateState() {
    const bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    return this._base64Url(bytes.buffer);
  },

  // Build the authorization URL for the interactive login flow.
  // Returns { url, verifier, state } — caller stores verifier + state for the
  // redirect handler and opens the URL in a tab for the user to log in.
  async buildAuthUrl() {
    const verifier = this._generateVerifier();
    const challenge = await this._computeChallenge(verifier);
    const state = this._generateState();
    const params = new URLSearchParams({
      client_id: this.CLIENT_ID,
      response_type: "code",
      redirect_uri: this.REDIRECT_URI,
      scope: this.SCOPES,
      code_challenge: challenge,
      code_challenge_method: "S256",
      state: state,
      prompt: "select_account",
    });
    return { url: this.AUTH_ENDPOINT + "?" + params.toString(), verifier, state };
  },

  // Exchange an authorization code for access + refresh tokens (PKCE).
  // `code` is from the redirect URL; `verifier` is the PKCE code_verifier.
  async exchangeCode(code, verifier) {
    const body = new URLSearchParams({
      client_id: this.CLIENT_ID,
      grant_type: "authorization_code",
      code: code,
      redirect_uri: this.REDIRECT_URI,
      code_verifier: verifier,
      scope: this.SCOPES,
    });
    console.log("[M365OWA] PKCE: exchanging authorization code for tokens");
    const resp = await fetch(this.TOKEN_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      credentials: "omit",
    });
    if (resp.status !== 200) {
      const txt = await resp.text().catch(() => "");
      throw new Error("token exchange failed (" + resp.status + "): " + txt.slice(0, 300));
    }
    const data = await resp.json();
    if (!data.access_token) {
      throw new Error("no access_token in token response: " + JSON.stringify(data).slice(0, 200));
    }
    console.log("[M365OWA] PKCE: token exchange succeeded" +
      " (access_token len=" + data.access_token.length +
      " refresh_token=" + (data.refresh_token ? "yes(len=" + data.refresh_token.length + ")" : "no") +
      " expires_in=" + data.expires_in +
      " scope=" + (data.scope || "?").slice(0, 100) + ")");
    return data;
  },

  // Refresh the access token using a stored refresh token (90-day rolling).
  // Returns { ok, data } on success or { ok: false, reason, invalidGrant } on failure.
  async refreshToken(refreshToken) {
    const body = new URLSearchParams({
      client_id: this.CLIENT_ID,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: this.SCOPES,
    });
    console.log("[M365OWA] PKCE: refreshing access token (refresh_token len=" + refreshToken.length + ")");
    const resp = await fetch(this.TOKEN_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      credentials: "omit",
    });
    if (resp.status !== 200) {
      const txt = await resp.text().catch(() => "");
      const invalidGrant = /"error"\s*:\s*"invalid_grant"/i.test(txt);
      return { ok: false, reason: resp.status + ": " + txt.slice(0, 300), invalidGrant };
    }
    const data = await resp.json();
    if (!data.access_token) {
      return { ok: false, reason: "no access_token in refresh response: " + JSON.stringify(data).slice(0, 200) };
    }
    console.log("[M365OWA] PKCE: refresh succeeded" +
      " (access_token len=" + data.access_token.length +
      " refresh_token=" + (data.refresh_token ? "rotated" : "same") + ")");
    return { ok: true, data };
  },
};
