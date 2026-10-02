// Bearer-token manager for OWA service.svc requests.
var Auth = {
  // current bearer string (without "Bearer " prefix)
  _token: null,
  // set true when OWA rejects with 401/403
  _expired: false,
  // ms epoch when the current token was captured or pasted
  _capturedAt: 0,
  // id of the OWA content tab used for login and refresh (memory only)
  _owaTabId: null,
  // treat stored tokens older than this as stale (OWA bearer tokens live ~60 min)
  _maxTokenAgeMs: 55 * 60 * 1000,

  // --- refresh-token harvesting (survives reboots, no SSO dependency) ---
  // ALL collected refresh tokens from multiple sources. We store every one we
  // find and try each at renewal time — the Azure AD token endpoint is the only
  // real validator, so we avoid format heuristics and let the server tell us
  // which token works. webRequest-sourced tokens (actual MS HTTP response) are
  // pinned at the front of the array; content-script tokens (MSAL cache) are
  // appended as fallback.
  // Entry shape: { refreshToken, clientId, url, scope, source, ts }
  _refreshTokens: [],
  // flag: true while we're making our own token refresh fetch (so the webRequest harvester skips us)
  _selfRefreshActive: false,

  // storage key for the token
  _key() { return "m365_owa_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the expired flag
  _expKey() { return "m365_owa_expired_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the capture timestamp
  _tsKey() { return "m365_owa_captured_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the refresh-token array (plural — distinct from legacy single-value keys)
  _rkey() { return "m365_owa_refresh_tokens_" + (CONFIG.CONNECTION_NAME || "default"); },

  // Restore token, expired flag, and capture timestamp from storage.local into memory.
  async loadStoredToken() {
    // resolve the token storage key
    const k = this._key();
    // read the persisted token
    const { [k]: tok } = await browser.storage.local.get(k);
    // resolve the expired-flag storage key
    const ek = this._expKey();
    // read the persisted expired flag
    const { [ek]: exp } = await browser.storage.local.get(ek);
    // resolve the timestamp storage key
    const tk = this._tsKey();
    // read the persisted capture timestamp
    const { [tk]: ts } = await browser.storage.local.get(tk);
    // restore token into memory (or null)
    this._token = tok || null;
    // restore expired flag as boolean
    this._expired = !!exp;
    // restore capture timestamp (or zero)
    this._capturedAt = ts || 0;
    // decide whether the stored token is still usable on startup:
    // prefer the JWT exp claim (authoritative) when the token is a JWT;
    // fall back to the 55-minute age heuristic for non-JWT tokens.
    if (this._token) {
      // try to read the JWT exp (ms epoch) — null for non-JWT tokens
      const jwtExp = this.getTokenExpiry();
      // a token is expired if the JWT exp has passed, or (no JWT) the age heuristic trips
      const stale = (jwtExp != null)
        ? jwtExp <= Date.now()
        : (Date.now() - this._capturedAt) > this._maxTokenAgeMs;
      // apply the stale flag in memory and storage
      if (stale) {
        // mark stale in memory
        this._expired = true;
        // persist the stale flag
        await browser.storage.local.set({ [this._expKey()]: true });
      } else if (this._expired) {
        // stored flag said expired but the token is actually still valid — clear it
        this._expired = false;
        await browser.storage.local.set({ [this._expKey()]: false });
      }
    }
  },

  // Restore the refresh-token array from storage.local into memory.
  // Also cleans up legacy single-value keys from older addon versions.
  async loadStoredRefreshToken() {
    const rk = this._rkey();
    const { [rk]: arr } = await browser.storage.local.get(rk);
    this._refreshTokens = Array.isArray(arr) ? arr : [];
    // clean up legacy single-value keys from older versions
    const conn = CONFIG.CONNECTION_NAME || "default";
    const legacyKeys = [
      "m365_owa_refresh_token_" + conn,
      "m365_owa_refresh_clientid_" + conn,
      "m365_owa_refresh_url_" + conn,
      "m365_owa_refresh_scope_" + conn,
    ];
    // migrate a legacy single-value token into the array when the new array is empty
    if (this._refreshTokens.length === 0) {
      const legacy = await browser.storage.local.get(legacyKeys);
      if (legacy[legacyKeys[0]]) {
        this._refreshTokens.push({
          refreshToken: legacy[legacyKeys[0]],
          clientId: legacy[legacyKeys[1]] || null,
          url: legacy[legacyKeys[2]] || null,
          scope: legacy[legacyKeys[3]] || null,
          source: "legacy-migration",
          ts: Date.now(),
        });
        await this._persistRefreshTokens();
        console.log("[M365OWA] migrated 1 legacy refresh token into array");
      }
    }
    // delete legacy keys regardless
    await browser.storage.local.remove(legacyKeys).catch(() => {});
    // log every loaded token
    if (this._refreshTokens.length) {
      console.log("[M365OWA] loaded " + this._refreshTokens.length + " refresh token(s):");
      for (let i = 0; i < this._refreshTokens.length; i++) {
        const e = this._refreshTokens[i];
        console.log("[M365OWA]   [" + i + "] source=" + (e.source || "?") +
          " start=" + (e.refreshToken || "").slice(0, 20) +
          " end=..." + (e.refreshToken || "").slice(-12) +
          " len=" + (e.refreshToken || "").length +
          " client_id=" + (e.clientId || "?").slice(0, 8) +
          " url=" + (e.url || "?"));
      }
    } else {
      console.log("[M365OWA] no refresh tokens stored");
    }
  },

  // Persist the refresh-token array to storage.local.
  async _persistRefreshTokens() {
    await browser.storage.local.set({ [this._rkey()]: this._refreshTokens });
  },

  // Extract a bare token from an Authorization header value ("Bearer xxx"); returns null if absent.
  _fromHeader(value) {
    // ignore missing values
    if (!value) return null;
    // match the "Bearer xxx" form
    const m = String(value).match(/Bearer\s+([A-Za-z0-9._\-~+/=]+)/i);
    // return the bare token or null
    return m ? m[1] : null;
  },

  // Accept a raw bearer string (or "Bearer xxx" / JSON envelope), persist it, and clear the expired flag.
  async setToken(raw) {
    // reject empty input
    if (!raw) throw new Error("Empty token");
    // normalise to a trimmed string
    let tok = String(raw).trim();
    // try to extract from "Bearer xxx" form
    const m = tok.match(/Bearer\s+([A-Za-z0-9._\-]+)/i);
    // use the bare token
    if (m) tok = m[1];
    else {
      // else try to extract access_token from a JSON envelope
      try { const j = JSON.parse(tok); if (j && j.access_token) tok = j.access_token; } catch {}
    }
    // store the token in memory
    this._token = tok;
    // mark it as valid
    this._expired = false;
    // record the capture time
    this._capturedAt = Date.now();
    // persist token, timestamp, and clear expired flag
    await browser.storage.local.set({ [this._key()]: tok, [this._expKey()]: false, [this._tsKey()]: this._capturedAt });
    console.log("[M365OWA] token stored for connection", CONFIG.CONNECTION_NAME);
  },

  // Store a token harvested from an OWA request header; returns true when it became the active token.
  async harvestFromHeader(value) {
    // extract the bare token
    const tok = this._fromHeader(value);
    // ignore values without a token
    if (!tok) return false;
    // skip tokens that are already active
    if (tok === this._token) {
      // refresh the last-seen timestamp anyway
      this._capturedAt = Date.now();
      // persist the refreshed timestamp
      await browser.storage.local.set({ [this._tsKey()]: this._capturedAt });
      // report that a valid token was seen
      return true;
    }
    // store the new token in memory
    this._token = tok;
    // mark it as valid
    this._expired = false;
    // record the capture time
    this._capturedAt = Date.now();
    // persist token, timestamp, and clear expired flag
    await browser.storage.local.set({ [this._key()]: tok, [this._expKey()]: false, [this._tsKey()]: this._capturedAt });
    console.log("[M365OWA] harvested fresh token from OWA session (len=" + tok.length + ")");
    // report that the active token changed
    return true;
  },

  // Clear the token and expired flag from memory and storage.
  async logout() {
    // clear the in-memory token
    this._token = null;
    // reset the expired flag
    this._expired = false;
    // reset the capture timestamp
    this._capturedAt = 0;
    // forget the OWA tab
    this._owaTabId = null;
    // clear the refresh-token array
    this._refreshTokens = [];
    // delete token, flag, timestamp, refresh-token array, and legacy single-value keys from storage
    const conn = CONFIG.CONNECTION_NAME || "default";
    const legacyKeys = [
      "m365_owa_refresh_token_" + conn,
      "m365_owa_refresh_clientid_" + conn,
      "m365_owa_refresh_url_" + conn,
      "m365_owa_refresh_scope_" + conn,
    ];
    await browser.storage.local.remove([this._key(), this._expKey(), this._tsKey(), this._rkey(), ...legacyKeys]);
  },

  // Return the current token, throwing if none exists or it is flagged expired.
  async getTokenAsync() {
    // error if no token at all
    if (!this._token) throw new Error("No OWA token. Open the addon options and click Connect to log in to OWA.");
    // error if the token was marked expired
    if (this._expired) throw new Error("OWA token expired. Open the addon options and click Connect to renew it.");
    // return the token
    return this._token;
  },

  // Synchronous token accessor; throws if no token or expired (no renewal attempt here).
  getToken() {
    // error if no token
    if (!this._token) throw new Error("No OWA token. Open the addon options and click Connect to log in to OWA.");
    // error if expired
    if (this._expired) throw new Error("OWA token expired. Open the addon options and click Connect to renew it.");
    // return the token
    return this._token;
  },

  // Flag the current token as expired and persist it.
  async markExpired() {
    // flag the current token as expired
    this._expired = true;
    // persist the expired flag
    await browser.storage.local.set({ [this._expKey()]: true });
  },

  // Harvest refresh token from a token endpoint request body (fallback when
  // filterResponseData is unavailable — reads the refresh_token grant's
  // request body instead of the response).
  async harvestRefreshToken(formData, url) {
    // skip when we're making our own refresh request
    if (this._selfRefreshActive) return false;
    // extract refresh_token from form data
    const rt = formData && formData.refresh_token && formData.refresh_token[0];
    // ignore requests without a refresh token
    if (!rt) return false;
    // extract client_id and scope from form data
    const cid = (formData.client_id && formData.client_id[0]) || null;
    const scope = (formData.scope && formData.scope[0]) || null;
    // delegate to harvestFromTokenResponse (no access token from request body)
    return this.harvestFromTokenResponse(rt, null, cid, url, scope, "webRequest-request");
  },

  // Add a refresh token to the collection. Dedup by token string.
  // webRequest-sourced tokens are pinned at the front (authoritative — reads
  // the actual Microsoft HTTP response). content-script tokens are appended.
  // Also sets the access token immediately when one is provided.
  // `source` is "webRequest", "webRequest-request", "content-script", or "legacy-migration".
  async harvestFromTokenResponse(refreshToken, accessToken, clientId, url, scope, source) {
    // skip when we're making our own refresh request
    if (this._selfRefreshActive) return false;
    // need a refresh token
    if (!refreshToken) return false;
    // log the incoming token details
    console.log("[M365OWA] harvestFromTokenResponse (" + (source || "?") + ")" +
      " start=" + refreshToken.slice(0, 25) +
      " end=..." + refreshToken.slice(-15) +
      " len=" + refreshToken.length +
      " client_id=" + (clientId || "?").slice(0, 12) +
      " url=" + (url || "?"));
    // dedup: check if this exact token string is already stored
    const existingIdx = this._refreshTokens.findIndex((e) => e.refreshToken === refreshToken);
    if (existingIdx !== -1) {
      // token already stored — update metadata (fill in missing fields)
      const e = this._refreshTokens[existingIdx];
      if (clientId && !e.clientId) e.clientId = clientId;
      if (url && !e.url) e.url = url;
      if (scope && !e.scope) e.scope = scope;
      e.ts = Date.now();
      await this._persistRefreshTokens();
      console.log("[M365OWA] refresh token already stored (dup, index=" + existingIdx + "), updated metadata; total: " + this._refreshTokens.length);
    } else {
      // new entry
      const entry = {
        refreshToken,
        clientId: clientId || null,
        url: url || null,
        scope: scope || null,
        source: source || "unknown",
        ts: Date.now(),
      };
      // webRequest tokens go to the front (authoritative); others appended
      if (source === "webRequest" || source === "webRequest-request") {
        this._refreshTokens.unshift(entry);
      } else {
        this._refreshTokens.push(entry);
      }
      await this._persistRefreshTokens();
      const insertedAt = (source === "webRequest" || source === "webRequest-request") ? 0 : this._refreshTokens.length - 1;
      console.log("[M365OWA] added refresh token (" + (source || "?") + ") at index " + insertedAt +
        ", total stored: " + this._refreshTokens.length);
    }
    // also set the access token if provided (immediately authenticates the addon)
    if (accessToken) {
      await this.setToken(accessToken);
    }
    return true;
  },

  // Transform v1-style scopes (lowercase) to v2-style (capitalized).
  // e.g., "https://outlook.office.com/calendars.readwrite" → "https://outlook.office.com/Calendars.ReadWrite"
  _transformScopeToV2(scope) {
    if (!scope) return null;
    return scope.split(" ").map(function(s) {
      var parts = s.split("/");
      var lastPart = parts[parts.length - 1];
      var permParts = lastPart.split(".").map(function(p) {
        return p.charAt(0).toUpperCase() + p.slice(1);
      });
      parts[parts.length - 1] = permParts.join(".");
      return parts.join("/");
    }).join(" ");
  },

  // Perform a single refresh-token POST to the given endpoint with the given params.
  // `entry` is the token entry providing client_id and refresh_token.
  // `params` may include: scope (v2), resource (v1), client_info, redirect_uri.
  // Returns { ok, data, reason }.
  async _doRefreshPost(url, params, entry) {
    try {
      const body = new URLSearchParams();
      body.set("grant_type", "refresh_token");
      body.set("client_id", entry.clientId);
      body.set("refresh_token", entry.refreshToken);
      if (params.scope) body.set("scope", params.scope);
      if (params.resource) body.set("resource", params.resource);
      if (params.client_info) body.set("client_info", "1");
      if (params.redirect_uri) body.set("redirect_uri", params.redirect_uri);
      console.log("[M365OWA] refresh-token POST " + url +
        " scope=" + (params.scope ? "yes" : "no") +
        " resource=" + (params.resource ? "yes" : "no") +
        " client_info=" + (params.client_info ? "yes" : "no") +
        " redirect_uri=" + (params.redirect_uri ? "yes" : "no") +
        " token_len=" + entry.refreshToken.length +
        " token_start=" + entry.refreshToken.slice(0, 15));
      const resp = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
        credentials: "omit",
      });
      if (resp.status !== 200) {
        const txt = await resp.text().catch(() => "");
        return { ok: false, reason: resp.status + ": " + txt.slice(0, 300) };
      }
      const data = await resp.json();
      if (!data.access_token) {
        return { ok: false, reason: "no access_token in response: " + JSON.stringify(data).slice(0, 200) };
      }
      return { ok: true, data };
    } catch (e) {
      return { ok: false, reason: (e && e.message) || String(e) };
    }
  },

  // Refresh the access token using stored refresh tokens; returns { ok, reason }.
  // Tries every stored token × every request strategy (endpoint/scope/resource/client_info).
  // First success wins. webRequest-sourced tokens are at the front of the array
  // (authoritative), so they are tried first.
  async refreshViaRefreshToken() {
    if (!this._refreshTokens.length) return { ok: false, reason: "no refresh tokens stored" };
    this._selfRefreshActive = true;
    try {
      // OWA resource URI used with the v1 token endpoint's `resource` parameter
      const OWA_RESOURCE = "https://outlook.office.com";
      let lastReason = "no strategy attempted";
      for (let ti = 0; ti < this._refreshTokens.length; ti++) {
        const entry = this._refreshTokens[ti];
        console.log("[M365OWA] refresh: trying token[" + ti + "] source=" + (entry.source || "?") +
          " start=" + entry.refreshToken.slice(0, 20) +
          " end=..." + entry.refreshToken.slice(-12) +
          " len=" + entry.refreshToken.length +
          " client_id=" + (entry.clientId || "?").slice(0, 8) +
          " url=" + (entry.url || "?"));
        // skip entries missing required fields
        if (!entry.url || !entry.clientId) {
          console.log("[M365OWA] refresh: token[" + ti + "] missing url or client_id, skipping");
          lastReason = "token[" + ti + "] missing url or client_id";
          continue;
        }
        // derive endpoint URLs from the stored URL
        // strip any query string (page-hook captures URLs with ?client-request-id=...)
        const baseUrl = entry.url.split("?")[0];
        // build the canonical OWA v2 token endpoint (what MSAL actually uses)
        const OWA_REDIRECT = "https://outlook.cloud.microsoft/mail/oauthRedirect.html";
        // the /organizations/ authority is what MSAL uses for multi-tenant public clients
        const orgsV2 = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token";
        const commonV2 = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        // also try the tenant-specific URL (from the entry's original URL)
        const tenantV2 = baseUrl;
        const v1Url = baseUrl.replace("/oauth2/v2.0/token", "/oauth2/token");
        const v1Common = "https://login.microsoftonline.com/common/oauth2/token";
        // also try the alternate authority hostname (login.windows.net ↔ login.microsoftonline.com)
        const altBaseUrl = baseUrl.indexOf("login.windows.net") !== -1
          ? baseUrl.replace("login.windows.net", "login.microsoftonline.com")
          : baseUrl.replace("login.microsoftonline.com", "login.windows.net");
        const altV1Url = altBaseUrl.replace("/oauth2/v2.0/token", "/oauth2/token");
        // transform the stored v1-style scope to v2-style (capitalized)
        const scopeV2 = this._transformScopeToV2(entry.scope);
        // Outlook-specific scope for OWA sync (what we actually need)
        const OWA_SCOPE = "https://outlook.office.com/Calendars.ReadWrite https://outlook.office.com/Contacts.ReadWrite openid profile offline_access";
        console.log("[M365OWA] refresh: token[" + ti + "] scope v1=" + (entry.scope ? entry.scope.slice(0, 80) : "null") + " v2=" + (scopeV2 ? scopeV2.slice(0, 80) : "null"));
        // build the list of strategies to try, in order:
        // Priority: match exactly what MSAL does — /organizations/ + redirect_uri + client_info=1
        // Then try variations with different scopes, endpoints, and authorities.
        const strategies = [
          // === EXACT MSAL pattern: /organizations/ + redirect_uri + client_info=1 ===
          { url: orgsV2,    scope: OWA_SCOPE, resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "orgs+redirect+clientinfo+owa-scope" },
          { url: orgsV2,    scope: null,      resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "orgs+redirect+clientinfo-noscope" },
          { url: orgsV2,    scope: scopeV2,   resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "orgs+redirect+clientinfo+scopeV2" },
          { url: orgsV2,    scope: "openid offline_access", resource: null, client_info: true, redirect_uri: OWA_REDIRECT, label: "orgs+redirect+clientinfo+minimal" },
          // === tenant-specific v2 + redirect_uri + client_info=1 ===
          { url: tenantV2,  scope: OWA_SCOPE, resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "tenant+redirect+clientinfo+owa-scope" },
          { url: tenantV2,  scope: null,      resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "tenant+redirect+clientinfo-noscope" },
          // === common v2 + redirect_uri + client_info=1 ===
          { url: commonV2,  scope: OWA_SCOPE, resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "common+redirect+clientinfo+owa-scope" },
          { url: commonV2,  scope: null,      resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "common+redirect+clientinfo-noscope" },
          // === v2 with redirect_uri but no client_info ===
          { url: orgsV2,    scope: OWA_SCOPE, resource: null, client_info: false, redirect_uri: OWA_REDIRECT, label: "orgs+redirect+owa-scope" },
          { url: tenantV2,  scope: OWA_SCOPE, resource: null, client_info: false, redirect_uri: OWA_REDIRECT, label: "tenant+redirect+owa-scope" },
          // === v2 without redirect_uri (simpler) ===
          { url: orgsV2,    scope: OWA_SCOPE, resource: null, client_info: false, redirect_uri: null, label: "orgs+owa-scope" },
          { url: orgsV2,    scope: null,      resource: null, client_info: false, redirect_uri: null, label: "orgs-noscope" },
          { url: tenantV2,  scope: null,      resource: null, client_info: false, redirect_uri: null, label: "tenant-noscope" },
          { url: tenantV2,  scope: scopeV2,   resource: null, client_info: false, redirect_uri: null, label: "tenant+scopeV2" },
          { url: commonV2,  scope: null,      resource: null, client_info: false, redirect_uri: null, label: "common-noscope" },
          // === v1 with resource (fallback) ===
          { url: v1Url,     scope: null,      resource: OWA_RESOURCE, client_info: false, redirect_uri: OWA_REDIRECT, label: "v1+redirect+resource" },
          { url: v1Url,     scope: null,      resource: OWA_RESOURCE, client_info: false, redirect_uri: null, label: "v1+resource" },
          { url: v1Common,  scope: null,      resource: OWA_RESOURCE, client_info: false, redirect_uri: null, label: "v1-common+resource" },
          // === alternate authority hostname ===
          { url: altBaseUrl, scope: OWA_SCOPE, resource: null, client_info: true,  redirect_uri: OWA_REDIRECT, label: "alt+redirect+clientinfo+owa-scope" },
          { url: altV1Url,  scope: null,      resource: OWA_RESOURCE, client_info: false, redirect_uri: null, label: "alt-v1+resource" },
        ];
        for (const s of strategies) {
          console.log("[M365OWA] refresh: token[" + ti + "] strategy: " + s.label);
          const r = await this._doRefreshPost(s.url, { scope: s.scope, resource: s.resource, client_info: s.client_info, redirect_uri: s.redirect_uri }, entry);
          if (r.ok) {
            await this.setToken(r.data.access_token);
            // update the refresh token if the server rotated it
            if (r.data.refresh_token && r.data.refresh_token !== entry.refreshToken) {
              entry.refreshToken = r.data.refresh_token;
              entry.ts = Date.now();
              await this._persistRefreshTokens();
              console.log("[M365OWA] refresh: token[" + ti + "] rotated (new start=" + r.data.refresh_token.slice(0, 20) + ")");
            }
            console.log("[M365OWA] access token refreshed (token[" + ti + "] strategy: " + s.label + ")");
            return { ok: true };
          }
          lastReason = "token[" + ti + "] " + s.label + " -> " + r.reason;
          console.log("[M365OWA] refresh: token[" + ti + "] " + s.label + " failed: " + r.reason);
        }
      }
      return { ok: false, reason: lastReason };
    } catch (e) {
      return { ok: false, reason: (e && e.message) || String(e) };
    } finally {
      this._selfRefreshActive = false;
    }
  },

  // Return true when at least one refresh token is stored.
  hasRefreshToken() { return this._refreshTokens.length > 0; },

  // Return true only when a token exists and is not flagged expired.
  isAuthenticated() { return !!this._token && !this._expired; },

  // Return the ms age of the current token (now minus capture time), or Infinity when unknown.
  tokenAgeMs() {
    // unknown when no capture timestamp exists
    if (!this._capturedAt) return Infinity;
    // return elapsed ms
    return Date.now() - this._capturedAt;
  },

  // Parse the JWT exp claim from the stored token; returns ms-epoch or null if not a parseable JWT.
  getTokenExpiry() {
    // no token means no expiry
    if (!this._token) return null;
    try {
      // split JWT into header.payload.signature
      const parts = this._token.split(".");
      // not a JWT if fewer than 2 parts
      if (parts.length < 2) return null;
      // decode the payload (base64url -> JSON)
      const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
      // return exp as ms epoch, or null if absent
      return payload.exp ? payload.exp * 1000 : null;
    } catch {
      // return null if parsing fails
      return null;
    }
  },

  // Build the OWA login URL from the configured host.
  owaLoginUrl() {
    // read the configured host (fall back to outlook default)
    const host = (CONFIG && CONFIG.OWA_HOST) || "https://outlook.office.com";
    // strip any trailing slashes
    const base = host.replace(/\/+$/, "");
    // return the OWA root URL
    return base + "/owa/";
  },

  // Return the tracked OWA tab id, verifying the tab still exists.
  async _liveOwaTabId() {
    // nothing tracked means no tab
    if (this._owaTabId == null) return null;
    try {
      // query the tracked tab
      await browser.tabs.get(this._owaTabId);
      // still alive, return it
      return this._owaTabId;
    } catch {
      // tab is gone, forget it
      this._owaTabId = null;
      // report no tab
      return null;
    }
  },

  // Open the OWA login page in a Thunderbird content tab (reusing the existing one) and return its id.
  async ensureOwaTab() {
    // check the tracked tab first
    const live = await this._liveOwaTabId();
    // reuse it when alive
    if (live != null) {
      // bring it to front
      await browser.tabs.update(live, { active: true });
      // return the reused id
      return live;
    }
    // build the OWA login URL
    const url = this.owaLoginUrl();
    // scan open tabs for an OWA tab to adopt
    const tabs = await browser.tabs.query({ url: ["https://outlook.office.com/owa/*", "https://outlook.office365.com/owa/*", "https://outlook.cloud.microsoft/owa/*"] });
    // adopt the first match when found
    if (tabs && tabs.length > 0) {
      // remember the adopted tab
      this._owaTabId = tabs[0].id;
      // bring it to front
      await browser.tabs.update(this._owaTabId, { active: true });
      // return the adopted id
      return this._owaTabId;
    }
    // create a fresh OWA content tab
    const created = await browser.tabs.create({ url, active: true });
    // remember the new tab
    this._owaTabId = created.id;
    console.log("[M365OWA] opened OWA login tab", this._owaTabId);
    // return the new id
    return this._owaTabId;
  },

  // Close the tracked OWA tab if it still exists.
  async closeOwaTab() {
    // check the tracked tab first
    const live = await this._liveOwaTabId();
    // forget the id either way
    this._owaTabId = null;
    // close it when alive
    if (live != null) {
      try { await browser.tabs.remove(live); } catch {}
    }
  },
};
