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
  // long-lived OAuth2 refresh token harvested from OWA's token endpoint calls (~90 days)
  _refreshToken: null,
  // OWA's public client_id used with the refresh token
  _refreshClientId: null,
  // full token endpoint URL (e.g. https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token)
  _refreshUrl: null,
  // scope string captured from OWA's token endpoint request body
  _refreshScope: null,
  // flag: true while we're making our own token refresh fetch (so the webRequest harvester skips us)
  _selfRefreshActive: false,

  // storage key for the token
  _key() { return "m365_owa_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the expired flag
  _expKey() { return "m365_owa_expired_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the capture timestamp
  _tsKey() { return "m365_owa_captured_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the refresh token
  _rkey() { return "m365_owa_refresh_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the refresh client_id
  _rckey() { return "m365_owa_refresh_clientid_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the refresh token endpoint URL
  _rukey() { return "m365_owa_refresh_url_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the refresh scope
  _rskey() { return "m365_owa_refresh_scope_" + (CONFIG.CONNECTION_NAME || "default"); },

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

  // Restore refresh token and related fields from storage.local into memory.
  async loadStoredRefreshToken() {
    // resolve the refresh-token storage key
    const rk = this._rkey();
    // read the persisted refresh token
    const { [rk]: rt } = await browser.storage.local.get(rk);
    // resolve the refresh client_id storage key
    const rck = this._rckey();
    // read the persisted client_id
    const { [rck]: rcid } = await browser.storage.local.get(rck);
    // resolve the refresh endpoint URL storage key
    const ru = this._rukey();
    // read the persisted endpoint URL
    const { [ru]: rurl } = await browser.storage.local.get(ru);
    // resolve the refresh scope storage key
    const rs = this._rskey();
    // read the persisted scope
    const { [rs]: rscope } = await browser.storage.local.get(rs);
    // restore into memory
    this._refreshToken = rt || null;
    this._refreshClientId = rcid || null;
    this._refreshUrl = rurl || null;
    this._refreshScope = rscope || null;
    // log when a refresh token was loaded
    if (this._refreshToken) console.log("[M365OWA] refresh token loaded (client_id=" + (this._refreshClientId || "?").slice(0, 8) + "...)");
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
    // clear refresh-token fields
    this._refreshToken = null;
    this._refreshClientId = null;
    this._refreshUrl = null;
    this._refreshScope = null;
    // delete token, flag, timestamp, and refresh-token data from storage
    await browser.storage.local.remove([this._key(), this._expKey(), this._tsKey(), this._rkey(), this._rckey(), this._rukey(), this._rskey()]);
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

  // Harvest refresh token + client_id + endpoint URL + scope from a token endpoint request body.
  async harvestRefreshToken(formData, url) {
    // skip when we're making our own refresh request
    if (this._selfRefreshActive) return false;
    // extract refresh_token from form data
    const rt = formData && formData.refresh_token && formData.refresh_token[0];
    // ignore requests without a refresh token
    if (!rt) return false;
    // extract client_id from form data
    const cid = (formData.client_id && formData.client_id[0]) || this._refreshClientId;
    // extract scope from form data
    const scope = (formData.scope && formData.scope[0]) || this._refreshScope;
    // store in memory
    this._refreshToken = rt;
    this._refreshClientId = cid;
    this._refreshUrl = url;
    this._refreshScope = scope;
    // persist all four fields
    await browser.storage.local.set({
      [this._rkey()]: rt,
      [this._rckey()]: cid,
      [this._rukey()]: url,
      [this._rskey()]: scope,
    });
    console.log("[M365OWA] harvested refresh token (client_id=" + (cid || "?").slice(0, 8) + "..., url=" + url + ")");
    // report that a refresh token was captured
    return true;
  },

  // Harvest refresh token + access token from a token endpoint RESPONSE body.
  // This captures refresh tokens from ALL grant types (auth_code, refresh_token)
  // by reading the response instead of the request.
  async harvestFromTokenResponse(refreshToken, accessToken, clientId, url, scope) {
    // skip when we're making our own refresh request
    if (this._selfRefreshActive) return false;
    // need a refresh token from the response
    if (!refreshToken) return false;
    // store the refresh token and related fields
    this._refreshToken = refreshToken;
    if (clientId) this._refreshClientId = clientId;
    if (url) this._refreshUrl = url;
    if (scope) this._refreshScope = scope;
    // persist all four fields
    await browser.storage.local.set({
      [this._rkey()]: refreshToken,
      [this._rckey()]: this._refreshClientId,
      [this._rukey()]: this._refreshUrl,
      [this._rskey()]: this._refreshScope,
    });
    console.log("[M365OWA] harvested refresh token from token endpoint response (client_id=" + (this._refreshClientId || "?").slice(0, 8) + "...)");
    // also set the access token if provided (immediately authenticates the addon)
    if (accessToken) {
      await this.setToken(accessToken);
    }
    return true;
  },

  // Refresh the access token using the stored refresh token; returns { ok, reason }.
  async refreshViaRefreshToken() {
    // need all three: refresh token, endpoint URL, and client_id
    if (!this._refreshToken) return { ok: false, reason: "no refresh token stored" };
    if (!this._refreshUrl) return { ok: false, reason: "no token endpoint URL stored" };
    if (!this._refreshClientId) return { ok: false, reason: "no client_id stored" };
    // mark that we're making our own request so the webRequest harvester skips
    this._selfRefreshActive = true;
    try {
      // build the form body
      const body = new URLSearchParams();
      body.set("grant_type", "refresh_token");
      body.set("client_id", this._refreshClientId);
      body.set("refresh_token", this._refreshToken);
      body.set("client_info", "1");
      // DO NOT send scope for refresh_token grant - Azure AD v2 returns the original scopes
      // log the request params (without the token itself) for debugging
      console.log("[M365OWA] refresh-token renewal: POST " + this._refreshUrl +
        " client_id=" + (this._refreshClientId || "?").slice(0, 8) + "..." +
        " no_scope token_len=" + (this._refreshToken ? this._refreshToken.length : 0) +
        " token_start=" + (this._refreshToken ? this._refreshToken.slice(0, 15) : "") +
        " token_end=" + (this._refreshToken ? this._refreshToken.slice(-15) : ""));
      // POST to the token endpoint (no cookies, no SSO dependency)
      const resp = await fetch(this._refreshUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
        credentials: "omit",
      });
      // non-200 means the refresh token is likely expired or revoked
      if (resp.status !== 200) {
        const txt = await resp.text().catch(() => "");
        return { ok: false, reason: "token endpoint returned " + resp.status + ": " + txt.slice(0, 400) };
      }
      const data = await resp.json();
      if (!data.access_token) return { ok: false, reason: "no access_token in response" };
      // store the new access token
      await this.setToken(data.access_token);
      // update the refresh token if a new one was returned (rotation)
      if (data.refresh_token && data.refresh_token !== this._refreshToken) {
        this._refreshToken = data.refresh_token;
        await browser.storage.local.set({ [this._rkey()]: data.refresh_token });
        console.log("[M365OWA] refresh token updated from token endpoint response");
      }
      console.log("[M365OWA] access token refreshed via refresh token");
      return { ok: true };
    } catch (e) {
      return { ok: false, reason: (e && e.message) || String(e) };
    } finally {
      this._selfRefreshActive = false;
    }
  },

  // Return true when a refresh token is stored.
  hasRefreshToken() { return !!this._refreshToken; },

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
