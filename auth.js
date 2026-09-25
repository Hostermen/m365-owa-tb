// Bearer-token manager for OWA service.svc requests.
var Auth = {
  // current bearer string (without "Bearer " prefix)
  _token: null,
  // set true when OWA rejects with 401/403
  _expired: false,
  // setInterval handle for periodic OAuth2 refresh
  _refreshTimer: null,
  // whether automatic OAuth2 refresh is active
  _autoRefreshEnabled: false,
  // OAuth2 host (empty string = auto-detect)
  _oauthHostname: null,
  // account email used for OAuth2 refresh
  _oauthUsername: null,

  // storage key for the token
  _key() { return "m365_owa_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the expired flag
  _expKey() { return "m365_owa_expired_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the auto-refresh toggle
  _autoKey() { return "m365_owa_auto_refresh"; },
  // storage key for the OAuth2 hostname
  _oauthHostKey() { return "m365_owa_oauth_host"; },
  // storage key for the OAuth2 username
  _oauthUserKey() { return "m365_owa_oauth_user"; },
  // storage key for the OAuth2 account type
  _oauthTypeKey() { return "m365_owa_oauth_type"; },

  // Restore token, expired flag, and auto-refresh settings from storage.local into memory.
  async loadStoredToken() {
    // resolve the token storage key
    const k = this._key();
    // read the persisted token
    const { [k]: tok } = await browser.storage.local.get(k);
    // resolve the expired-flag storage key
    const ek = this._expKey();
    // read the persisted expired flag
    const { [ek]: exp } = await browser.storage.local.get(ek);
    // restore token into memory (or null)
    this._token = tok || null;
    // restore expired flag as boolean
    this._expired = !!exp;

    // batch-read all auto-refresh settings
    const cfg = await browser.storage.local.get([
      this._autoKey(), this._oauthHostKey(),
      this._oauthUserKey(), this._oauthTypeKey()
    ]);
    // restore auto-refresh toggle
    this._autoRefreshEnabled = !!cfg[this._autoKey()];
    // restore OAuth2 hostname
    this._oauthHostname = cfg[this._oauthHostKey()] || null;
    // restore OAuth2 username
    this._oauthUsername = cfg[this._oauthUserKey()] || null;
    // restore account type (default EWS)
    this._oauthType = cfg[this._oauthTypeKey()] || "ews";

    // start periodic refresh if it was enabled
    if (this._autoRefreshEnabled) {
      this._startRefreshTimer();
    }
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
    // persist token and clear expired flag
    await browser.storage.local.set({ [this._key()]: tok, [this._expKey()]: false });
    console.log("[M365OWA] token stored for connection", CONFIG.CONNECTION_NAME);
    // (re)start the refresh timer if enabled
    if (this._autoRefreshEnabled) this._startRefreshTimer();
  },

  // Clear the token and expired flag from memory and storage, and stop the refresh timer.
  async logout() {
    // clear the in-memory token
    this._token = null;
    // reset the expired flag
    this._expired = false;
    // stop any pending refresh timer
    this._stopRefreshTimer();
    // delete token and expired flag from storage
    await browser.storage.local.remove([this._key(), this._expKey()]);
  },

  // Enable/disable automatic token refresh via Thunderbird's OAuth2 module; persists the config and attempts an immediate refresh when enabled.
  async configureAutoRefresh(hostname, username, accountType) {
    // store the OAuth2 hostname
    this._oauthHostname = hostname || null;
    // store the OAuth2 username
    this._oauthUsername = username || null;
    // store the account type (default EWS)
    this._oauthType = accountType || "ews";
    // enable auto-refresh only when a username is present
    this._autoRefreshEnabled = !!this._oauthUsername;
    // persist all auto-refresh settings
    await browser.storage.local.set({
      [this._autoKey()]: this._autoRefreshEnabled,
      [this._oauthHostKey()]: hostname || "",
      [this._oauthUserKey()]: username || "",
      [this._oauthTypeKey()]: this._oauthType,
    });
    // when enabled, start timer and attempt an immediate refresh
    if (this._autoRefreshEnabled) {
      console.log("[M365OWA] auto-refresh configured for", username, hostname ? "at " + hostname : "(auto-detect)");
      this._startRefreshTimer();
      await this._tryAutoRefresh();
    } else {
      // when disabled, stop the timer
      this._stopRefreshTimer();
      console.log("[M365OWA] auto-refresh disabled");
    }
  },

  // Clear the pending refresh interval if one is running.
  _stopRefreshTimer() {
    // if a timer is running
    if (this._refreshTimer) {
      // cancel it
      clearTimeout(this._refreshTimer);
      // clear the handle
      this._refreshTimer = null;
    }
  },

  // Start (or restart) the 5-minute interval that triggers periodic OAuth2 token refresh.
  _startRefreshTimer() {
    // clear any existing timer first
    this._stopRefreshTimer();
    // schedule a refresh check every 5 minutes
    this._refreshTimer = setInterval(() => {
      this._tryAutoRefresh().catch(e => console.warn("[M365OWA] auto-refresh timer error:", e.message || e));
    }, 5 * 60 * 1000);
    console.log("[M365OWA] auto-refresh timer started (every 5 min)");
  },

  // Request a fresh access token from Thunderbird's OAuth2 module; on success, store and persist it. Returns true if a token was obtained.
  async _tryAutoRefresh() {
    // bail out if not configured
    if (!this._autoRefreshEnabled || !this._oauthUsername) return false;
    try {
      console.log("[M365OWA] auto-refresh: requesting token from TB OAuth2 for", this._oauthUsername);
      // ask Thunderbird's OAuth2 module for a fresh token
      const result = await messenger.oauth.getAccessToken(
        this._oauthHostname || "",
        this._oauthUsername,
        this._oauthType || "ews"
      );
      // if a token was returned
      if (result && result.accessToken) {
        // store it in memory
        this._token = result.accessToken;
        // mark as valid
        this._expired = false;
        // persist token and clear expired flag
        await browser.storage.local.set({ [this._key()]: result.accessToken, [this._expKey()]: false });
        console.log("[M365OWA] auto-refresh: got fresh token (len=" + result.accessToken.length + ")");
        return true;
      }
      console.warn("[M365OWA] auto-refresh: no token returned");
      return false;
    } catch (e) {
      console.warn("[M365OWA] auto-refresh failed:", e.message || e);
      return false;
    }
  },

  // Return the current token, refreshing it if expired and auto-refresh is enabled; throws if no token or refresh fails.
  async getTokenAsync() {
    // error if no token at all
    if (!this._token) throw new Error("No OWA token. Capture one (options page > Generate bookmarklet) or enable auto-refresh.");
    // if the token was marked expired
    if (this._expired) {
      // try auto-refresh first
      if (this._autoRefreshEnabled) {
        const ok = await this._tryAutoRefresh();
        if (!ok) throw new Error("OWA token expired and auto-refresh failed. Re-capture it or reconfigure auto-refresh.");
      } else {
        // otherwise tell the user to re-capture
        throw new Error("OWA token expired/rejected (401). Re-capture it from Outlook on the web.");
      }
    }
    // return the (possibly refreshed) token
    return this._token;
  },

  // Synchronous token accessor; throws if no token or expired (no auto-refresh attempt here).
  getToken() {
    // error if no token
    if (!this._token) throw new Error("No OWA token. Capture one (options page > Generate bookmarklet) and paste it.");
    // error if expired
    if (this._expired) throw new Error("OWA token expired/rejected (401). Re-capture it from Outlook on the web.");
    // return the token
    return this._token;
  },

  // Flag the current token as expired and persist it; attempts an immediate refresh if auto-refresh is enabled.
  async markExpired() {
    // flag the current token as expired
    this._expired = true;
    // persist the expired flag
    await browser.storage.local.set({ [this._expKey()]: true });
    // if auto-refresh is on, try to recover immediately
    if (this._autoRefreshEnabled) {
      const ok = await this._tryAutoRefresh();
      // clear the flag if a fresh token was obtained
      if (ok) this._expired = false;
    }
  },

  // Return true only when a token exists and is not flagged expired.
  isAuthenticated() { return !!this._token && !this._expired; },

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

  // Return whether automatic token refresh is active.
  isAutoRefreshEnabled() { return this._autoRefreshEnabled; },
  // Return a snapshot of the current OAuth2 config (hostname + username).
  getOAuthConfig() { return { hostname: this._oauthHostname || "(auto-detect)", username: this._oauthUsername }; },

  // Generate a javascript: URL bookmarklet that intercepts OWA's Authorization header and exposes the bearer token to the user.
  bookmarklet() {
    // build the bookmarklet source as an array of lines
    const src = [
      "(function(){",
      // guard against double-arming
      "if(window.__owaCap){alert('already armed - reload OWA and click again');return;}",
      "window.__owaCap=true;",
      // regex to extract the bearer token
      "function grab(v){var m=String(v||'').match(/Bearer\\s+([A-Za-z0-9._\\-]+)/);",
      // on match, disarm and capture
      "if(m){var t=m[1];window.__owaCap=false;",
      // create a textarea to enable clipboard copy
      "var h=document.createElement('textarea');h.value=t;document.body.appendChild(h);h.select();",
      // copy the token to the clipboard
      "try{document.execCommand('copy');}catch(e){}",
      // also show it in a prompt
      "prompt('M365 OWA bearer token (copied to clipboard). Paste into Thunderbird > M365 OWA Sync options:',t);",
      // clean up the textarea
      "h.remove();}}",
      // wrap fetch to intercept Authorization headers
      "var of=window.fetch;window.fetch=function(u,o){o=o||{};var h=o.headers||{};",
      // grab from plain-object headers
      "if(h.Authorization)grab(h.Authorization);",
      // or from Headers instance
      "else if(h instanceof Headers&&h.get&&h.get('Authorization'))grab(h.get('Authorization'));",
      // call the original fetch
      "return of.apply(this,arguments);};",
      // wrap setRequestHeader to intercept Authorization
      "var os=XMLHttpRequest.prototype.setRequestHeader;",
      "XMLHttpRequest.prototype.setRequestHeader=function(k,v){if(/authorization/i.test(k))grab(v);return os.apply(this,arguments);};",
      "alert('M365 OWA capture armed. Open or refresh Calendar/Mail, then the token prompt will appear.');",
      "})();"
    ].join("");
    // return as a javascript: URL
    return "javascript:" + encodeURIComponent(src);
  }
};
