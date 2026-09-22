// OWA bearer-token auth.
//
// Ported from m365-owa-cli's auth model: there is NO OAuth and NO client_id.
// The user captures an OWA bearer token from a logged-in Outlook on the web
// tab (bookmarklet or browser DevTools) and pastes it here. We store the raw
// token in browser.storage.local and attach it as `Authorization: Bearer ...`
// to every OWA service.svc request. When OWA returns 401/403 we mark the token
// expired so the options page can tell the user to re-capture.
//
// AUTO-REFRESH: If the user has a Thunderbird mail account configured with
// OAuth2 (IMAP or EWS), the addon can automatically get fresh access tokens
// via the `messenger.oauth.getAccessToken` experiment API. This uses TB's
// built-in OAuth2 module which stores refresh tokens in the login manager.
var Auth = {
  _token: null,        // raw bearer (no "Bearer " prefix)
  _expired: false,
  _refreshTimer: null,
  _autoRefreshEnabled: false,
  _oauthHostname: null,
  _oauthUsername: null,

  _key() { return "m365_owa_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  _expKey() { return "m365_owa_expired_" + (CONFIG.CONNECTION_NAME || "default"); },
  _autoKey() { return "m365_owa_auto_refresh"; },
  _oauthHostKey() { return "m365_owa_oauth_host"; },
  _oauthUserKey() { return "m365_owa_oauth_user"; },
  _oauthTypeKey() { return "m365_owa_oauth_type"; },

  async loadStoredToken() {
    const k = this._key();
    const { [k]: tok } = await browser.storage.local.get(k);
    const ek = this._expKey();
    const { [ek]: exp } = await browser.storage.local.get(ek);
    this._token = tok || null;
    this._expired = !!exp;

    // Load auto-refresh config
    const cfg = await browser.storage.local.get([this._autoKey(), this._oauthHostKey(), this._oauthUserKey(), this._oauthTypeKey()]);
    this._autoRefreshEnabled = !!cfg[this._autoKey()];
    this._oauthHostname = cfg[this._oauthHostKey()] || null;
    this._oauthUsername = cfg[this._oauthUserKey()] || null;
    this._oauthType = cfg[this._oauthTypeKey()] || "ews";

    if (this._autoRefreshEnabled) {
      this._startRefreshTimer();
    }
  },

  async setToken(raw) {
    if (!raw) throw new Error("Empty token");
    // Accept either the raw bearer or a full "Authorization: Bearer xxx" header,
    // or an MSAL/id-token response JSON. We only want the access_token.
    let tok = String(raw).trim();
    const m = tok.match(/Bearer\s+([A-Za-z0-9._\-]+)/i);
    if (m) tok = m[1];
    else {
      try { const j = JSON.parse(tok); if (j && j.access_token) tok = j.access_token; } catch {}
    }
    this._token = tok;
    this._expired = false;
    await browser.storage.local.set({ [this._key()]: tok, [this._expKey()]: false });
    console.log("[M365OWA] token stored for connection", CONFIG.CONNECTION_NAME);
    if (this._autoRefreshEnabled) this._startRefreshTimer();
  },

  async logout() {
    this._token = null;
    this._expired = false;
    this._stopRefreshTimer();
    await browser.storage.local.remove([this._key(), this._expKey()]);
  },

  // Configure auto-refresh using Thunderbird's built-in OAuth2.
  // hostname: e.g. "outlook.office365.com"
  // username: e.g. "user@example.com"
  async configureAutoRefresh(hostname, username, accountType) {
    this._oauthHostname = hostname || null;
    this._oauthUsername = username || null;
    this._oauthType = accountType || "ews";
    this._autoRefreshEnabled = !!this._oauthUsername;
    await browser.storage.local.set({
      [this._autoKey()]: this._autoRefreshEnabled,
      [this._oauthHostKey()]: hostname || "",
      [this._oauthUserKey()]: username || "",
      [this._oauthTypeKey()]: this._oauthType,
    });
    if (this._autoRefreshEnabled) {
      console.log("[M365OWA] auto-refresh configured for", username, hostname ? "at " + hostname : "(auto-detect)");
      this._startRefreshTimer();
      // Try an immediate refresh
      await this._tryAutoRefresh();
    } else {
      this._stopRefreshTimer();
      console.log("[M365OWA] auto-refresh disabled");
    }
  },

  _stopRefreshTimer() {
    if (this._refreshTimer) {
      clearTimeout(this._refreshTimer);
      this._refreshTimer = null;
    }
  },

  _startRefreshTimer() {
    this._stopRefreshTimer();
    // Check every 5 minutes; the experiment will use the stored refresh token
    // to get a fresh access token, which is cheap (no UI).
    this._refreshTimer = setInterval(() => {
      this._tryAutoRefresh().catch(e => console.warn("[M365OWA] auto-refresh timer error:", e.message || e));
    }, 5 * 60 * 1000);
    console.log("[M365OWA] auto-refresh timer started (every 5 min)");
  },

  // Try to get a fresh token from TB's OAuth2 system.
  // Returns true if a new token was obtained.
  async _tryAutoRefresh() {
    if (!this._autoRefreshEnabled || !this._oauthUsername) return false;
    try {
      console.log("[M365OWA] auto-refresh: requesting token from TB OAuth2 for", this._oauthUsername);
      // Pass empty hostname — the experiment auto-detects from TB's accounts
      const result = await messenger.oauth.getAccessToken(
        this._oauthHostname || "",
        this._oauthUsername,
        this._oauthType || "ews"
      );
      if (result && result.accessToken) {
        this._token = result.accessToken;
        this._expired = false;
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

  // Get the current token, auto-refreshing if needed.
  async getTokenAsync() {
    if (!this._token) throw new Error("No OWA token. Capture one (options page > Generate bookmarklet) or enable auto-refresh.");
    if (this._expired) {
      // Try auto-refresh before giving up
      if (this._autoRefreshEnabled) {
        const ok = await this._tryAutoRefresh();
        if (!ok) throw new Error("OWA token expired and auto-refresh failed. Re-capture it or reconfigure auto-refresh.");
      } else {
        throw new Error("OWA token expired/rejected (401). Re-capture it from Outlook on the web.");
      }
    }
    return this._token;
  },

  // Synchronous version for backward compatibility (calendar code uses this).
  getToken() {
    if (!this._token) throw new Error("No OWA token. Capture one (options page > Generate bookmarklet) and paste it.");
    if (this._expired) throw new Error("OWA token expired/rejected (401). Re-capture it from Outlook on the web.");
    return this._token;
  },

  async markExpired() {
    this._expired = true;
    await browser.storage.local.set({ [this._expKey()]: true });
    // If auto-refresh is on, try to get a new token immediately
    if (this._autoRefreshEnabled) {
      const ok = await this._tryAutoRefresh();
      if (ok) this._expired = false;
    }
  },

  isAuthenticated() { return !!this._token && !this._expired; },

  // Parse the JWT exp timestamp (seconds since epoch) from the stored token.
  // Returns null if the token is not a JWT or can't be parsed.
  getTokenExpiry() {
    if (!this._token) return null;
    try {
      const parts = this._token.split(".");
      if (parts.length < 2) return null;
      const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
      return payload.exp ? payload.exp * 1000 : null;
    } catch { return null; }
  },

  isAutoRefreshEnabled() { return this._autoRefreshEnabled; },
  getOAuthConfig() { return { hostname: this._oauthHostname || "(auto-detect)", username: this._oauthUsername }; },

  // --- Bookmarklet generator (mirrors `m365-owa-cli auth bookmarklet`) ---
  //
  // Returns a `javascript:` URL the user can save as a bookmark. When clicked
  // while Outlook on the web is open, it monkey-patches fetch + XHR to capture
  // the next `Authorization: Bearer ...` header OWA sends, then shows it via
  // prompt() for copying into the addon options page.
  bookmarklet() {
    const src = [
      "(function(){",
      "if(window.__owaCap){alert('already armed - reload OWA and click again');return;}",
      "window.__owaCap=true;",
      "function grab(v){var m=String(v||'').match(/Bearer\\s+([A-Za-z0-9._\\-]+)/);",
      "if(m){var t=m[1];window.__owaCap=false;",
      "var h=document.createElement('textarea');h.value=t;document.body.appendChild(h);h.select();",
      "try{document.execCommand('copy');}catch(e){}",
      "prompt('M365 OWA bearer token (copied to clipboard). Paste into Thunderbird > M365 OWA Sync options:',t);",
      "h.remove();}}",
      "var of=window.fetch;window.fetch=function(u,o){o=o||{};var h=o.headers||{};",
      "if(h.Authorization)grab(h.Authorization);",
      "else if(h instanceof Headers&&h.get&&h.get('Authorization'))grab(h.get('Authorization'));",
      "return of.apply(this,arguments);};",
      "var os=XMLHttpRequest.prototype.setRequestHeader;",
      "XMLHttpRequest.prototype.setRequestHeader=function(k,v){if(/authorization/i.test(k))grab(v);return os.apply(this,arguments);};",
      "alert('M365 OWA capture armed. Open or refresh Calendar/Mail, then the token prompt will appear.');",
      "})();"
    ].join("");
    return "javascript:" + encodeURIComponent(src);
  }
};
