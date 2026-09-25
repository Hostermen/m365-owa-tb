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

  // storage key for the token
  _key() { return "m365_owa_token_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the expired flag
  _expKey() { return "m365_owa_expired_" + (CONFIG.CONNECTION_NAME || "default"); },
  // storage key for the capture timestamp
  _tsKey() { return "m365_owa_captured_" + (CONFIG.CONNECTION_NAME || "default"); },

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
    // flag tokens older than the max age as expired so they get renewed
    if (this._token && (Date.now() - this._capturedAt) > this._maxTokenAgeMs) {
      // mark stale in memory
      this._expired = true;
      // persist the stale flag
      await browser.storage.local.set({ [this._expKey()]: true });
    }
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
    // delete token, flag, and timestamp from storage
    await browser.storage.local.remove([this._key(), this._expKey(), this._tsKey()]);
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

  // Reload the OWA tab so OWA mints a fresh token that gets harvested; returns true when reloaded.
  async reloadOwaTab() {
    // check the tracked tab first
    const live = await this._liveOwaTabId();
    // bail out when no OWA tab is open
    if (live == null) return false;
    // reload it without activating
    await browser.tabs.reload(live);
    console.log("[M365OWA] reloaded OWA tab to renew token");
    // report success
    return true;
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
