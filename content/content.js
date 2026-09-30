// Content script (isolated world): patches fetch() and XMLHttpRequest in the
// page world via wrappedJSObject/exportFunction to intercept token endpoint
// responses and relay refresh tokens to the background script.
(function() {
  console.log("[M365OWA] content script loaded on:", location.href);

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

  // Relay captured token data to the background script
  function relayToBackground(refreshToken, accessToken, clientId, scope, url) {
    try {
      browser.runtime.sendMessage({
        type: "m365-owa-refresh-token-harvest",
        refresh_token: refreshToken,
        access_token: accessToken || null,
        client_id: clientId || null,
        scope: scope || null,
        url: url || null,
      }).catch(function() {});
    } catch (e) {
      console.warn("[M365OWA] content: relay failed:", e.message || e);
    }
  }

  // Access the page world via wrappedJSObject (Firefox/Thunderbird XPCNativeWrapper)
  var pageWindow = window.wrappedJSObject;

  // --- Patch fetch() in the page world ---
  var originalFetch = pageWindow.fetch;
  console.log("[M365OWA] content: patching fetch, original typeof:", typeof originalFetch);

  var patchedFetch = exportFunction(function() {
    var args = arguments;
    var url = typeof args[0] === "string" ? args[0] : (args[0] && args[0].url) || "";
    var promise = originalFetch.apply(this, args);
    if (isTokenEndpoint(url)) {
      console.log("[M365OWA] content: fetch to token endpoint detected:", url);
      promise.then(function(response) {
        try {
          var bodyStr = typeof args[1] === "object" && args[1] ? args[1].body : null;
          var meta = parseFormBody(bodyStr);
          var clone = response.clone();
          clone.json().then(function(data) {
            console.log("[M365OWA] content: token endpoint response keys:", Object.keys(data).join(","));
            if (data.refresh_token) {
              console.log("[M365OWA] content: refresh_token found in fetch response!");
              relayToBackground(data.refresh_token, data.access_token, meta.clientId, meta.scope, url);
            }
          }).catch(function(e) {
            console.warn("[M365OWA] content: failed to parse fetch response:", e.message || e);
          });
        } catch (e) {}
      }).catch(function() {});
    }
    return promise;
  }, window);

  pageWindow.fetch = patchedFetch;

  // --- Patch XMLHttpRequest in the page world ---
  var OriginalXHR = pageWindow.XMLHttpRequest;
  var originalOpen = OriginalXHR.prototype.open;
  var originalSend = OriginalXHR.prototype.send;

  var patchedOpen = exportFunction(function(method, url) {
    this._m365owa_url = url;
    return originalOpen.apply(this, arguments);
  }, window);

  var patchedSend = exportFunction(function(body) {
    var self = this;
    if (isTokenEndpoint(this._m365owa_url)) {
      console.log("[M365OWA] content: XHR to token endpoint detected:", this._m365owa_url);
      this.addEventListener("load", function() {
        try {
          var data = JSON.parse(self.responseText);
          console.log("[M365OWA] content: XHR token endpoint response keys:", Object.keys(data).join(","));
          if (data.refresh_token) {
            console.log("[M365OWA] content: refresh_token found in XHR response!");
            var meta = parseFormBody(body);
            relayToBackground(data.refresh_token, data.access_token, meta.clientId, meta.scope, self._m365owa_url);
          }
        } catch (e) {
          console.warn("[M365OWA] content: failed to parse XHR response:", e.message || e);
        }
      });
    }
    return originalSend.apply(this, arguments);
  }, window);

  OriginalXHR.prototype.open = patchedOpen;
  OriginalXHR.prototype.send = patchedSend;

  console.log("[M365OWA] content: fetch and XHR patched in page world");
})();
