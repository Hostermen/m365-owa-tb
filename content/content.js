// Content script (isolated world): injects page-hook.js into the page context
// and relays token endpoint response data to the background script.
(function() {
  // Inject the page hook script into the page's main world
  try {
    var script = document.createElement("script");
    script.src = browser.runtime.getURL("content/page-hook.js");
    (document.head || document.documentElement).appendChild(script);
    script.remove();
  } catch (e) {
    console.warn("[M365OWA] content script: failed to inject page hook:", e.message || e);
  }

  // Relay token responses from the page world to the background
  window.addEventListener("message", function(event) {
    if (event.source !== window) return;
    if (event.data && event.data.type === "M365OWA_TOKEN_RESPONSE") {
      browser.runtime.sendMessage({
        type: "m365-owa-refresh-token-harvest",
        refresh_token: event.data.refresh_token,
        access_token: event.data.access_token,
        client_id: event.data.client_id,
        scope: event.data.scope,
        url: event.data.url,
      }).catch(function() {});
    }
  });
})();
