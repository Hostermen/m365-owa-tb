/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

var { ExtensionCommon: { ExtensionAPI } } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");

var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");

// List all mail accounts that could be used for OAuth2 token refresh.
// Returns an array of { id, hostname, username, type, name }
function listAllAccounts() {
  const results = [];
  try {
    const accounts = MailServices.accounts.accounts;
    if (!accounts) return results;
    for (const account of accounts) {
      const server = account.incomingServer;
      if (!server) continue;
      // Skip "none" type (Local Folders) and empty usernames
      const type = server.type || "";
      const username = server.username || "";
      const hostname = server.hostname || server.realHostName || "";
      if (type === "none" || type === "rss" || !username || !hostname) continue;
      results.push({
        id: account.key || "",
        hostname: hostname,
        username: username,
        type: type,
        name: server.prettyName || username,
      });
    }
  } catch (e) {
    console.error("[M365OWA] listAllAccounts error:", e.message || e);
  }
  return results;
}

// Find the OAuth2 server config from TB's existing mail accounts by matching email.
function detectAccountFromEmail(email) {
  if (!email) return null;
  const target = String(email).toLowerCase();
  try {
    const accounts = listAllAccounts();
    for (const a of accounts) {
      if (a.username && a.username.toLowerCase() === target) {
        console.log("[M365OWA] detectAccount: matched", a.hostname, "type", a.type, "user", a.username);
        return { hostname: a.hostname, username: a.username, type: a.type || "ews" };
      }
    }
    console.log("[M365OWA] detectAccount: no match for", email, "among", accounts.length, "accounts");
  } catch (e) {
    console.error("[M365OWA] detectAccount error:", e.message || e);
  }
  return null;
}

this.oauth = class extends ExtensionAPI {
  getAPI(context) {
    return {
      oauth: {
        // List all mail accounts available for OAuth2.
        async listAccounts() {
          return listAllAccounts();
        },

        async detectAccount(email) {
          return detectAccountFromEmail(email);
        },

        async getAccessToken(hostname, username, type = "ews") {
          try {
            // Auto-detect hostname if not provided
            if (!hostname) {
              const detected = detectAccountFromEmail(username);
              if (detected && detected.hostname) {
                hostname = detected.hostname;
                type = detected.type || type;
                console.log("[M365OWA] oauth: auto-detected", hostname, "type", type);
              } else {
                hostname = "outlook.office365.com";
                console.log("[M365OWA] oauth: no match, using default", hostname);
              }
            }

            const { OAuth2Module } = ChromeUtils.importESModule(
              "resource:///modules/OAuth2Module.sys.mjs"
            );
            const oauthModule = new OAuth2Module();
            const ok = oauthModule.initFromHostname(hostname, username, type);
            if (!ok) {
              throw new Error(`No OAuth2 provider configured for ${hostname} (${type})`);
            }

            const token = await new Promise((resolve, reject) => {
              const listener = {
                onSuccess(token) { resolve(token); },
                onFailure(code) {
                  reject(Components.Exception("OAuth2 token fetch failed", code));
                },
                QueryInterface: ChromeUtils.generateQI(["msgIOAuth2ModuleListener"]),
              };
              try {
                oauthModule.getAccessToken(listener);
              } catch (e) {
                reject(e);
              }
            });

            return {
              accessToken: token,
              username: username,
              hostname: hostname,
              type: type,
            };
          } catch (e) {
            console.error("[M365OWA] oauth.getAccessToken failed:", e.message || e);
            throw e;
          }
        },
      },
    };
  }
};
