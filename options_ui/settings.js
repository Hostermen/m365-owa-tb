// Options page logic. Talks to the background via browser.runtime messages.
const $ = (id) => document.getElementById(id);

async function send(msg) {
  return await browser.runtime.sendMessage(msg);
}

async function refreshStatus() {
  try {
    const s = await send({ type: "m365-owa-status" });
    $("status").textContent = JSON.stringify(s, null, 2);
  } catch (e) {
    $("status").textContent = "status error: " + (e.message || e);
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  // prefill from storage.local
  const s = await browser.storage.local.get([
    "owaHost", "connectionName", "pullDaysBack", "pullDaysForward",
    "m365_owa_auto_refresh", "m365_owa_oauth_host", "m365_owa_oauth_user",
  ]);
  if (s.owaHost) $("owaHost").value = s.owaHost;
  if (s.connectionName) $("connectionName").value = s.connectionName;
  if (s.pullDaysBack) $("pullDaysBack").value = s.pullDaysBack;
  if (s.pullDaysForward) $("pullDaysForward").value = s.pullDaysForward;
  if (s.m365_owa_oauth_user) {
    // Will be selected after accounts load
    window._savedOAuthUser = s.m365_owa_oauth_user;
  }
  if (s.m365_owa_auto_refresh) {
    $("autoRefreshStatus").textContent = "Auto-refresh ENABLED for " + (s.m365_owa_oauth_user || "?");
  }
  await refreshStatus();

  // Load available accounts into dropdown
  try {
    const accounts = await messenger.oauth.listAccounts();
    const sel = $("oauthAccountSelect");
    sel.innerHTML = "";
    if (!accounts || accounts.length === 0) {
      sel.innerHTML = '<option value="">No mail accounts found</option>';
    } else {
      sel.innerHTML = '<option value="">— Select an account —</option>';
      for (const a of accounts) {
        const label = `${a.name} (${a.username}) — ${a.type} @ ${a.hostname}`;
        const opt = document.createElement("option");
        opt.value = JSON.stringify({ hostname: a.hostname, username: a.username, type: a.type });
        opt.textContent = label;
        if (window._savedOAuthUser && a.username === window._savedOAuthUser) {
          opt.selected = true;
        }
        sel.appendChild(opt);
      }
    }
  } catch (e) {
    $("oauthAccountSelect").innerHTML = '<option value="">Error loading accounts: ' + (e.message || e) + '</option>';
  }
});

$("saveConfig").addEventListener("click", async () => {
  const r = await send({
    type: "m365-owa-save-config",
    owaHost: $("owaHost").value.trim(),
    connectionName: $("connectionName").value.trim() || "default",
    pullDaysBack: $("pullDaysBack").value || 30,
    pullDaysForward: $("pullDaysForward").value || 90,
  });
  $("err").textContent = r && r.ok ? "Settings saved." : "Save failed.";
  await refreshStatus();
});

$("genBookmarklet").addEventListener("click", async () => {
  const r = await send({ type: "m365-owa-bookmarklet" });
  if (!r || !r.url) { $("err").textContent = "Could not generate bookmarklet."; return; }
  const out = $("bookmarkletOut");
  out.innerHTML = "";
  const a = document.createElement("a");
  a.className = "bm";
  a.href = r.url;
  a.textContent = "Capture M365 OWA token  (drag me to your bookmarks bar)";
  out.appendChild(a);
  const p = document.createElement("p");
  p.className = "hint";
  p.textContent = "Drag the link above to your bookmarks bar. Then open Outlook on the web and click it.";
  out.appendChild(p);
});

$("saveToken").addEventListener("click", async () => {
  const tok = $("token").value.trim();
  if (!tok) { $("err").textContent = "Paste a token first."; return; }
  try {
    await send({ type: "m365-owa-set-token", token: tok });
    await send({ type: "m365-owa-relogin" });
    $("err").textContent = "Token saved & sync started.";
    $("token").value = "";
    await refreshStatus();
  } catch (e) {
    $("err").textContent = "Token save failed: " + (e.message || e);
  }
});

$("logout").addEventListener("click", async () => {
  await send({ type: "m365-owa-logout" });
  $("err").textContent = "Token cleared.";
  await refreshStatus();
});

$("saveAutoRefresh").addEventListener("click", async () => {
  const sel = $("oauthAccountSelect");
  const raw = sel.value;
  if (!raw) {
    $("err").textContent = "Select an account from the dropdown.";
    return;
  }
  let acct;
  try { acct = JSON.parse(raw); } catch { $("err").textContent = "Invalid selection."; return; }
  try {
    const r = await send({
      type: "m365-owa-configure-auto-refresh",
      hostname: acct.hostname,
      username: acct.username,
      accountType: acct.type,
    });
    if (r && r.ok) {
      $("autoRefreshStatus").textContent = "Auto-refresh enabled for " + acct.username + ". Token will be refreshed automatically.";
      $("err").textContent = "Auto-refresh configured. Testing token fetch…";
      await send({ type: "m365-owa-relogin" });
      $("err").textContent = "Auto-refresh active ✓";
    } else {
      $("err").textContent = "Auto-refresh setup failed.";
    }
  } catch (e) {
    $("err").textContent = "Auto-refresh failed: " + (e.message || e);
  }
  await refreshStatus();
});

$("disableAutoRefresh").addEventListener("click", async () => {
  const r = await send({ type: "m365-owa-configure-auto-refresh", hostname: "", username: "" });
  $("autoRefreshStatus").textContent = "Auto-refresh disabled.";
  $("err").textContent = r && r.ok ? "Auto-refresh disabled." : "Failed.";
  await refreshStatus();
});

$("testConn").addEventListener("click", async () => {
  $("err").textContent = "Testing…";
  try {
    const s = await send({ type: "m365-owa-status" });
    $("err").textContent = s && s.owaProbe === "ok" ? "Connection OK ✓" : "Probe failed: " + (s && s.owaProbe);
  } catch (e) {
    $("err").textContent = "Test failed: " + (e.message || e);
  }
});

$("syncContacts").addEventListener("click", async () => {
  $("err").textContent = "Syncing contacts…";
  const r = await send({ type: "m365-owa-sync-contacts" });
  $("err").textContent = r && r.ok ? "Contacts sync done." : "Contacts sync failed (check error console).";
  await refreshStatus();
});

$("syncCalendar").addEventListener("click", async () => {
  $("err").textContent = "Syncing calendar…";
  const r = await send({ type: "m365-owa-sync-calendar" });
  $("err").textContent = r && r.ok ? "Calendar sync done." : "Calendar sync failed.";
  await refreshStatus();
});
