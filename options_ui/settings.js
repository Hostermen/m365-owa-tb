// Options page logic. Talks to the background via browser.runtime messages.
const $ = (id) => document.getElementById(id);

async function send(msg) {
  return await browser.runtime.sendMessage(msg);
}

function setErr(msg, cls) {
  const el = $("err");
  el.textContent = msg;
  el.className = cls || "ok";
}

async function refreshStatus() {
  try {
    return await send({ type: "m365-owa-status" });
  } catch (e) {
    return null;
  }
}

function updateAutoRefreshBadge(status) {
  const el = $("autoRefreshStatus");
  if (!status) {
    el.innerHTML = "";
    return;
  }
  if (status.autoRefresh) {
    const cfg = status.oauthConfig || {};
    const user = cfg.username || "?";
    el.innerHTML = `<span class="status-badge ok"><span class="dot"></span>Auto-refresh active for ${user}</span>`;
  } else {
    el.innerHTML = `<span class="status-badge off"><span class="dot"></span>Auto-refresh disabled — select an account to enable</span>`;
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  const s = await browser.storage.local.get([
    "owaHost", "connectionName", "pullDaysBack", "pullDaysForward",
    "m365_owa_auto_refresh", "m365_owa_oauth_host", "m365_owa_oauth_user",
  ]);
  if (s.owaHost) $("owaHost").value = s.owaHost;
  if (s.connectionName) $("connectionName").value = s.connectionName;
  if (s.pullDaysBack) $("pullDaysBack").value = s.pullDaysBack;
  if (s.pullDaysForward) $("pullDaysForward").value = s.pullDaysForward;
  if (s.m365_owa_oauth_user) window._savedOAuthUser = s.m365_owa_oauth_user;

  await updateConnectButton();
  await loadAccounts();
});

async function updateBadge() {
  await updateConnectButton();
}

async function loadAccounts() {
  try {
    const accounts = await messenger.oauth.listAccounts();
    const sel = $("oauthAccountSelect");
    sel.innerHTML = "";
    if (!accounts || accounts.length === 0) {
      sel.innerHTML = '<option value="">No mail accounts found</option>';
    } else {
      sel.innerHTML = '<option value="">— Select an account —</option>';
      for (const a of accounts) {
        const label = `${a.name} (${a.username})`;
        const opt = document.createElement("option");
        opt.value = JSON.stringify({ hostname: a.hostname, username: a.username, type: a.type });
        opt.textContent = label;
        if (window._savedOAuthUser && a.username === window._savedOAuthUser) opt.selected = true;
        sel.appendChild(opt);
      }
    }
  } catch (e) {
    $("oauthAccountSelect").innerHTML = '<option value="">Error: ' + (e.message || e) + '</option>';
  }
}

async function getSelectedAccount() {
  const raw = $("oauthAccountSelect").value;
  if (!raw) { setErr("Select an account first.", "err"); return null; }
  try { return JSON.parse(raw); } catch { setErr("Invalid selection.", "err"); return null; }
}

// --- Main: Connect / Disconnect ---
async function updateConnectButton() {
  const status = await refreshStatus();
  const btn = $("connectBtn");
  if (status && status.autoRefresh) {
    btn.textContent = "Disconnect";
    btn.classList.remove("primary");
  } else {
    btn.textContent = "Connect";
    btn.classList.add("primary");
  }
  updateAutoRefreshBadge(status);
}

$("connectBtn").addEventListener("click", async () => {
  const btn = $("connectBtn");
  if (btn.textContent === "Disconnect") {
    await send({ type: "m365-owa-configure-auto-refresh", hostname: "", username: "" });
    setErr("Disconnected.", "ok");
  } else {
    const acct = await getSelectedAccount();
    if (!acct) return;
    setErr("Connecting…", "working");
    try {
      const r = await send({
        type: "m365-owa-configure-auto-refresh",
        hostname: acct.hostname, username: acct.username, accountType: acct.type,
      });
      if (r && r.ok) {
        await send({ type: "m365-owa-relogin" });
        setErr("Connected ✓", "ok");
      } else {
        setErr("Connection failed.", "err");
      }
    } catch (e) {
      setErr("Failed: " + (e.message || e), "err");
    }
  }
  await updateConnectButton();
});

// --- Main: Test connection ---
$("testConn").addEventListener("click", async () => {
  setErr("Testing…", "working");
  try {
    const s = await send({ type: "m365-owa-status" });
    if (s && s.owaProbe === "ok") setErr("Connection OK ✓", "ok");
    else setErr("Probe failed: " + (s && s.owaProbe), "err");
  } catch (e) {
    setErr("Test failed: " + (e.message || e), "err");
  }
});

// --- Main: Sync now (contacts + calendar) ---
$("syncAll").addEventListener("click", async () => {
  setErr("Syncing…", "working");
  try {
    const [c, cal] = await Promise.all([
      send({ type: "m365-owa-sync-contacts" }),
      send({ type: "m365-owa-sync-calendar" }),
    ]);
    const parts = [];
    if (c && c.ok) parts.push("contacts ✓");
    else parts.push("contacts ✗");
    if (cal && cal.ok) parts.push("calendar ✓");
    else parts.push("calendar ✗");
    setErr("Sync done — " + parts.join(", "), "ok");
  } catch (e) {
    setErr("Sync failed: " + (e.message || e), "err");
  }
});

// --- Advanced: Save config ---
$("saveConfig").addEventListener("click", async () => {
  const r = await send({
    type: "m365-owa-save-config",
    owaHost: $("owaHost").value.trim(),
    connectionName: $("connectionName").value.trim() || "default",
    pullDaysBack: $("pullDaysBack").value || 30,
    pullDaysForward: $("pullDaysForward").value || 90,
  });
  setErr(r && r.ok ? "Settings saved." : "Save failed.", r && r.ok ? "ok" : "err");
});

// --- Advanced: Manual token ---
$("genBookmarklet").addEventListener("click", async () => {
  const r = await send({ type: "m365-owa-bookmarklet" });
  if (!r || !r.url) { setErr("Could not generate bookmarklet.", "err"); return; }
  const out = $("bookmarkletOut");
  out.innerHTML = "";
  const a = document.createElement("a");
  a.className = "bm";
  a.href = r.url;
  a.textContent = "Capture M365 OWA token (drag to bookmarks bar)";
  out.appendChild(a);
});

$("saveToken").addEventListener("click", async () => {
  const tok = $("token").value.trim();
  if (!tok) { setErr("Paste a token first.", "err"); return; }
  try {
    await send({ type: "m365-owa-set-token", token: tok });
    await send({ type: "m365-owa-relogin" });
    setErr("Token saved & sync started ✓", "ok");
    $("token").value = "";
  } catch (e) {
    setErr("Token save failed: " + (e.message || e), "err");
  }
});

$("logout").addEventListener("click", async () => {
  await send({ type: "m365-owa-logout" });
  setErr("Token cleared.", "ok");
});

// --- Advanced: Diagnostics ---
$("showStatus").addEventListener("click", async () => {
  const el = $("status");
  if (el.style.display !== "none") { el.style.display = "none"; return; }
  const s = await send({ type: "m365-owa-status" });
  el.textContent = JSON.stringify(s, null, 2);
  el.style.display = "block";
});
