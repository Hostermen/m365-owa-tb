// Options page logic. Talks to the background via browser.runtime messages.
const $ = (id) => document.getElementById(id);

async function send(msg) {
  return await browser.runtime.sendMessage(msg);
}

// --- Dark mode toggle (persisted in localStorage) ---
(function initTheme() {
  const root = document.documentElement;
  const saved = localStorage.getItem("m365-owa-theme");
  if (saved === "dark") root.setAttribute("data-theme", "dark");
  const toggle = $("themeToggle");
  if (toggle) {
    toggle.addEventListener("click", () => {
      const isDark = root.getAttribute("data-theme") === "dark";
      if (isDark) {
        root.removeAttribute("data-theme");
        localStorage.setItem("m365-owa-theme", "light");
      } else {
        root.setAttribute("data-theme", "dark");
        localStorage.setItem("m365-owa-theme", "dark");
      }
    });
  }
})();

function setErr(msg, cls) {
  const el = $("err");
  el.textContent = msg;
  el.className = cls || "ok";
}

function badge(cls, text) {
  const span = document.createElement("span");
  span.className = "status-badge " + cls;
  const dot = document.createElement("span");
  dot.className = "dot";
  span.appendChild(dot);
  span.appendChild(document.createTextNode(text));
  return span;
}

async function refreshStatus() {
  try {
    return await send({ type: "m365-owa-status" });
  } catch (e) {
    return null;
  }
}

function updateConnectBadge(status) {
  const el = $("connectStatus");
  if (!status) {
    el.innerHTML = "";
    return;
  }
  // offline takes priority over everything — show red even when the access
  // token is still valid in memory (sync can't run without a network).
  if (status.online === false) {
    el.replaceChildren(badge("danger", "Offline ✗"));
  } else if (status.authenticated) {
    el.replaceChildren(badge("ok", "Connected ✓"));
  } else if (status.configured) {
    // account is set up (refresh token present) but the access token is not
    // live yet — a background renewal is running or pending. Show an amber
    // "pending" badge instead of the grey "Not connected" one so the user
    // knows the addon is recovering on its own (e.g. cold boot, offline).
    if (status.renewing) {
      el.replaceChildren(badge("warn", "Reconnecting…"));
    } else if (status.retryPending) {
      el.replaceChildren(badge("warn", "Token expired — retrying…"));
    } else {
      el.replaceChildren(badge("warn", "Token expired — click Renew token or wait"));
    }
  } else if (status.owaTabOpen) {
    el.replaceChildren(badge("off", "Waiting for OWA login — complete it in the opened tab"));
  } else {
    el.replaceChildren(badge("off", "Not connected — click Connect to log in to OWA"));
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  if (window.particlesJS) {
    particlesJS("particles-js", {
      particles: {
        number: { value: 50, density: { enable: true, value_area: 900 } },
        color: { value: "#16a34a" },
        shape: { type: "circle" },
        opacity: { value: 0.35, random: true, anim: { enable: true, speed: 0.8, opacity_min: 0.1, sync: false } },
        size: { value: 4, random: true, anim: { enable: true, speed: 3, size_min: 0.3, sync: false } },
        line_linked: { enable: true, distance: 150, color: "#16a34a", opacity: 0.18, width: 1 },
        move: { enable: true, speed: 1.2, direction: "none", random: true, straight: false, out_mode: "out", bounce: false }
      },
      interactivity: {
        detect_on: "canvas",
        events: {
          onhover: { enable: true, mode: "grab" },
          onclick: { enable: true, mode: "push" },
          resize: true
        },
        modes: {
          grab: { distance: 180, line_linked: { opacity: 0.5 } },
          push: { particles_nb: 3 },
        }
      },
      retina_detect: true
    });
  }
  const s = await browser.storage.local.get([
    "pullDaysBack", "pullDaysForward",
  ]);
  if (s.pullDaysBack) $("pullDaysBack").value = s.pullDaysBack;
  if (s.pullDaysForward) $("pullDaysForward").value = s.pullDaysForward;

  await updateConnectButton();

  // auto-refresh the connect badge while the options page is open, so a
  // background renewal (e.g. after a cold boot) flips the badge to
  // "Connected" on its own without requiring a button click. Polls every 5s
  // while not authenticated, stops once connected. Stops after consecutive
  // failures (e.g. stale options tab after an addon update) to avoid
  // spamming Conduits errors — reopen the options page to restart polling.
  let failCount = 0;
  let badgePoll = setInterval(async () => {
    const s = await refreshStatus();
    if (!s) {
      // sendMessage failed — count and stop after 3 consecutive failures
      if (++failCount >= 3) {
        clearInterval(badgePoll);
        badgePoll = null;
      }
      return;
    }
    failCount = 0;
    applyStatus(s);
    // keep polling for the lifetime of the options page so both
    // online→offline and offline→online transitions update the badge
    // within ~5s without the user reopening the page. The sendMessage
    // is a cheap in-process call, so running it indefinitely is fine.
  }, 5000);

  // push-based badge updates: the background broadcasts connectivity
  // changes the instant they happen (window events + probe transitions),
  // so the badge reacts in real time instead of waiting for the 5s poll.
  browser.runtime.onMessage.addListener((msg) => {
    if (msg && msg.type === "m365-owa-connectivity") {
      refreshStatus().then((s) => { if (s) applyStatus(s); }).catch(() => {});
    }
    return false;
  });
});

// shared helper used by both the poll and the push listener
function applyStatus(s) {
  const btn = $("connectBtn");
  if (s.authenticated || s.configured) {
    btn.textContent = "Disconnect";
    btn.classList.remove("primary");
  } else {
    btn.textContent = "Connect";
    btn.classList.add("primary");
  }
  updateConnectBadge(s);
}

async function waitForAuth(timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const s = await refreshStatus();
    if (s && s.authenticated) return s;
    await new Promise((r) => setTimeout(r, 2000));
  }
  return await refreshStatus();
}

// --- Main: Connect / Disconnect ---
async function updateConnectButton() {
  const status = await refreshStatus();
  const btn = $("connectBtn");
  // Show "Disconnect" when the account is configured (a refresh token is
  // stored or a live access token exists) even if the access token is
  // currently expired/pending renewal — the user has set the account up, so
  // the button should not flip back to "Connect" on every cold boot.
  if (status && (status.authenticated || status.configured)) {
    btn.textContent = "Disconnect";
    btn.classList.remove("primary");
  } else {
    btn.textContent = "Connect";
    btn.classList.add("primary");
  }
  updateConnectBadge(status);
}

$("connectBtn").addEventListener("click", async () => {
  const btn = $("connectBtn");
  if (btn.textContent === "Disconnect") {
    await send({ type: "m365-owa-disconnect" });
    setErr("Disconnected.", "ok");
  } else {
    setErr("Opening OWA — log in in the opened tab…", "working");
    try {
      const r = await send({ type: "m365-owa-connect" });
      if (!r || !r.ok) { setErr("Could not open OWA tab.", "err"); return; }
      const s = await waitForAuth(120000);
      if (s && s.authenticated) {
        await send({ type: "m365-owa-relogin" });
        setErr("Connected ✓", "ok");
      } else {
        setErr("Still waiting for login — complete it in the OWA tab.", "err");
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
    pullDaysBack: $("pullDaysBack").value || 30,
    pullDaysForward: $("pullDaysForward").value || 90,
  });
  setErr(r && r.ok ? "Settings saved." : "Save failed.", r && r.ok ? "ok" : "err");
});

// --- Advanced: Diagnostics ---
$("forceRenew").addEventListener("click", async () => {
  setErr("Renewing token in background…", "working");
  try {
    const r = await send({ type: "m365-owa-debug-renew" });
    setErr(r && r.ok ? "Token renewed ✓" : "Renewal failed — " + ((r && r.reason) || "unknown"), r && r.ok ? "ok" : "err");
  } catch (e) {
    setErr("Renewal failed: " + (e.message || e), "err");
  }
  await updateConnectButton();
});

$("showStatus").addEventListener("click", async () => {
  const el = $("status");
  if (el.style.display !== "none") { el.style.display = "none"; return; }
  const s = await send({ type: "m365-owa-status" });
  el.textContent = JSON.stringify(s, null, 2);
  el.style.display = "block";
});
