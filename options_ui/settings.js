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
  const el = $("autoRefreshStatus");
  if (!status) {
    el.innerHTML = "";
    return;
  }
  if (status.authenticated) {
    const age = status.tokenAgeSec;
    const ageMin = (typeof age === "number" && isFinite(age)) ? Math.floor(age / 60) : "?";
    el.replaceChildren(badge("ok", "Connected ✓ (token " + ageMin + " min old)"));
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
});

async function updateBadge() {
  await updateConnectButton();
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
  if (status && status.authenticated) {
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
