// File watcher for Thunderbird addon development.
// Watches source files for changes and (optionally) auto-reloads the addon
// via Thunderbird's remote debugging protocol.
//
// Usage:
//   node dev-watch.js            # watch + rebuild xpi + notify
//   node dev-watch.js --reload   # watch + rebuild + auto-reload via RDP
//
// Reload workflow (no --reload flag):
//   1. Thunderbird > Tools > Developer Tools > Debug Add-ons (about:debugging)
//   2. "Load Temporary Add-on" > select manifest.json (loads from source dir)
//   3. After each file change this script rebuilds the xpi and logs.
//   4. Reload: click "Reload" in about:debugging, or run M365OWA.reload()
//      in the Browser Toolbox console (Tools > Developer Tools > Browser Toolbox).
const fs = require("fs");
const path = require("path");
const { spawn, execSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const XPI = path.join(ROOT, "m365-owa-tb.xpi");

// extensions to watch
const EXTS = new Set([".js", ".json", ".html", ".svg", ".png", ".css"]);
// paths to ignore
const IGNORE = ["node_modules", ".git", "m365-owa-tb.xpi"];

let debounce = null;
let building = false;

function rebuild() {
  if (building) return;
  building = true;
  const t0 = Date.now();
  try {
    // remove old xpi
    try { fs.rmSync(XPI, { force: true }); } catch {}
    // zip the source dir
    execSync(
      `zip -r -q m365-owa-tb.xpi . -x '.git/*' 'm365-owa-tb.xpi' 'node_modules/*' 'package.json' 'package-lock.json' '*.log' '.eslintcache' 'dev/watch.js'`,
      { cwd: ROOT, stdio: "pipe" }
    );
    const ms = Date.now() - t0;
    const size = (fs.statSync(XPI).size / 1024).toFixed(0);
    console.log(`[dev-watch] rebuilt m365-owa-tb.xpi (${size} KB) in ${ms}ms — reload addon now`);
  } catch (e) {
    console.error("[dev-watch] build failed:", e.message);
  } finally {
    building = false;
  }
}

function watch(dir) {
  for (const name of fs.readdirSync(dir)) {
    if (IGNORE.includes(name)) continue;
    const full = path.join(dir, name);
    let st;
    try { st = fs.statSync(full); } catch { continue; }
    if (st.isDirectory()) {
      watch(full);
      fs.watch(full, (event, filename) => {
        if (!filename) return;
        const f = path.join(full, filename);
        if (IGNORE.some((i) => f.includes(i))) return;
        if (!EXTS.has(path.extname(f))) return;
        schedule(f);
      });
    }
  }
}

function schedule(file) {
  if (debounce) clearTimeout(debounce);
  debounce = setTimeout(() => {
    debounce = null;
    console.log(`[dev-watch] change detected: ${path.relative(ROOT, file)}`);
    rebuild();
  }, 300);
}

console.log("[dev-watch] watching", ROOT);
watch(ROOT);
rebuild();
