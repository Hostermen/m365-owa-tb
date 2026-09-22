// OWA-backed calendar provider sync.
//
// m365-owa-cli is default-calendar-only (no calendar enumeration) and
// supports list / create / delete — NOT update (its UpdateEvent is explicitly
// `not_implemented`). So we expose exactly ONE provider calendar and:
//   * onSync        -> pull GetCalendarView into the cache
//   * onItemCreated -> push via CreateItem
//   * onItemRemoved -> push via DeleteItem
//   * onItemUpdated -> refused (MODIFY_FAILED) with a clear log line
var CalendarSync = {
  PROVIDER_TYPE: null,
  tbCalId: null,        // the single provider calendar id
  syncing: false,
  syncTimer: null,

  init() {
    this.PROVIDER_TYPE = "ext-" + messenger.runtime.id;
    this._registerProviderListeners();
    setTimeout(() => this._bootstrap().catch(e => console.error("[M365OWA] calendar bootstrap failed", e)), 2000);
  },

  _registerProviderListeners() {
    const cal = messenger.calendar;

    cal.provider.onInit.addListener(async (calendar) => {});

    cal.provider.onSync.addListener(async (calendar) => {
      await this._pull(calendar).catch(e => console.error("[M365OWA] onSync pull failed", e));
    });

    cal.provider.onResetSync.addListener(async (calendar) => {
      console.log("[M365OWA] reset sync for", calendar.id);
    });

    cal.provider.onItemCreated.addListener(async (calendar, item) => {
      if (this.syncing) return null;
      const owa = JCal.jcalToOwa(item.item);
      if (!owa || !owa.Subject || !owa.Start || !owa.End) return { error: "MODIFY_FAILED" };
      try {
        const created = await OWA.createEvent(owa);
        const jcalItem = JCal.eventToJcal(created);
        const etag = OWA.eventEtag(created) || (created.ItemId && created.ItemId.ChangeKey) || "";
        setTimeout(() => this._refresh(), 500);
        return { type: "event", format: "jcal", item: jcalItem, metadata: { etag } };
      } catch (e) {
        console.error("[M365OWA] push create failed", e.message || e);
        return { error: "MODIFY_FAILED" };
      }
    }, { returnFormat: "jcal" });

    cal.provider.onItemUpdated.addListener(async (_cal, item) => {
      if (this.syncing) return { error: "MODIFY_FAILED" };
      const ve = new _JCalComp(item.item);
      const ev = ve && ve.sub("vevent");
      const eventId = ev && ev.prop("uid");
      if (!eventId) return { error: "MODIFY_FAILED" };
      const etag = item.metadata && item.metadata.etag || "";
      const changeKey = etag.split("|")[0];
      const owa = JCal.jcalToOwa(item.item);
      if (!owa || !owa.Subject || !owa.Start || !owa.End) return { error: "MODIFY_FAILED" };
      try {
        const updated = await OWA.updateEvent(eventId, changeKey, owa);
        const newEtag = OWA.eventEtag(updated) + "|" + (updated.LastModifiedTime || "");
        setTimeout(() => this._refresh(), 500);
        return { type: "event", format: "jcal", item: JCal.eventToJcal(updated), metadata: { etag: newEtag } };
      } catch (e) {
        console.error("[M365OWA] push update failed", e.message || e);
        return { error: "MODIFY_FAILED" };
      }
    }, { returnFormat: "jcal" });

    cal.provider.onItemRemoved.addListener(async (calendar, item) => {
      if (this.syncing) return {};
      const ve = new _JCalComp(item.item);
      const ev = ve.sub("vevent");
      const eventId = ev && ev.prop("uid");
      if (!eventId) return {};
      try { await OWA.deleteEvent(eventId); }
      catch (e) { console.warn("[M365OWA] push delete failed (may be gone):", e.message || e); }
      setTimeout(() => this._refresh(), 500);
      return {};
    }, { returnFormat: "jcal" });
  },

  _refresh() {
    if (this.tbCalId) {
      messenger.calendar.calendars.synchronize().catch(() => {});
    }
  },

  async _bootstrap() {
    if (!Auth.isAuthenticated()) { console.log("[M365OWA] calendar bootstrap: not authenticated; waiting for token."); return; }
    if (this._booted) { console.log("[M365OWA] calendar already bootstrapped."); return; }
    this._booted = true;
    try {
      // Fetch calendar name from OWA; fall back to config default
      let calName = CONFIG.CAL_NAME;
      try {
        const owaName = await OWA.getCalendarFolderName();
        if (owaName) { calName = owaName; console.log("[M365OWA] fetched calendar name from OWA:", owaName); }
      } catch (e) { console.warn("[M365OWA] could not fetch calendar name from OWA:", e.message || e); }

      const existing = await messenger.calendar.calendars.query({ type: this.PROVIDER_TYPE });
      let cal = existing && existing[0];
      if (!cal) {
        cal = await messenger.calendar.calendars.create({
          type: this.PROVIDER_TYPE,
          url: "m365owa://default",
          name: calName,
          enabled: true,
          visible: true,
        });
        console.log("[M365OWA] created provider calendar ->", cal.id);
      } else {
        // update name + make sure it's visible+enabled after a relogin
        try {
          await messenger.calendar.calendars.update(cal.id, { name: calName, enabled: true, visible: true });
        } catch {}
      }
      this.tbCalId = cal.id;
      await messenger.calendar.calendars.synchronize();
      if (this.syncTimer) clearInterval(this.syncTimer);
      this.syncTimer = setInterval(() => messenger.calendar.calendars.synchronize(), CONFIG.SYNC_INTERVAL_MS);
      console.log("[M365OWA] calendar bootstrap done. provider cal:", this.tbCalId);
    } catch (e) {
      this._booted = false;
      console.error("[M365OWA] calendar bootstrap failed", e.message || e);
    }
  },

  async _pull(calendar) {
    if (this.syncing) return;
    this.syncing = true;
    try {
      const cacheId = calendar.cacheId;
      const cached = await messenger.calendar.items.query({ calendarId: cacheId, returnFormat: "jcal" });
      const cacheById = {};
      for (const it of (cached || [])) cacheById[it.id] = it.metadata?.etag || null;

      const now = new Date();
      const start = new Date(now.getTime() - CONFIG.PULL_DAYS_BACK * 86400000);
      const end = new Date(now.getTime() + CONFIG.PULL_DAYS_FORWARD * 86400000);
      console.log("[M365OWA] pulling", start.toISOString(), "->", end.toISOString());
      const items = await OWA.listEvents(start.toISOString(), end.toISOString());

      const seen = new Set();
      let created = 0, updated = 0, skipped = 0;
      for (const it of items) {
        const id = OWA.eventId(it);
        if (!id) { skipped++; continue; }
        seen.add(id);
        const jcal = JCal.eventToJcal(it);
        const etag = OWA.eventEtag(it) + "|" + (it.LastModifiedTime || it.lastModifiedTime || "");
        if (cacheById[id] === undefined) {
          await messenger.calendar.items.create(cacheId, { id, type: "event", format: "jcal", item: jcal, metadata: { etag } });
          created++;
        } else if (cacheById[id] !== etag) {
          await messenger.calendar.items.update(cacheId, id, { format: "jcal", item: jcal, metadata: { etag } });
          updated++;
        }
      }
      for (const id of Object.keys(cacheById)) {
        if (!seen.has(id)) { try { await messenger.calendar.items.remove(cacheId, id); } catch {} }
      }
      console.log("[M365OWA] pull done. total:", items.length, "created:", created, "updated:", updated, "skipped:", skipped);
      if (items.length && created === 0 && updated === 0 && skipped === items.length) {
        console.warn("[M365OWA] ALL items were skipped — likely the id field name differs. First item:", JSON.stringify(items[0]).slice(0, 500));
      }
    } catch (e) {
      console.error("[M365OWA] pull error", e.message || e);
    } finally {
      this.syncing = false;
    }
  },
};
