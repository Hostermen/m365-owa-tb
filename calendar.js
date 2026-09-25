// OWA-backed calendar provider sync.
var CalendarSync = {
  // provider type string (prefixed with the addon's runtime id)
  PROVIDER_TYPE: null,
  // the single Thunderbird provider calendar id
  tbCalId: null,
  // true while a pull is in progress (suppresses push to avoid echo)
  syncing: false,
  // setInterval handle for periodic syncs
  syncTimer: null,

  // Initialise the provider: set the type, register listeners, and bootstrap after a short delay.
  init() {
    // build the provider type from the addon's runtime id
    this.PROVIDER_TYPE = "ext-" + messenger.runtime.id;
    // register all calendar provider event listeners
    this._registerProviderListeners();
    // bootstrap after 2 seconds (allows config/token to load first)
    setTimeout(() => this._bootstrap().catch(e => console.error("[M365OWA] calendar bootstrap failed", e)), 2000);
  },

  // Register handlers for all calendar provider events (onInit, onSync, onItemCreated/Updated/Removed).
  _registerProviderListeners() {
    // shortcut to the calendar provider API
    const cal = messenger.calendar;

    // no-op onInit (required by the API contract)
    cal.provider.onInit.addListener(async (calendar) => {});

    // onSync: pull from OWA into the local cache
    cal.provider.onSync.addListener(async (calendar) => {
      // delegate to _pull, log on failure
      await this._pull(calendar).catch(e => console.error("[M365OWA] onSync pull failed", e));
    });

    // onResetSync: just log (no special handling needed)
    cal.provider.onResetSync.addListener(async (calendar) => {
      console.log("[M365OWA] reset sync for", calendar.id);
    });

    // onItemCreated: push a locally-created event to OWA via CreateItem
    cal.provider.onItemCreated.addListener(async (calendar, item) => {
      // suppress push during a pull to avoid echoing server changes
      if (this.syncing) return null;
      // convert the jCal item to OWA shape
      const owa = JCal.jcalToOwa(item.item);
      // validate required fields
      if (!owa || !owa.Subject || !owa.Start || !owa.End) return { error: "MODIFY_FAILED" };
      try {
        // push to OWA
        const created = await OWA.createEvent(owa);
        // convert the created item back to jCal
        const jcalItem = JCal.eventToJcal(created);
        // extract the etag (ChangeKey)
        const etag = OWA.eventEtag(created) || (created.ItemId && created.ItemId.ChangeKey) || "";
        // trigger a refresh so the cache picks up the server-side version
        setTimeout(() => this._refresh(), 500);
        return { type: "event", format: "jcal", item: jcalItem, metadata: { etag } };
      } catch (e) {
        // log and report failure
        console.error("[M365OWA] push create failed", e.message || e);
        return { error: "MODIFY_FAILED" };
      }
    }, { returnFormat: "jcal" });

    // onItemUpdated: push a locally-modified event to OWA via UpdateItem
    cal.provider.onItemUpdated.addListener(async (_cal, item) => {
      // suppress during pull
      if (this.syncing) return { error: "MODIFY_FAILED" };
      // parse the jCal to find the vevent
      const ve = new _JCalComp(item.item);
      const ev = ve && ve.sub("vevent");
      // extract the event UID (= OWA ItemId)
      const eventId = ev && ev.prop("uid");
      // no UID -> can't update
      if (!eventId) return { error: "MODIFY_FAILED" };
      // extract the etag (ChangeKey|LastModifiedTime)
      const etag = item.metadata && item.metadata.etag || "";
      // ChangeKey is the part before "|"
      const changeKey = etag.split("|")[0];
      // convert jCal to OWA shape
      const owa = JCal.jcalToOwa(item.item);
      // validate required fields
      if (!owa || !owa.Subject || !owa.Start || !owa.End) return { error: "MODIFY_FAILED" };
      try {
        // push the update to OWA
        const updated = await OWA.updateEvent(eventId, changeKey, owa);
        // build the new etag
        const newEtag = OWA.eventEtag(updated) + "|" + (updated.LastModifiedTime || "");
        // trigger a refresh
        setTimeout(() => this._refresh(), 500);
        return { type: "event", format: "jcal", item: JCal.eventToJcal(updated), metadata: { etag: newEtag } };
      } catch (e) {
        // log and report failure
        console.error("[M365OWA] push update failed", e.message || e);
        return { error: "MODIFY_FAILED" };
      }
    }, { returnFormat: "jcal" });

    // onItemRemoved: delete the event from OWA via DeleteItem
    cal.provider.onItemRemoved.addListener(async (calendar, item) => {
      // suppress during pull
      if (this.syncing) return {};
      // parse jCal to find the vevent
      const ve = new _JCalComp(item.item);
      const ev = ve.sub("vevent");
      // extract the event UID
      const eventId = ev && ev.prop("uid");
      // no UID -> nothing to delete
      if (!eventId) return {};
      try { await OWA.deleteEvent(eventId); }
      // log but don't fail (may already be gone)
      catch (e) { console.warn("[M365OWA] push delete failed (may be gone):", e.message || e); }
      // trigger a refresh
      setTimeout(() => this._refresh(), 500);
      return {};
    }, { returnFormat: "jcal" });
  },

  // Trigger a calendar sync if the provider calendar exists.
  _refresh() {
    if (this.tbCalId) {
      messenger.calendar.calendars.synchronize().catch(() => {});
    }
  },

  // Create the provider calendar (if missing), fetch its name from OWA, and start periodic syncs.
  async _bootstrap() {
    // wait for authentication
    if (!Auth.isAuthenticated()) { console.log("[M365OWA] calendar bootstrap: not authenticated; waiting for token."); return; }
    // avoid double bootstrap
    if (this._booted) { console.log("[M365OWA] calendar already bootstrapped."); return; }
    try {
      // fetch the calendar name from OWA (no fallback)
      const calName = await OWA.getCalendarFolderName();
      // abort if no name; retry later
      if (!calName) {
        console.warn("[M365OWA] could not fetch calendar name from OWA; will retry.");
        return;
      }
      console.log("[M365OWA] fetched calendar name from OWA:", calName);
      // mark as bootstrapped
      this._booted = true;

      // check if a provider calendar already exists
      const existing = await messenger.calendar.calendars.query({ type: this.PROVIDER_TYPE });
      let cal = existing && existing[0];
      if (!cal) {
        // create the provider calendar
        cal = await messenger.calendar.calendars.create({
          type: this.PROVIDER_TYPE,
          url: "m365owa://default",
          name: calName,
          enabled: true,
          visible: true,
        });
        console.log("[M365OWA] created provider calendar ->", cal.id);
      } else {
        // update name and ensure visible+enabled after relogin
        try {
          await messenger.calendar.calendars.update(cal.id, { name: calName, enabled: true, visible: true });
        } catch {}
      }
      // store the provider calendar id
      this.tbCalId = cal.id;
      // trigger an initial sync
      await messenger.calendar.calendars.synchronize();
      // start the periodic sync timer
      if (this.syncTimer) clearInterval(this.syncTimer);
      this.syncTimer = setInterval(() => messenger.calendar.calendars.synchronize(), CONFIG.SYNC_INTERVAL_MS);
      console.log("[M365OWA] calendar bootstrap done. provider cal:", this.tbCalId);
    } catch (e) {
      // reset bootstrapped flag on failure
      this._booted = false;
      console.error("[M365OWA] calendar bootstrap failed", e.message || e);
    }
  },

  // Pull events from OWA (GetCalendarView) into the local cache, creating/updating/removing items as needed.
  async _pull(calendar) {
    // prevent re-entrant pulls
    if (this.syncing) return;
    // mark as syncing (suppresses push)
    this.syncing = true;
    try {
      // get the cache id for this calendar
      const cacheId = calendar.cacheId;
      // read all currently cached items
      const cached = await messenger.calendar.items.query({ calendarId: cacheId, returnFormat: "jcal" });
      // build a lookup map: id -> etag
      const cacheById = {};
      for (const it of (cached || [])) cacheById[it.id] = it.metadata?.etag || null;

      // compute the pull window
      const now = new Date();
      const start = new Date(now.getTime() - CONFIG.PULL_DAYS_BACK * 86400000);
      const end = new Date(now.getTime() + CONFIG.PULL_DAYS_FORWARD * 86400000);
      console.log("[M365OWA] pulling", start.toISOString(), "->", end.toISOString());
      // fetch events from OWA
      const items = await OWA.listEvents(start.toISOString(), end.toISOString());

      // track which server ids we've seen
      const seen = new Set();
      // counters
      let created = 0, updated = 0, skipped = 0;
      // iterate each server event
      for (const it of items) {
        // extract the event id
        const id = OWA.eventId(it);
        // skip if no id
        if (!id) { skipped++; continue; }
        seen.add(id);
        // convert to jCal
        const jcal = JCal.eventToJcal(it);
        // build the etag (ChangeKey|LastModifiedTime)
        const etag = OWA.eventEtag(it) + "|" + (it.LastModifiedTime || it.lastModifiedTime || "");
        // not in cache -> create
        if (cacheById[id] === undefined) {
          await messenger.calendar.items.create(cacheId, { id, type: "event", format: "jcal", item: jcal, metadata: { etag } });
          created++;
        } else if (cacheById[id] !== etag) {
          // etag changed -> update
          await messenger.calendar.items.update(cacheId, id, { format: "jcal", item: jcal, metadata: { etag } });
          updated++;
        }
      }
      // remove cached items that no longer exist on the server
      for (const id of Object.keys(cacheById)) {
        if (!seen.has(id)) { try { await messenger.calendar.items.remove(cacheId, id); } catch {} }
      }
      console.log("[M365OWA] pull done. total:", items.length, "created:", created, "updated:", updated, "skipped:", skipped);
      // warn if all items were skipped (likely an id field name mismatch)
      if (items.length && created === 0 && updated === 0 && skipped === items.length) {
        console.warn("[M365OWA] ALL items were skipped — likely the id field name differs. First item:", JSON.stringify(items[0]).slice(0, 500));
      }
    } catch (e) {
      console.error("[M365OWA] pull error", e.message || e);
    } finally {
      // always clear the syncing flag
      this.syncing = false;
    }
  },
};
