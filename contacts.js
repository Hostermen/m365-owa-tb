// OWA-backed contacts sync (bidirectional).
var ContactsSync = {
  // Thunderbird address book id for our synced contacts
  abId: null,
  // map: { owaId: { localId, vcard, changeKey } } — vcard = last known state
  mapByOwa: {},
  // true while a sync is in progress (prevents re-entrant syncs)
  busy: false,
  // setInterval handle for periodic syncs
  syncTimer: null,

  // Initialise the contacts sync: ensure the address book, load the map, run a sync, and start the periodic timer.
  async init() {
    // if already bootstrapped, just run a sync
    if (this._booted) { await this.sync(); return; }
    this._booted = true;
    try {
      // ensure the address book exists
      await this._ensureAB();
      // load the contact map from storage
      await this._loadMap();
      // run an initial sync
      await this.sync();
      // start the periodic sync timer
      if (this.syncTimer) clearInterval(this.syncTimer);
      this.syncTimer = setInterval(() => this.sync(), CONFIG.SYNC_INTERVAL_MS);
      // register listeners for local contact changes
      this._registerContactEvents();
      console.log("[M365OWA] contacts init done. AB:", this.abId, "mapped:", Object.keys(this.mapByOwa).length);
    } catch (e) {
      // reset bootstrapped flag on failure
      this._booted = false;
      console.error("[M365OWA] contacts init failed", e.message || e);
    }
  },

  // Register Thunderbird contact event listeners that trigger a debounced push sync.
  _registerContactEvents() {
    // avoid double-registration
    if (this._eventsRegistered) return;
    this._eventsRegistered = true;
    // debounce timer handle
    let debounce = null;
    // schedule a sync after 2 seconds of quiet (debounces rapid changes)
    const trigger = (reason) => {
      console.log("[M365OWA] contact event:", reason, "-> scheduling push sync");
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => {
        debounce = null;
        this.sync().catch(e => console.warn("[M365OWA] event-triggered sync failed:", e.message || e));
      }, 2000);
    };
    try {
      // local contact created -> trigger
      messenger.contacts.onCreated.addListener((node) => {
        if (this._isOurAB(node.parentId)) trigger("created:" + node.id);
      });
      // local contact updated -> trigger
      messenger.contacts.onUpdated.addListener((node) => {
        if (this._isOurAB(node.parentId)) trigger("updated:" + node.id);
      });
      // local contact deleted -> trigger
      messenger.contacts.onDeleted.addListener((parentId, id) => {
        if (this._isOurAB(parentId)) trigger("deleted:" + id);
      });
      console.log("[M365OWA] contact event listeners registered");
    } catch (e) {
      console.warn("[M365OWA] could not register contact events:", e.message || e);
    }
  },

  // Return true if the given parent id matches our address book id.
  _isOurAB(parentId) {
    return String(parentId) === String(this.abId);
  },

  // Ensure our address book exists (fetch its name from OWA first, then find or create it).
  async _ensureAB() {
    // fetch the contacts folder name from OWA (no fallback)
    const abName = await OWA.getContactsFolderName();
    if (!abName) {
      console.warn("[M365OWA] could not fetch contacts folder name from OWA; will retry.");
      throw new Error("OWA contacts folder name not available");
    }
    console.log("[M365OWA] fetched contacts folder name from OWA:", abName);

    // list all address books (including sub-books)
    const all = await messenger.addressBooks.list(true);
    // find an existing AB with the matching name
    let ab = all.find(a => a.name === abName);
    if (ab) {
      // use the existing AB id
      this.abId = ab.id;
    } else {
      // create a new AB (returns id string directly)
      this.abId = await messenger.addressBooks.create({ name: abName });
    }
  },

  // Load the contact map from storage.local, migrating the old string-only format.
  async _loadMap() {
    // read the stored map
    const { m365_owa_contact_map } = await browser.storage.local.get("m365_owa_contact_map");
    const raw = m365_owa_contact_map || {};
    // migrate old format: { owaId: "localIdString" } -> { owaId: { localId, vcard: "" } }
    this.mapByOwa = {};
    for (const [k, v] of Object.entries(raw)) {
      if (typeof v === "string") this.mapByOwa[k] = { localId: v, vcard: "" };
      else this.mapByOwa[k] = v;
    }
  },

  // Persist the contact map to storage.local.
  async _saveMap() { await browser.storage.local.set({ m365_owa_contact_map: this.mapByOwa }); },

  // Build and return a reverse map: localId -> owaId.
  _reverseMap() {
    const rev = {};
    for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
      if (entry && entry.localId) rev[entry.localId] = owaId;
    }
    return rev;
  },

  // Read a single contact's vCard from Thunderbird, handling both vCard and property-bag formats.
  async _readContactVCard(id) {
    try {
      // fetch the full contact
      const full = await messenger.contacts.get(id);
      // not found -> empty
      if (!full) return "";
      // if TB provides a vCard field, use it
      if (full.vCard) return VCard.vcardNormalized(full.vCard);
      // if TB provides a property bag, convert it
      if (full.properties) {
        const v = VCard.fromProperties(full.properties);
        if (v) return VCard.vcardNormalized(v);
      }
      // no vCard at all
      return "";
    } catch (e) {
      console.warn("[M365OWA] _readContactVCard: contacts.get() threw for", id, "-", e.message || e);
      return "";
    }
  },

  // Read all contacts in our address book and return a map: localId -> normalised vCard.
  async _readLocalContacts() {
    // list all contacts in our AB
    const nodes = await messenger.contacts.list(this.abId);
    const result = {};
    for (const node of nodes) {
      // try the node's vCard field first
      let vcard = node.vCard;
      // if no vCard, try converting from properties
      if (!vcard && node.properties) {
        vcard = VCard.fromProperties(node.properties);
      }
      // if still no vCard, try fetching it individually
      if (!vcard) {
        vcard = await this._readContactVCard(node.id);
      }
      // warn if no vCard could be obtained
      if (!vcard) {
        console.warn("[M365OWA] _readLocalContacts: NO vCard for", node.id,
          "| node keys:", Object.keys(node).join(","),
          "| properties keys:", node.properties
            ? (typeof node.properties.keys === "function"
                ? Array.from(node.properties.keys()).join(",")
                : Object.keys(node.properties).join(","))
            : "(none)");
      }
      // store the normalised vCard (or empty string)
      result[node.id] = vcard ? VCard.vcardNormalized(vcard) : "";
    }
    return result;
  },

  // Run a full bidirectional sync: pull server contacts, reconcile deletions, then push local changes.
  async sync() {
    // skip if already syncing
    if (this.busy) return;
    // skip if not authenticated
    if (!Auth.isAuthenticated()) return;
    this.busy = true;
    // stats counters
    let stats = { sCreated: 0, sUpdated: 0, sDeleted: 0, lCreated: 0, lUpdated: 0, lDeleted: 0 };
    try {
      // --- 1. PULL: server -> local ---
      // fetch all server contacts via FindItem
      const contacts = await OWA.findContacts();
      // track which server owaIds we've seen
      const seen = new Set();
      // read all local contacts
      const localContacts = await this._readLocalContacts();
      // build reverse map for lookups
      const reverseMap = this._reverseMap();

      // iterate each server contact
      for (const c of contacts) {
        // convert server contact to vCard
        const serverVcard = VCard.fromContact(c);
        if (!serverVcard) continue;
        // extract the OWA id from the vCard
        const owaId = VCard.idFromVcard(serverVcard);
        if (!owaId) continue;
        // extract the ChangeKey
        const changeKey = OWA.contactChangeKey(c);
        seen.add(owaId);
        // look up the existing map entry
        const entry = this.mapByOwa[owaId];
        // stored vCard (last known state)
        const storedVcard = entry ? VCard.vcardNormalized(entry.vcard) : null;
        // local contact id
        const localId = entry && entry.localId;
        // local vCard (current state)
        const localVcard = localId ? localContacts[localId] : null;

        if (!entry || !localId) {
          // new on server -> create locally
          const node = await messenger.contacts.create(this.abId, { vCard: serverVcard });
          const newId = typeof node === "string" ? node : (node && node.id);
          // read back the actual stored vCard (TB may store as property-bag)
          const actualVcard = await this._readContactVCard(newId) || VCard.vcardNormalized(serverVcard);
          this.mapByOwa[owaId] = { localId: newId, vcard: actualVcard, changeKey };
          stats.sCreated++;
        } else if (localVcard == null || localVcard === "") {
          // local contact was deleted by user OR vCard not readable
          if (localVcard == null) {
            // truly deleted locally -> local deletion wins (handled in PUSH phase)
          } else {
            // vCard not readable -> trust server: update local from server
            if (VCard.vcardCanonical(serverVcard) !== VCard.vcardCanonical(storedVcard)) {
              try { await messenger.contacts.update(localId, { vCard: serverVcard }); } catch {}
              const actualVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(serverVcard);
              this.mapByOwa[owaId].vcard = actualVcard;
              this.mapByOwa[owaId].changeKey = changeKey;
              stats.sUpdated++;
            }
          }
        } else {
          // both local and server copies exist: compare to stored baseline
          const localChanged = VCard.vcardCanonical(localVcard) !== VCard.vcardCanonical(storedVcard);
          const serverChanged = VCard.vcardCanonical(serverVcard) !== VCard.vcardCanonical(storedVcard);
          if (!localChanged) {
            // no local edit since last sync
            if (serverChanged) {
              // server changed -> update local
              try { await messenger.contacts.update(localId, { vCard: serverVcard }); } catch {}
              const actualVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(serverVcard);
              this.mapByOwa[owaId].vcard = actualVcard;
              this.mapByOwa[owaId].changeKey = changeKey;
              stats.sUpdated++;
            }
          } else {
            // local was edited
            if (serverChanged) {
              // both changed -> conflict: server wins
              try { await messenger.contacts.update(localId, { vCard: serverVcard }); } catch {}
              const actualVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(serverVcard);
              this.mapByOwa[owaId].vcard = actualVcard;
              this.mapByOwa[owaId].changeKey = changeKey;
              stats.sUpdated++;
              console.warn("[M365OWA] contact conflict for", owaId, "- server wins");
            }
            // else: local changed only -> push in PUSH phase
          }
          // always update ChangeKey from FindItem
          if (changeKey) this.mapByOwa[owaId].changeKey = changeKey;
        }
      }

      // --- 2. Remove locally: contacts deleted on server ---
      for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
        if (!seen.has(owaId) && entry && entry.localId) {
          // only delete locally if the user didn't edit it
          const localVcard = localContacts[entry.localId];
          if (localVcard && VCard.vcardCanonical(localVcard) === VCard.vcardCanonical(entry.vcard)) {
            try { await messenger.contacts.delete(entry.localId); } catch {}
            stats.sDeleted++;
          }
          // remove from map regardless
          delete this.mapByOwa[owaId];
        }
      }

      // --- 3. PUSH: local -> server ---
      // re-read local contacts (may have changed during PULL)
      const freshLocal = await this._readLocalContacts();
      const freshReverse = this._reverseMap();
      console.log("[M365OWA] PUSH phase. freshLocal ids:", Object.keys(freshLocal), "map:", JSON.stringify(this.mapByOwa));

      // locally created contacts (not in map)
      for (const [localId, localVcard] of Object.entries(freshLocal)) {
        // skip if already mapped
        if (freshReverse[localId]) {
          console.log("[M365OWA] push skip (already mapped):", localId);
          continue;
        }
        // skip if no readable vCard
        if (!localVcard) {
          console.log("[M365OWA] push skip (no readable vCard):", localId);
          continue;
        }
        // extract OWA id from vCard (may have a stale one from a previous pull)
        const owaId = VCard.idFromVcard(localVcard);
        // skip if the OWA id is already mapped to another local contact
        if (owaId && this.mapByOwa[owaId]) {
          console.log("[M365OWA] push skip (vCard OWA id already mapped to another localId):", localId, "owaId:", owaId);
          continue;
        }
        if (owaId) {
          // vCard carries an OWA id from a previous pull — two sub-cases:
          if (seen.has(owaId)) {
            // contact still exists on server -> map was lost, just re-associate
            console.log("[M365OWA] push re-associate (still on server):", localId, "owaId:", owaId);
            this.mapByOwa[owaId] = { localId, vcard: localVcard };
            continue;
          }
          // OWA id not in current results -> deleted on server -> push as NEW
          console.log("[M365OWA] push new (stale OWA id, not on server):", localId, "old owaId:", owaId);
          // fall through to the push block
        } else {
          // no OWA id -> genuinely local contact
          console.log("[M365OWA] push new (no OWA id, genuinely local):", localId);
        }

        // push the contact to the server via CreateItem
        try {
          // convert vCard to OWA contact shape
          const owaContact = VCard.vcardToOwa(localVcard);
          // create on server
          const created = await OWA.createContact(owaContact);
          // extract the new ItemId
          const newOwaId = OWA.eventId(created)
            || (created && created.ItemId && created.ItemId.Id)
            || (created && created.Id);
          if (!newOwaId) {
            console.warn("[M365OWA] push create: server gave no ItemId for", localId,
              "-> contact may exist on server untracked. Response:", JSON.stringify(created).slice(0, 300));
            continue;
          }
          // if the local vCard had a stale OWA id, rewrite it to the new one
          let updatedVcard = localVcard;
          if (owaId && String(owaId) !== String(newOwaId)) {
            updatedVcard = localVcard.replace(
              /^X-M365-OWA-ID:.*$/m,
              "X-M365-OWA-ID:" + String(newOwaId)
            );
            try {
              await messenger.contacts.update(localId, { vCard: updatedVcard });
              updatedVcard = VCard.vcardNormalized(updatedVcard);
            } catch (e) {
              console.warn("[M365OWA] could not rewrite stale OWA id in local vCard:", e.message || e);
              updatedVcard = VCard.vcardNormalized(localVcard);
            }
          }
          // read back the actual vCard from TB (may be property-bag format)
          const storedVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(updatedVcard);
          this.mapByOwa[String(newOwaId)] = { localId, vcard: storedVcard };
          stats.lCreated++;
          console.log("[M365OWA] push created on server. localId:", localId, "newOwaId:", newOwaId);
        } catch (e) {
          console.warn("[M365OWA] push create contact failed:", e.message || e);
        }
      }

      // locally modified / deleted contacts
      for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
        if (!entry || !entry.localId) continue;
        const localVcard = freshLocal[entry.localId];
        if (localVcard == null) {
          // locally deleted -> delete on server
          console.log("[M365OWA] pushing contact DELETE to OWA:", owaId, "entry.localId:", entry.localId, "not in freshLocal");
          try {
            await OWA.deleteContact(owaId, entry.changeKey);
            stats.lDeleted++;
            console.log("[M365OWA] push deleted on server. owaId:", owaId);
          } catch (e) {
            console.warn("[M365OWA] push delete contact failed:", e.message || e);
          }
          // remove from map
          delete this.mapByOwa[owaId];
        } else if (localVcard && VCard.vcardCanonical(localVcard) !== VCard.vcardCanonical(entry.vcard)) {
          // locally modified -> push update (implemented as delete + create)
          console.log("[M365OWA] pushing contact update (delete+create) to OWA:", owaId, "localVcard changed");
          try {
            // convert vCard to OWA contact shape
            const owaContact = VCard.vcardToOwa(localVcard);
            // 1. delete old contact
            try {
              await OWA.deleteContact(owaId, entry.changeKey);
              console.log("[M365OWA] update: deleted old contact", owaId);
            } catch (e) {
              console.warn("[M365OWA] update: delete old contact failed:", e.message || e, "- trying create anyway");
            }
            // remove old mapping
            delete this.mapByOwa[owaId];
            // 2. create new contact with updated data
            const created = await OWA.createContact(owaContact);
            // extract new ItemId and ChangeKey
            const newId = created && created.ItemId && (created.ItemId.Id || created.ItemId.id);
            const newCk = created && created.ItemId && created.ItemId.ChangeKey;
            if (newId) {
              const newOwaId = String(newId);
              // read back actual vCard from TB
              const storedVcard = await this._readContactVCard(entry.localId) || VCard.vcardNormalized(localVcard);
              this.mapByOwa[newOwaId] = { localId: entry.localId, vcard: storedVcard, changeKey: newCk ? String(newCk) : "" };
              // rewrite X-M365-OWA-ID in local vCard to the new id
              try {
                const rewritten = VCard.rewriteOwaId(localVcard, newOwaId);
                if (rewritten) {
                  await messenger.contacts.update(entry.localId, { vCard: rewritten });
                  const actualVcard = await this._readContactVCard(entry.localId) || VCard.vcardNormalized(rewritten);
                  this.mapByOwa[newOwaId].vcard = actualVcard;
                }
              } catch (e) {
                console.warn("[M365OWA] update: could not rewrite OWA id in local vCard:", e.message || e);
              }
              stats.lUpdated++;
              console.log("[M365OWA] push updated on server. oldOwaId:", owaId, "newOwaId:", newOwaId);
            } else {
              console.warn("[M365OWA] push update: createContact returned no ItemId");
            }
          } catch (e) {
            console.warn("[M365OWA] push update contact failed:", e.message || e);
          }
        }
      }

      // persist the updated map
      await this._saveMap();
      console.log("[M365OWA] contacts sync done.",
        "pull: created:", stats.sCreated, "updated:", stats.sUpdated, "deleted:", stats.sDeleted,
        "| push: created:", stats.lCreated, "updated:", stats.lUpdated, "deleted:", stats.lDeleted,
        "| mapped:", Object.keys(this.mapByOwa).length);
    } catch (e) {
      console.error("[M365OWA] contacts sync failed:", e.message || e);
    } finally {
      // always clear the busy flag
      this.busy = false;
    }
  },
};
