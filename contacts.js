// OWA-backed contacts sync (BIDIRECTIONAL).
//
// m365-owa-cli has no contacts surface; we reuse its OWA-service strategy to
// query people via FindPeople and write back via CreateItem/UpdateItem/DeleteItem.
//
// Sync strategy (runs every SYNC_INTERVAL_MS):
//   1. PULL: FindPeople -> server personas. Create/update local contacts when
//      the local vCard matches our stored "last known" copy (no local edit).
//   2. PUSH: Read all local contacts. Detect local-only changes by comparing
//      current vCard to stored "last known" vCard:
//        - new local contact (not in map)  -> CreateItem on server
//        - local vCard changed             -> UpdateItem on server
//        - local contact deleted            -> DeleteItem on server
//   3. Reconcile: server-deleted contacts are removed locally.
//
// Conflict policy: if both sides changed, server wins (local overwritten).
var ContactsSync = {
  abId: null,
  mapByOwa: {},       // { owaId: { localId, vcard } }  vcard = last known state
  busy: false,
  syncTimer: null,

  async init() {
    if (this._booted) { await this.sync(); return; }
    this._booted = true;
    try {
      await this._ensureAB();
      await this._loadMap();
      await this.sync();
      if (this.syncTimer) clearInterval(this.syncTimer);
      this.syncTimer = setInterval(() => this.sync(), CONFIG.SYNC_INTERVAL_MS);
      // Listen for local contact changes to trigger immediate push sync
      this._registerContactEvents();
      console.log("[M365OWA] contacts init done. AB:", this.abId, "mapped:", Object.keys(this.mapByOwa).length);
    } catch (e) {
      this._booted = false;
      console.error("[M365OWA] contacts init failed", e.message || e);
    }
  },

  _registerContactEvents() {
    if (this._eventsRegistered) return;
    this._eventsRegistered = true;
    let debounce = null;
    const trigger = (reason) => {
      console.log("[M365OWA] contact event:", reason, "-> scheduling push sync");
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => {
        debounce = null;
        this.sync().catch(e => console.warn("[M365OWA] event-triggered sync failed:", e.message || e));
      }, 2000);
    };
    try {
      messenger.contacts.onCreated.addListener((node) => {
        if (this._isOurAB(node.parentId)) trigger("created:" + node.id);
      });
      messenger.contacts.onUpdated.addListener((node) => {
        if (this._isOurAB(node.parentId)) trigger("updated:" + node.id);
      });
      messenger.contacts.onDeleted.addListener((parentId, id) => {
        if (this._isOurAB(parentId)) trigger("deleted:" + id);
      });
      console.log("[M365OWA] contact event listeners registered");
    } catch (e) {
      console.warn("[M365OWA] could not register contact events:", e.message || e);
    }
  },

  _isOurAB(parentId) {
    return String(parentId) === String(this.abId);
  },

  async _ensureAB() {
    // Fetch contacts folder name from OWA; fall back to config default
    let abName = CONFIG.AB_NAME;
    try {
      const owaName = await OWA.getContactsFolderName();
      if (owaName) { abName = owaName; console.log("[M365OWA] fetched contacts folder name from OWA:", owaName); }
    } catch (e) { console.warn("[M365OWA] could not fetch contacts folder name from OWA:", e.message || e); }

    const all = await messenger.addressBooks.list(true);
    let ab = all.find(a => a.name === abName);
    if (!ab) ab = await messenger.addressBooks.create({ name: abName });
    this.abId = ab.id;
  },

  async _loadMap() {
    const { m365_owa_contact_map } = await browser.storage.local.get("m365_owa_contact_map");
    const raw = m365_owa_contact_map || {};
    // Migrate old format: { owaId: "localIdString" } -> { owaId: { localId, vcard: "" } }
    this.mapByOwa = {};
    for (const [k, v] of Object.entries(raw)) {
      if (typeof v === "string") this.mapByOwa[k] = { localId: v, vcard: "" };
      else this.mapByOwa[k] = v;
    }
  },
  async _saveMap() { await browser.storage.local.set({ m365_owa_contact_map: this.mapByOwa }); },

  // Reverse map: localId -> owaId
  _reverseMap() {
    const rev = {};
    for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
      if (entry && entry.localId) rev[entry.localId] = owaId;
    }
    return rev;
  },

  // Read a single contact's vCard, handling both vCard-format and
  // property-bag-format Thunderbird address books.
  async _readContactVCard(id) {
    try {
      const full = await messenger.contacts.get(id);
      if (!full) return "";
      if (full.vCard) return VCard.vcardNormalized(full.vCard);
      if (full.properties) {
        const v = VCard.fromProperties(full.properties);
        if (v) return VCard.vcardNormalized(v);
      }
      return "";
    } catch (e) {
      console.warn("[M365OWA] _readContactVCard: contacts.get() threw for", id, "-", e.message || e);
      return "";
    }
  },

  async _readLocalContacts() {
    const nodes = await messenger.contacts.list(this.abId);
    const result = {};
    for (const node of nodes) {
      let vcard = node.vCard;
      if (!vcard && node.properties) {
        vcard = VCard.fromProperties(node.properties);
      }
      if (!vcard) {
        vcard = await this._readContactVCard(node.id);
      }
      if (!vcard) {
        console.warn("[M365OWA] _readLocalContacts: NO vCard for", node.id,
          "| node keys:", Object.keys(node).join(","),
          "| properties keys:", node.properties
            ? (typeof node.properties.keys === "function"
                ? Array.from(node.properties.keys()).join(",")
                : Object.keys(node.properties).join(","))
            : "(none)");
      }
      result[node.id] = vcard ? VCard.vcardNormalized(vcard) : "";
    }
    return result;
  },

  async sync() {
    if (this.busy) return;
    if (!Auth.isAuthenticated()) return;
    this.busy = true;
    let stats = { sCreated: 0, sUpdated: 0, sDeleted: 0, lCreated: 0, lUpdated: 0, lDeleted: 0 };
    try {
      // --- 1. PULL: server -> local ---
      const contacts = await OWA.findContacts();
      const seen = new Set();
      const localContacts = await this._readLocalContacts();
      const reverseMap = this._reverseMap();

      for (const c of contacts) {
        const serverVcard = VCard.fromContact(c);
        if (!serverVcard) continue;
        const owaId = VCard.idFromVcard(serverVcard);
        if (!owaId) continue;
        const changeKey = OWA.contactChangeKey(c);
        seen.add(owaId);
        const entry = this.mapByOwa[owaId];
        const storedVcard = entry ? VCard.vcardNormalized(entry.vcard) : null;
        const localId = entry && entry.localId;
        const localVcard = localId ? localContacts[localId] : null;

        if (!entry || !localId) {
          // New on server -> create locally
          const node = await messenger.contacts.create(this.abId, { vCard: serverVcard });
          const newId = typeof node === "string" ? node : (node && node.id);
          // Read back the actual stored vCard (TB may store as property-bag)
          const actualVcard = await this._readContactVCard(newId) || VCard.vcardNormalized(serverVcard);
          this.mapByOwa[owaId] = { localId: newId, vcard: actualVcard, changeKey };
          stats.sCreated++;
        } else if (localVcard == null || localVcard === "") {
          // Local contact was deleted by user OR vCard not readable
          if (localVcard == null) {
            // Truly deleted locally -> local deletion wins: push delete to server (handled in PUSH phase)
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
          const localChanged = VCard.vcardCanonical(localVcard) !== VCard.vcardCanonical(storedVcard);
          const serverChanged = VCard.vcardCanonical(serverVcard) !== VCard.vcardCanonical(storedVcard);
          if (!localChanged) {
            // No local edit since last sync
            if (serverChanged) {
              // Server changed -> update local
              try { await messenger.contacts.update(localId, { vCard: serverVcard }); } catch {}
              const actualVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(serverVcard);
              this.mapByOwa[owaId].vcard = actualVcard;
              this.mapByOwa[owaId].changeKey = changeKey;
              stats.sUpdated++;
            }
          } else {
            // Local was edited
            if (serverChanged) {
              // Both changed -> conflict: server wins
              try { await messenger.contacts.update(localId, { vCard: serverVcard }); } catch {}
              const actualVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(serverVcard);
              this.mapByOwa[owaId].vcard = actualVcard;
              this.mapByOwa[owaId].changeKey = changeKey;
              stats.sUpdated++;
              console.warn("[M365OWA] contact conflict for", owaId, "- server wins");
            }
            // else: local changed only -> push in PUSH phase
          }
          // Always update ChangeKey from FindItem (server may have changed it)
          if (changeKey) this.mapByOwa[owaId].changeKey = changeKey;
        }
      }

      // --- 2. Remove locally: contacts deleted on server ---
      for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
        if (!seen.has(owaId) && entry && entry.localId) {
          // Only delete locally if the user didn't edit it
          const localVcard = localContacts[entry.localId];
          if (localVcard && VCard.vcardCanonical(localVcard) === VCard.vcardCanonical(entry.vcard)) {
            try { await messenger.contacts.delete(entry.localId); } catch {}
            stats.sDeleted++;
          }
          delete this.mapByOwa[owaId];
        }
      }

      // --- 3. PUSH: local -> server ---
      const freshLocal = await this._readLocalContacts();
      const freshReverse = this._reverseMap();
      console.log("[M365OWA] PUSH phase. freshLocal ids:", Object.keys(freshLocal), "map:", JSON.stringify(this.mapByOwa));

      // Locally created contacts (not in map)
      for (const [localId, localVcard] of Object.entries(freshLocal)) {
        if (freshReverse[localId]) {
          console.log("[M365OWA] push skip (already mapped):", localId);
          continue;
        }
        if (!localVcard) {
          console.log("[M365OWA] push skip (no readable vCard):", localId);
          continue;
        }
        const owaId = VCard.idFromVcard(localVcard);
        if (owaId && this.mapByOwa[owaId]) {
          console.log("[M365OWA] push skip (vCard OWA id already mapped to another localId):", localId, "owaId:", owaId);
          continue;
        }
        if (owaId) {
          // vCard carries an OWA id from a previous pull. Two sub-cases:
          if (seen.has(owaId)) {
            // Contact still exists on server -> map was lost, just re-associate.
            console.log("[M365OWA] push re-associate (still on server):", localId, "owaId:", owaId);
            this.mapByOwa[owaId] = { localId, vcard: localVcard };
            continue;
          }
          // OWA id is NOT in current FindPeople results -> contact was deleted on
          // server (or never existed). Push it as a NEW contact and rewrite the
          // stale X-M365-OWA-ID line below with the freshly assigned id.
          console.log("[M365OWA] push new (stale OWA id, not on server):", localId, "old owaId:", owaId);
          // fall through to the push block
        } else {
          console.log("[M365OWA] push new (no OWA id, genuinely local):", localId);
        }

        // Genuinely new local contact -> push to server
        try {
          const owaContact = VCard.vcardToOwa(localVcard);
          const created = await OWA.createContact(owaContact);
          const newOwaId = OWA.eventId(created)
            || (created && created.ItemId && created.ItemId.Id)
            || (created && created.Id);
          if (!newOwaId) {
            console.warn("[M365OWA] push create: server gave no ItemId for", localId,
              "-> contact may exist on server untracked. Response:", JSON.stringify(created).slice(0, 300));
            continue;
          }
          // If the local vCard had a stale OWA id, rewrite it to the new one so
          // we don't loop forever (re-push) on subsequent syncs.
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
          // Read back actual vCard from TB (may be property-bag format)
          const storedVcard = await this._readContactVCard(localId) || VCard.vcardNormalized(updatedVcard);
          this.mapByOwa[String(newOwaId)] = { localId, vcard: storedVcard };
          stats.lCreated++;
          console.log("[M365OWA] push created on server. localId:", localId, "newOwaId:", newOwaId);
        } catch (e) {
          console.warn("[M365OWA] push create contact failed:", e.message || e);
        }
      }

      // Locally modified / deleted contacts
      for (const [owaId, entry] of Object.entries(this.mapByOwa)) {
        if (!entry || !entry.localId) continue;
        const localVcard = freshLocal[entry.localId];
        if (localVcard == null) {
          // Locally deleted -> delete on server
          console.log("[M365OWA] pushing contact DELETE to OWA:", owaId, "entry.localId:", entry.localId, "not in freshLocal");
          try {
            await OWA.deleteContact(owaId, entry.changeKey);
            stats.lDeleted++;
            console.log("[M365OWA] push deleted on server. owaId:", owaId);
          } catch (e) {
            console.warn("[M365OWA] push delete contact failed:", e.message || e);
          }
          delete this.mapByOwa[owaId];
        } else if (localVcard && VCard.vcardCanonical(localVcard) !== VCard.vcardCanonical(entry.vcard)) {
          // Locally modified -> push update to server
          // OWA's People module uses PeopleGraphVx REST API (not service.svc UpdateItem),
          // so we implement update as delete + create.
          console.log("[M365OWA] pushing contact update (delete+create) to OWA:", owaId, "localVcard changed");
          try {
            const owaContact = VCard.vcardToOwa(localVcard);
            // 1. Delete old contact
            try {
              await OWA.deleteContact(owaId, entry.changeKey);
              console.log("[M365OWA] update: deleted old contact", owaId);
            } catch (e) {
              console.warn("[M365OWA] update: delete old contact failed:", e.message || e, "- trying create anyway");
            }
            // Remove old mapping
            delete this.mapByOwa[owaId];
            // 2. Create new contact with updated data
            const created = await OWA.createContact(owaContact);
            const newId = created && created.ItemId && (created.ItemId.Id || created.ItemId.id);
            const newCk = created && created.ItemId && created.ItemId.ChangeKey;
            if (newId) {
              const newOwaId = String(newId);
              // Read back actual vCard from TB (may be property-bag format)
              const storedVcard = await this._readContactVCard(entry.localId) || VCard.vcardNormalized(localVcard);
              this.mapByOwa[newOwaId] = { localId: entry.localId, vcard: storedVcard, changeKey: newCk ? String(newCk) : "" };
              // Rewrite X-M365-OWA-ID in local vCard to new ID
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

      await this._saveMap();
      console.log("[M365OWA] contacts sync done.",
        "pull: created:", stats.sCreated, "updated:", stats.sUpdated, "deleted:", stats.sDeleted,
        "| push: created:", stats.lCreated, "updated:", stats.lUpdated, "deleted:", stats.lDeleted,
        "| mapped:", Object.keys(this.mapByOwa).length);
    } catch (e) {
      console.error("[M365OWA] contacts sync failed:", e.message || e);
    } finally {
      this.busy = false;
    }
  },
};
