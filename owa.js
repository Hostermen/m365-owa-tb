// OWA service.svc client — JS port of m365-owa-cli's `owa/client.py`.
//
// All operations POST to:  <OWA_HOST>/owa/service.svc?action=<Action>&app=<App>
// with `Authorization: Bearer <token>`, `Action: <Action>`, and an
// `__type`-tagged Exchange JSON-RPC-ish body. Responses are JSON with a `Body`
// object whose `ResponseCode` must be "NoError".
//
// Calendar (app=Calendar): GetCalendarFolders, GetCalendarView, CreateItem, DeleteItem
// People   (app=People)  : FindPeople   (best-effort — m365-owa-cli has no contacts)
var OWA = {
  _first(obj, keys) { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; },

  // OWA event id is nested: ItemId is {"__type":"ItemId:#Exchange","Id":"...","ChangeKey":"..."}.
  // Extract the string Id (and ChangeKey for etag) from any shape we've seen.
  eventId(it) {
    if (!it) return null;
    const fid = it.ItemId;
    if (fid && typeof fid === "object" && fid.Id) return String(fid.Id);
    if (typeof fid === "string") return fid;
    const v = this._first(it, ["Id", "id", "UID", "Uid", "itemId"]);
    return v != null ? String(v) : null;
  },
  eventEtag(it) {
    if (!it) return "";
    const fid = it.ItemId;
    if (fid && typeof fid === "object" && fid.ChangeKey) return String(fid.ChangeKey);
    return String(this._first(it, ["ChangeKey", "changeKey"]) || "");
  },


  _origin() {
    try { return new URL(CONFIG.OWA_HOST).origin; } catch { return String(CONFIG.OWA_HOST || "").replace(/\/+$/, ""); }
  },
  _url(action, app) {
    const q = new URLSearchParams({ action });
    if (app) q.set("app", app);
    return `${this._origin()}/owa/service.svc?${q}`;
  },

  async _headers(action) {
    const token = await Auth.getTokenAsync();
    return {
      Accept: "application/json",
      Action: action,
      "Content-Type": "application/json; charset=utf-8",
      "X-Requested-With": "XMLHttpRequest",
      Authorization: "Bearer " + token,
    };
  },

  async _postJson(action, app, payload) {
    const resp = await fetch(this._url(action, app), {
      method: "POST",
      headers: await this._headers(action),
      body: JSON.stringify(payload || {}),
      credentials: "omit",
    });
    if (resp.status === 401 || resp.status === 403) {
      await Auth.markExpired();
      throw new Error("OWA auth rejected (" + resp.status + "). Re-capture the bearer token.");
    }
    let data = null;
    try { data = await resp.json(); }
    catch { throw new Error("OWA returned non-JSON (HTTP " + resp.status + ")"); }

    const body = data && data.Body;
    const code = body && body.ResponseCode;
    if (resp.status >= 400 || (typeof code === "string" && code.toLowerCase() !== "noerror")) {
      const detail = body && body.Message
        ? (code + ": " + body.Message + (body.MessageXml ? " | " + JSON.stringify(body.MessageXml) : ""))
        : (JSON.stringify(data).slice(0, 600));
      throw new Error("OWA error " + resp.status + " / " + (code || "unknown") + ": " + detail);
    }
    return data;
  },

  // --- date formatting (ported from client._format_owa_range_boundary) ---
  _fmtBoundary(value, end) {
    let base;
    if (value instanceof Date) base = value.toISOString().split(".")[0];
    else {
      const t = String(value).replace(/Z$/, "");
      base = t.includes("T") ? t.split("+")[0] : `${t}T00:00:00`;
    }
    base = base.split(".")[0];
    return base + (end ? ".000" : ".001");
  },
  _fmtCreate(value) {
    if (value instanceof Date) return value.toISOString().split(".")[0] + ".000";
    const d = String(value).replace(/Z$/, "").split("+")[0];
    return d.split(".")[0] + ".000";
  },
  _isAllDay(start, end) {
    const s = new Date(start), e = new Date(end);
    if (isNaN(s) || isNaN(e)) return false;
    return s.toTimeString().slice(0, 8) === "00:00:00" && e.toTimeString().slice(0, 8) === "00:00:00" && (e - s) >= 86400000 && (e - s) % 86400000 === 0;
  },

  // Default calendar via DistinguishedFolderId — no folder enumeration needed.
  // "calendar" is the standard EWS distinguished id for the user's primary
  // calendar, so we skip GetCalendarFolders entirely (its response shape varies
  // by tenant and "Monarch" OWA doesn't always return CalendarGroups).
  _calendarId() {
    return {
      __type: "TargetFolderId:#Exchange",
      BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "calendar" },
    };
  },

  // Best-effort folder enumeration kept as an optional probe. Tolerant of any
  // response shape — never throws hard; returns null if it can't parse.
  async getCalendarFolders() {
    try {
      const data = await this._postJson("GetCalendarFolders", "Calendar", {});
      const body = data.Body || data;
      const groups = body.CalendarGroups || body.Folders || body.Items;
      if (!Array.isArray(groups)) return null;
      for (const g of groups) {
        const cals = g && (g.Calendars || [g]);
        if (!Array.isArray(cals)) continue;
        for (const c of cals) {
          const fid = c && (c.CalendarFolderId || c.FolderId || c);
          if (fid && fid.Id) return { Id: fid.Id, ChangeKey: fid.ChangeKey };
        }
      }
      return null;
    } catch {
      return null;
    }
  },

  // Fetch the display name of the primary calendar from OWA.
  async getCalendarFolderName() {
    try {
      const data = await this._postJson("GetCalendarFolders", "Calendar", {});
      const body = data.Body || data;
      const groups = body.CalendarGroups || body.Folders || body.Items;
      if (!Array.isArray(groups)) return null;
      for (const g of groups) {
        const cals = g && (g.Calendars || [g]);
        if (!Array.isArray(cals)) continue;
        for (const c of cals) {
          const name = this._first(c, ["DisplayName", "FolderName", "Name", "FolderDisplayName"]);
          if (name) return String(name);
        }
      }
      return null;
    } catch {
      return null;
    }
  },

  // Fetch the display name of the default contacts folder via EWS FindFolder.
  async getContactsFolderName() {
    try {
      const payload = {
        __type: "FindFolderJsonRequest:#Exchange",
        Header: {
          __type: "JsonRequestHeaders:#Exchange",
          RequestServerVersion: "Exchange2013",
          TimeZoneContext: {
            __type: "TimeZoneContext:#Exchange",
            TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
          },
        },
        Body: {
          __type: "FindFolderRequest:#Exchange",
          FolderShape: { __type: "FolderResponseShape:#Exchange", BaseShape: "Default" },
          ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "contacts" }],
          Traversal: "Shallow",
          Paging: { __type: "FolderView:#Exchange", MaxEntriesReturned: 100, BasePoint: "Beginning" },
        },
      };
      const data = await this._postJson("FindFolder", "People", payload);
      const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
      const m = Array.isArray(msgs) && msgs[0];
      if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") return null;
      const folders = m.RootFolder && m.RootFolder.Folders;
      if (Array.isArray(folders) && folders.length) {
        const name = this._first(folders[0], ["DisplayName", "FolderName", "Name"]);
        if (name) return String(name);
      }
      return null;
    } catch {
      return null;
    }
  },

  // --- Calendar: list events in a range (uses default calendar directly) ---
  async listEvents(startISO, endISO) {
    const payload = {
      __type: "GetCalendarViewJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "Exchange2013",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: {
        __type: "GetCalendarViewRequest:#Exchange",
        CalendarId: this._calendarId(),
        RangeStart: this._fmtBoundary(startISO, false),
        RangeEnd: this._fmtBoundary(endISO, true),
      },
    };
    const data = await this._postJson("GetCalendarView", "Calendar", payload);
    const body = data.Body || data;
    const items = body.Items || body.CalendarItems || body.Events || body.calendarItems || body.events;
    if (Array.isArray(items)) {
      console.log("[M365OWA] GetCalendarView returned", items.length, "items. Body keys:", Object.keys(body), "first item keys:", items[0] ? Object.keys(items[0]).join(",") : "(empty)");
      return items;
    }
    console.log("[M365OWA] GetCalendarView: no items array found. Response keys:", Object.keys(data), "Body keys:", body && Object.keys(body), "Body type:", typeof body);
    return [];
  },

  // --- Calendar: create an event (ported from create_event / _create_item_payload) ---
  // `item` is the OWA item shape built by jcal.js (Subject, Start, End, Body, Categories).
  async createEvent(item) {
    const calItem = Object.assign({
      __type: "CalendarItem:#Exchange",
      Subject: String(item.Subject || ""),
      Start: this._fmtCreate(item.Start),
      End: this._fmtCreate(item.End),
      IsAllDayEvent: !!item.IsAllDayEvent,
      ReminderIsSet: false,
    }, item.Body ? { Body: item.Body } : {}, item.Categories ? { Categories: item.Categories } : {});
    const payload = {
      __type: "CreateItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "Exchange2013",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: {
        __type: "CreateItemRequest:#Exchange",
        Items: [calItem],
        SendMeetingInvitations: "SendToNone",
      },
    };
    console.log("[M365OWA] createEvent payload:", JSON.stringify(calItem).slice(0, 500));
    const data = await this._postJson("CreateItem", "Calendar", payload);
    console.log("[M365OWA] createEvent response:", JSON.stringify(data).slice(0, 800));
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      const detail = m ? (m.ResponseCode + ": " + (m.MessageText || "") + " | " + JSON.stringify(m).slice(0, 400)) : "no response";
      throw new Error("OWA create failed: " + detail);
    }
    const created = Array.isArray(m.Items) && m.Items[0];
    console.log("[M365OWA] createEvent created item:", created ? JSON.stringify(created).slice(0, 500) : "NONE");
    return created || { Subject: item.Subject, Start: item.Start, End: item.End };
  },

  // --- Calendar: update an event (OWA UpdateItem — not in m365-owa-cli) ---
  // `owaFields` is the OWA item shape built by jcal.js jcalToOwa().
  // Each changed field becomes a SetItemField entry with its FieldURI.
  async updateEvent(itemId, changeKey, owaFields) {
    const Updates = [];
    const field = (uri, item) => ({ __type: "SetItemField:#Exchange", FieldURI: { __type: "FieldURI:#Exchange", FieldURI: uri }, Item: Object.assign({ __type: "CalendarItem:#Exchange" }, item) });
    if (owaFields.Subject != null) Updates.push(field("item:Subject", { Subject: String(owaFields.Subject) }));
    if (owaFields.Start != null) Updates.push(field("calendar:Start", { Start: this._fmtCreate(owaFields.Start) }));
    if (owaFields.End != null) Updates.push(field("calendar:End", { End: this._fmtCreate(owaFields.End) }));
    if (owaFields.IsAllDayEvent != null) Updates.push(field("calendar:IsAllDayEvent", { IsAllDayEvent: !!owaFields.IsAllDayEvent }));
    if (owaFields.Location != null) Updates.push(field("calendar:Location", { Location: String(owaFields.Location) }));
    if (owaFields.Body != null) Updates.push(field("item:Body", { Body: owaFields.Body }));
    if (owaFields.Categories != null) Updates.push(field("item:Categories", { Categories: owaFields.Categories }));
    if (!Updates.length) throw new Error("OWA update: no fields to update");

    const payload = {
      __type: "UpdateItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "Exchange2013",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: {
        __type: "UpdateItemRequest:#Exchange",
        ItemChanges: [{
          __type: "ItemChange:#Exchange",
          ItemId: { __type: "ItemId:#Exchange", Id: String(itemId), ChangeKey: String(changeKey || "") },
          Updates,
        }],
        SendMeetingInvitations: "SendToNone",
        ConflictResolution: "AlwaysOverwrite",
      },
    };
    const data = await this._postJson("UpdateItem", "Calendar", payload);
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA update failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    const updated = Array.isArray(m.Items) && m.Items[0];
    return updated || owaFields;
  },

  // --- Calendar: delete an event (ported from delete_event / _delete_item_payload) ---
  async deleteEvent(eventId) {
    const payload = {
      __type: "DeleteItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "DeleteItemRequest:#Exchange",
        ItemIds: [{ __type: "ItemId:#Exchange", Id: String(eventId) }],
        DeleteType: "MoveToDeletedItems",
        SendMeetingCancellations: "SendToNone",
        AffectedTaskOccurrences: "SpecifiedOccurrenceOnly",
      },
    };
    const data = await this._postJson("DeleteItem", "Calendar", payload);
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (m && String(m.ResponseClass || "").toLowerCase() !== "success" && String(m.ResponseCode || "").toLowerCase() !== "noerror") {
      throw new Error("OWA delete failed: " + (m.ResponseCode || "unknown"));
    }
  },

  // --- People: contacts CRUD (m365-owa-cli has no contacts) ---
  //
  // FindPeople is read-only (search endpoint). For create/update/delete we
  // use the generic EWS CreateItem/UpdateItem/DeleteItem on the contacts
  // folder (app=People). PersonaId from FindPeople is used as the ItemId
  // for update/delete — for personal contacts this is the EWS ItemId.

  // Create a contact. `contact` is an OWA Contact item built by VCard.vcardToOwa().
  async createContact(contact) {
    const item = Object.assign({ __type: "Contact:#Exchange" }, contact);
    const payload = {
      __type: "CreateItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "CreateItemRequest:#Exchange",
        Items: [item],
        ParentFolderId: {
          __type: "TargetFolderId:#Exchange",
          BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "contacts" },
        },
      },
    };
    console.log("[M365OWA] createContact payload:", JSON.stringify(item).slice(0, 500));
    const data = await this._postJson("CreateItem", "People", payload);
    console.log("[M365OWA] createContact response:", JSON.stringify(data).slice(0, 800));
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact create failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    const created = Array.isArray(m.Items) && m.Items[0];
    console.log("[M365OWA] createContact created item:", created ? JSON.stringify(created).slice(0, 500) : "NONE (m.Items empty)");
    return created || contact;
  },

  // Fetch a contact item by ItemId to get its current ChangeKey.
  // FindPeople returns PersonaId but not ChangeKey; UpdateItem requires it.
  async getContact(itemId) {
    const payload = {
      __type: "GetItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "GetItemRequest:#Exchange",
        ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" },
        ItemIds: [{ __type: "ItemId:#Exchange", Id: String(itemId) }],
      },
    };
    const data = await this._postJson("GetItem", "People", payload);
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA getContact failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    const item = Array.isArray(m.Items) && m.Items[0];
    return item || null;
  },

  // Update a contact. `itemId` is the EWS ItemId string.
  // `changeKey` is REQUIRED by EWS — obtained from findContacts().
  async updateContact(itemId, changeKey, contact) {
    const Updates = [];
    // Simple scalar fields use plain FieldURI
    const f = (uri, val) => ({
      __type: "SetItemField:#Exchange",
      FieldURI: { __type: "FieldURI:#Exchange", FieldURI: uri },
      Item: Object.assign({ __type: "Contact:#Exchange" }, val),
    });
    // Dictionary fields (email, phone) use IndexedFieldURI per entry
    const fi = (uri, idx, val) => ({
      __type: "SetItemField:#Exchange",
      FieldURI: { __type: "IndexedFieldURI:#Exchange", FieldURI: uri, FieldIndex: idx },
      Item: Object.assign({ __type: "Contact:#Exchange" }, val),
    });

    if (contact.GivenName != null) Updates.push(f("contacts:GivenName", { GivenName: String(contact.GivenName) }));
    if (contact.Surname != null) Updates.push(f("contacts:Surname", { Surname: String(contact.Surname) }));
    if (contact.CompanyName != null) Updates.push(f("contacts:CompanyName", { CompanyName: String(contact.CompanyName) }));
    if (contact.Department != null) Updates.push(f("contacts:Department", { Department: String(contact.Department) }));
    if (contact.JobTitle != null) Updates.push(f("contacts:JobTitle", { JobTitle: String(contact.JobTitle) }));
    if (contact.FileAs != null) Updates.push(f("contacts:FileAs", { FileAs: String(contact.FileAs) }));

    // EmailAddresses — each entry needs its own IndexedFieldURI
    if (Array.isArray(contact.EmailAddresses)) {
      for (const e of contact.EmailAddresses) {
        const key = e.Key || "EmailAddress1";
        const dictEntry = { __type: "EmailAddressDictionaryEntry:#Exchange", Key: key, Address: String(e.Address || ""), Name: String(e.Name || e.Address || "") };
        Updates.push(fi("contacts:EmailAddress", key, { EmailAddresses: [dictEntry] }));
      }
    }

    // PhoneNumbers — each entry needs its own IndexedFieldURI
    if (Array.isArray(contact.PhoneNumbers)) {
      for (const p of contact.PhoneNumbers) {
        const key = p.Key || "OtherTelephone";
        const dictEntry = { __type: "PhoneNumberDictionaryEntry:#Exchange", Key: key, Value: String(p.Value || "") };
        Updates.push(fi("contacts:PhoneNumber", key, { PhoneNumbers: [dictEntry] }));
      }
    }

    if (contact.Body != null) Updates.push(f("item:Body", { Body: contact.Body }));
    if (!Updates.length) throw new Error("OWA contact update: no fields to update");
    const itemIdObj = { __type: "ItemId:#Exchange", Id: String(itemId) };
    if (changeKey) itemIdObj.ChangeKey = String(changeKey);
    const payload = {
      __type: "UpdateItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "UpdateItemRequest:#Exchange",
        ItemChanges: [{
          __type: "ItemChange:#Exchange",
          ItemId: itemIdObj,
          Updates,
        }],
        ConflictResolution: "AlwaysOverwrite",
      },
    };
    console.log("[M365OWA] updateContact payload:", JSON.stringify(payload, null, 2));
    const data = await this._postJson("UpdateItem", "People", payload);
    console.log("[M365OWA] updateContact response:", JSON.stringify(data, null, 2).slice(0, 2000));
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact update failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    const updated = Array.isArray(m.Items) && m.Items[0];
    return updated || contact;
  },

  // Delete a contact by ItemId/PersonaId.
  async deleteContact(itemId, changeKey) {
    const itemObj = { __type: "ItemId:#Exchange", Id: String(itemId) };
    if (changeKey) itemObj.ChangeKey = String(changeKey);
    const payload = {
      __type: "DeleteItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "DeleteItemRequest:#Exchange",
        ItemIds: [itemObj],
        DeleteType: "HardDelete",
      },
    };
    console.log("[M365OWA] deleteContact payload:", JSON.stringify(payload));
    const data = await this._postJson("DeleteItem", "People", payload);
    console.log("[M365OWA] deleteContact response:", JSON.stringify(data).slice(0, 1000));
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    if (m && String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact delete failed: " + (m.ResponseCode || "unknown") + ": " + (m.MessageText || ""));
    }
  },

  // --- FindPeople: read contacts (app=People) ---
  // Uses OWA's FindPeople action against the default contacts
  // folder via DistinguishedFolderId "contacts" — same approach as the
  // calendar fix. If your tenant rejects this shape, contacts sync fails
  // gracefully; calendar is unaffected. See error console for the OWA
  // error detail so the payload can be adjusted.
  async findPeople(searchTerm, maxEntries = 200) {
    const payload = {
      __type: "FindPeopleJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: {
        __type: "FindPeopleRequest:#Exchange",
        IndexedPageItemView: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: maxEntries },
        QueryString: searchTerm || null,
        ParentFolderId: {
          __type: "TargetFolderId:#Exchange",
          BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "contacts" },
        },
        PersonaShape: { __type: "PersonaResponseShape:#Exchange", BaseShape: "Default" },
        ShouldResolveAmbiguousContacts: false,
      },
    };
    const data = await this._postJson("FindPeople", "People", payload);
    const body = data.Body || data;
    const ppl = body.ResultSet || body.People || body.Personas || body.Items || body.personas;
    if (Array.isArray(ppl)) {
      console.log("[M365OWA] FindPeople returned", ppl.length, "personas (TotalInView:", body.TotalNumberOfPeopleInView, "). first keys:", ppl[0] ? Object.keys(ppl[0]).join(",").slice(0,200) : "(empty)");
      return ppl;
    }
    console.log("[M365OWA] FindPeople: no personas array. Body keys:", body && Object.keys(body));
    return [];
  },

  // --- FindItem: list contacts (returns real EWS Contact items with ItemId+ChangeKey) ---
  // Unlike FindPeople (which returns aggregated Personas with PersonaId),
  // FindItem returns actual contact store items with proper EWS ItemId and
  // ChangeKey — both required for UpdateItem/DeleteItem.
  async findContacts(maxEntries = 200) {
    const payload = {
      __type: "FindItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "Exchange2013",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: {
        __type: "FindItemRequest:#Exchange",
        ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "Default" },
        ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "contacts" }],
        Traversal: "Shallow",
        Paging: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: maxEntries },
        ViewFilter: "All",
        ClutterFilter: "All",
        IsWarmUpSearch: 0,
        ShapeName: "MailListItem",
        SortOrder: [{ __type: "SortResults:#Exchange", Order: "Descending", Path: { __type: "PropertyUri:#Exchange", FieldURI: "DateTimeReceived" } }],
      },
    };
    const data = await this._postJson("FindItem", "People", payload);
    // Response: Body.ResponseMessages.Items[0].RootFolder.Items
    let items = [];
    try {
      const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
      const m = Array.isArray(msgs) && msgs[0];
      if (m && m.RootFolder && m.RootFolder.Items) items = m.RootFolder.Items;
    } catch {}
    if (!Array.isArray(items)) items = [];
    console.log("[M365OWA] FindItem returned", items.length, "contacts. first keys:", items[0] ? Object.keys(items[0]).join(",").slice(0,300) : "(empty)");
    return items;
  },

  // Extract ChangeKey from a FindItem contact's ItemId
  contactChangeKey(item) {
    if (!item || !item.ItemId) return "";
    return String(item.ItemId.ChangeKey || "");
  },

  // --- auth probe (mirrors `m365-owa-cli auth test`) ---
  // Tests auth + endpoint + default-calendar access in one call. Uses a tiny
  // 1-minute GetCalendarView window so we don't pull real data.
  async probe() {
    const now = new Date();
    const start = new Date(now.getTime() - 30000);
    const end = new Date(now.getTime() + 30000);
    await this.listEvents(start.toISOString(), end.toISOString());
    return { ok: true };
  },
};
