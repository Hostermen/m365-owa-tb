// OWA service.svc client.
var OWA = {
  // Return the first non-null value found under any of the given keys.
  _first(obj, keys) { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; },

  // Extract the string ItemId from an OWA event item (handles object and string shapes).
  eventId(it) {
    if (!it) return null;
    // ItemId is the primary field
    const fid = it.ItemId;
    // object form: { Id: "..." }
    if (fid && typeof fid === "object" && fid.Id) return String(fid.Id);
    // string form
    if (typeof fid === "string") return fid;
    // fallback: try alternate keys
    const v = this._first(it, ["Id", "id", "UID", "Uid", "itemId"]);
    return v != null ? String(v) : null;
  },

  // Extract the ChangeKey (etag) from an OWA event item's ItemId.
  eventEtag(it) {
    if (!it) return "";
    // ItemId object with ChangeKey
    const fid = it.ItemId;
    if (fid && typeof fid === "object" && fid.ChangeKey) return String(fid.ChangeKey);
    // fallback: try alternate keys
    return String(this._first(it, ["ChangeKey", "changeKey"]) || "");
  },

  // Return the OWA origin (scheme + host + port) from CONFIG.OWA_HOST.
  _origin() {
    // parse as URL, fall back to stripping trailing slashes
    try { return new URL(CONFIG.OWA_HOST).origin; } catch { return String(CONFIG.OWA_HOST || "").replace(/\/+$/, ""); }
  },

  // Build the service.svc URL for a given action and app.
  _url(action, app) {
    // build query string with the action
    const q = new URLSearchParams({ action });
    // add app if provided
    if (app) q.set("app", app);
    // assemble the full URL
    return `${this._origin()}/owa/service.svc?${q}`;
  },

  // Build the HTTP headers for a service.svc request (includes the bearer token).
  async _headers(action) {
    // get the current token (may auto-refresh)
    const token = await Auth.getTokenAsync();
    return {
      Accept: "application/json",
      // Action header must match the URL action parameter
      Action: action,
      "Content-Type": "application/json; charset=utf-8",
      "X-Requested-With": "XMLHttpRequest",
      // bearer token
      Authorization: "Bearer " + token,
    };
  },

  // POST a JSON payload to service.svc and return the parsed response; throws on auth failure or OWA error.
  async _postJson(action, app, payload) {
    // send the POST request
    const resp = await fetch(this._url(action, app), {
      method: "POST",
      headers: await this._headers(action),
      body: JSON.stringify(payload || {}),
      credentials: "omit",
    });
    // 401/403 -> mark token expired and throw
    if (resp.status === 401 || resp.status === 403) {
      await Auth.markExpired();
      throw new Error("OWA auth rejected (" + resp.status + "). Re-capture the bearer token.");
    }
    // parse the JSON response
    let data = null;
    try { data = await resp.json(); }
    // non-JSON response -> error
    catch { throw new Error("OWA returned non-JSON (HTTP " + resp.status + ")"); }

    // extract the response body and code
    const body = data && data.Body;
    const code = body && body.ResponseCode;
    // check for HTTP errors or OWA error codes
    if (resp.status >= 400 || (typeof code === "string" && code.toLowerCase() !== "noerror")) {
      // build a detail string
      const detail = body && body.Message
        ? (code + ": " + body.Message + (body.MessageXml ? " | " + JSON.stringify(body.MessageXml) : ""))
        : (JSON.stringify(data).slice(0, 600));
      throw new Error("OWA error " + resp.status + " / " + (code || "unknown") + ": " + detail);
    }
    return data;
  },

  // Format a date boundary string for OWA GetCalendarView (with millisecond precision).
  _fmtBoundary(value, end) {
    let base;
    if (value instanceof Date) base = value.toISOString().split(".")[0];
    else {
      // strip Z suffix for naive strings
      const t = String(value).replace(/Z$/, "");
      // if it has a time component, use it; else add midnight
      base = t.includes("T") ? t.split("+")[0] : `${t}T00:00:00`;
    }
    // strip any existing milliseconds
    base = base.split(".")[0];
    // end boundary uses .000, start uses .001
    return base + (end ? ".000" : ".001");
  },

  // Format a date for OWA CreateItem/UpdateItem (UTC, with .000 milliseconds).
  _fmtCreate(value) {
    // Date object -> ISO with .000
    if (value instanceof Date) return value.toISOString().split(".")[0] + ".000";
    // string: strip Z and offset, add .000
    const d = String(value).replace(/Z$/, "").split("+")[0];
    return d.split(".")[0] + ".000";
  },

  // Detect if a start/end pair represents an all-day event (midnight to midnight, >= 1 day).
  _isAllDay(start, end) {
    const s = new Date(start), e = new Date(end);
    // invalid dates -> not all-day
    if (isNaN(s) || isNaN(e)) return false;
    // both at midnight and duration >= 1 day and a multiple of 1 day
    return s.toTimeString().slice(0, 8) === "00:00:00" && e.toTimeString().slice(0, 8) === "00:00:00" && (e - s) >= 86400000 && (e - s) % 86400000 === 0;
  },

  // Return the EWS TargetFolderId for the user's primary calendar (distinguished folder id "calendar").
  _calendarId() {
    return {
      __type: "TargetFolderId:#Exchange",
      BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "calendar" },
    };
  },

  // Best-effort enumeration of calendar folders; returns { Id, ChangeKey } or null on any parse failure.
  async getCalendarFolders() {
    try {
      // POST GetCalendarFolders
      const data = await this._postJson("GetCalendarFolders", "Calendar", {});
      const body = data.Body || data;
      // find the groups array
      const groups = body.CalendarGroups || body.Folders || body.Items;
      if (!Array.isArray(groups)) return null;
      // iterate groups looking for a calendar folder id
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

  // Fetch the display name of the primary calendar via GetFolder on the "calendar" distinguished folder id.
  async getCalendarFolderName() {
    // try GetFolder first (most reliable)
    try {
      const payload = {
        __type: "GetFolderJsonRequest:#Exchange",
        Header: {
          __type: "JsonRequestHeaders:#Exchange",
          RequestServerVersion: "Exchange2013",
          TimeZoneContext: {
            __type: "TimeZoneContext:#Exchange",
            TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
          },
        },
        Body: {
          __type: "GetFolderRequest:#Exchange",
          FolderShape: { __type: "FolderResponseShape:#Exchange", BaseShape: "Default" },
          FolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "calendar" }],
        },
      };
      const data = await this._postJson("GetFolder", "Calendar", payload);
      // navigate to the first response message
      const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
      const m = Array.isArray(msgs) && msgs[0];
      // check for success
      if (m && String(m.ResponseClass || "").toLowerCase() === "success") {
        const folders = m.Folders;
        if (Array.isArray(folders) && folders.length) {
          // extract the display name
          const name = this._first(folders[0], ["DisplayName", "FolderName", "Name"]);
          if (name) return String(name);
        }
      }
    } catch (e) { console.warn("[M365OWA] GetFolder(calendar) error:", e.message || e); }

    // fallback: parse GetCalendarFolders
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
    } catch (e) {
      console.warn("[M365OWA] getCalendarFolderName error:", e.message || e);
      return null;
    }
  },

  // Fetch the display name of the default contacts folder via EWS GetFolder.
  async getContactsFolderName() {
    try {
      const payload = {
        __type: "GetFolderJsonRequest:#Exchange",
        Header: {
          __type: "JsonRequestHeaders:#Exchange",
          RequestServerVersion: "Exchange2013",
          TimeZoneContext: {
            __type: "TimeZoneContext:#Exchange",
            TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
          },
        },
        Body: {
          __type: "GetFolderRequest:#Exchange",
          FolderShape: { __type: "FolderResponseShape:#Exchange", BaseShape: "Default" },
          FolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "contacts" }],
        },
      };
      const data = await this._postJson("GetFolder", "People", payload);
      // navigate to the first response message
      const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
      const m = Array.isArray(msgs) && msgs[0];
      // check for success
      if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") return null;
      const folders = m.Folders;
      if (Array.isArray(folders) && folders.length) {
        // extract the display name
        const name = this._first(folders[0], ["DisplayName", "FolderName", "Name"]);
        if (name) return String(name);
      }
      return null;
    } catch (e) {
      console.warn("[M365OWA] getContactsFolderName error:", e.message || e);
      return null;
    }
  },

  // List calendar events in a date range via GetCalendarView on the default calendar.
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
        // use the default calendar (distinguished folder id)
        CalendarId: this._calendarId(),
        // format the boundary dates for OWA
        RangeStart: this._fmtBoundary(startISO, false),
        RangeEnd: this._fmtBoundary(endISO, true),
      },
    };
    const data = await this._postJson("GetCalendarView", "Calendar", payload);
    const body = data.Body || data;
    // find the items array (field name varies by OWA version)
    const items = body.Items || body.CalendarItems || body.Events || body.calendarItems || body.events;
    if (Array.isArray(items)) {
      console.log("[M365OWA] GetCalendarView returned", items.length, "items. Body keys:", Object.keys(body), "first item keys:", items[0] ? Object.keys(items[0]).join(",") : "(empty)");
      return items;
    }
    console.log("[M365OWA] GetCalendarView: no items array found. Response keys:", Object.keys(data), "Body keys:", body && Object.keys(body), "Body type:", typeof body);
    return [];
  },

  // Create a calendar event via CreateItem; returns the created item from the OWA response.
  async createEvent(item) {
    // build the calendar item, merging in Body and Categories if present
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
        // don't send meeting invitations
        SendMeetingInvitations: "SendToNone",
      },
    };
    console.log("[M365OWA] createEvent payload:", JSON.stringify(calItem).slice(0, 500));
    const data = await this._postJson("CreateItem", "Calendar", payload);
    console.log("[M365OWA] createEvent response:", JSON.stringify(data).slice(0, 800));
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      const detail = m ? (m.ResponseCode + ": " + (m.MessageText || "") + " | " + JSON.stringify(m).slice(0, 400)) : "no response";
      throw new Error("OWA create failed: " + detail);
    }
    // extract the created item
    const created = Array.isArray(m.Items) && m.Items[0];
    console.log("[M365OWA] createEvent created item:", created ? JSON.stringify(created).slice(0, 500) : "NONE");
    // return the created item, or a fallback with the input fields
    return created || { Subject: item.Subject, Start: item.Start, End: item.End };
  },

  // Update a calendar event via UpdateItem with SetItemField entries for each changed field.
  async updateEvent(itemId, changeKey, owaFields) {
    // build the Updates array
    const Updates = [];
    // helper to create a SetItemField entry
    const field = (uri, item) => ({ __type: "SetItemField:#Exchange", FieldURI: { __type: "FieldURI:#Exchange", FieldURI: uri }, Item: Object.assign({ __type: "CalendarItem:#Exchange" }, item) });
    // add an update for each changed field
    if (owaFields.Subject != null) Updates.push(field("item:Subject", { Subject: String(owaFields.Subject) }));
    if (owaFields.Start != null) Updates.push(field("calendar:Start", { Start: this._fmtCreate(owaFields.Start) }));
    if (owaFields.End != null) Updates.push(field("calendar:End", { End: this._fmtCreate(owaFields.End) }));
    if (owaFields.IsAllDayEvent != null) Updates.push(field("calendar:IsAllDayEvent", { IsAllDayEvent: !!owaFields.IsAllDayEvent }));
    if (owaFields.Location != null) Updates.push(field("calendar:Location", { Location: String(owaFields.Location) }));
    if (owaFields.Body != null) Updates.push(field("item:Body", { Body: owaFields.Body }));
    if (owaFields.Categories != null) Updates.push(field("item:Categories", { Categories: owaFields.Categories }));
    // error if nothing to update
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
          // identify the item by Id + ChangeKey
          ItemId: { __type: "ItemId:#Exchange", Id: String(itemId), ChangeKey: String(changeKey || "") },
          Updates,
        }],
        SendMeetingInvitations: "SendToNone",
        // always overwrite on conflict
        ConflictResolution: "AlwaysOverwrite",
      },
    };
    const data = await this._postJson("UpdateItem", "Calendar", payload);
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA update failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    // extract the updated item
    const updated = Array.isArray(m.Items) && m.Items[0];
    return updated || owaFields;
  },

  // Delete a calendar event via DeleteItem (moves to deleted items).
  async deleteEvent(eventId) {
    const payload = {
      __type: "DeleteItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "DeleteItemRequest:#Exchange",
        ItemIds: [{ __type: "ItemId:#Exchange", Id: String(eventId) }],
        // move to Deleted Items folder
        DeleteType: "MoveToDeletedItems",
        SendMeetingCancellations: "SendToNone",
        AffectedTaskOccurrences: "SpecifiedOccurrenceOnly",
      },
    };
    const data = await this._postJson("DeleteItem", "Calendar", payload);
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success (NoError is also acceptable)
    if (m && String(m.ResponseClass || "").toLowerCase() !== "success" && String(m.ResponseCode || "").toLowerCase() !== "noerror") {
      throw new Error("OWA delete failed: " + (m.ResponseCode || "unknown"));
    }
  },

  // Create a contact via CreateItem on the contacts folder (app=People).
  async createContact(contact) {
    // merge the Contact type tag
    const item = Object.assign({ __type: "Contact:#Exchange" }, contact);
    const payload = {
      __type: "CreateItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "CreateItemRequest:#Exchange",
        Items: [item],
        // target the default contacts folder
        ParentFolderId: {
          __type: "TargetFolderId:#Exchange",
          BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "contacts" },
        },
      },
    };
    console.log("[M365OWA] createContact payload:", JSON.stringify(item).slice(0, 500));
    const data = await this._postJson("CreateItem", "People", payload);
    console.log("[M365OWA] createContact response:", JSON.stringify(data).slice(0, 800));
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact create failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    // extract the created item
    const created = Array.isArray(m.Items) && m.Items[0];
    console.log("[M365OWA] createContact created item:", created ? JSON.stringify(created).slice(0, 500) : "NONE (m.Items empty)");
    return created || contact;
  },

  // Fetch a contact by ItemId to obtain its current ChangeKey (needed for UpdateItem).
  async getContact(itemId) {
    const payload = {
      __type: "GetItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "GetItemRequest:#Exchange",
        // only need the id
        ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" },
        ItemIds: [{ __type: "ItemId:#Exchange", Id: String(itemId) }],
      },
    };
    const data = await this._postJson("GetItem", "People", payload);
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA getContact failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    // extract the item
    const item = Array.isArray(m.Items) && m.Items[0];
    return item || null;
  },

  // Update a contact via UpdateItem with SetItemField entries (scalar fields use FieldURI, dict fields use IndexedFieldURI).
  async updateContact(itemId, changeKey, contact) {
    // build the Updates array
    const Updates = [];
    // helper for scalar fields (plain FieldURI)
    const f = (uri, val) => ({
      __type: "SetItemField:#Exchange",
      FieldURI: { __type: "FieldURI:#Exchange", FieldURI: uri },
      Item: Object.assign({ __type: "Contact:#Exchange" }, val),
    });
    // helper for dictionary fields (IndexedFieldURI per entry)
    const fi = (uri, idx, val) => ({
      __type: "SetItemField:#Exchange",
      FieldURI: { __type: "IndexedFieldURI:#Exchange", FieldURI: uri, FieldIndex: idx },
      Item: Object.assign({ __type: "Contact:#Exchange" }, val),
    });

    // scalar fields
    if (contact.GivenName != null) Updates.push(f("contacts:GivenName", { GivenName: String(contact.GivenName) }));
    if (contact.Surname != null) Updates.push(f("contacts:Surname", { Surname: String(contact.Surname) }));
    if (contact.CompanyName != null) Updates.push(f("contacts:CompanyName", { CompanyName: String(contact.CompanyName) }));
    if (contact.Department != null) Updates.push(f("contacts:Department", { Department: String(contact.Department) }));
    if (contact.JobTitle != null) Updates.push(f("contacts:JobTitle", { JobTitle: String(contact.JobTitle) }));
    if (contact.FileAs != null) Updates.push(f("contacts:FileAs", { FileAs: String(contact.FileAs) }));

    // email addresses — each needs its own IndexedFieldURI
    if (Array.isArray(contact.EmailAddresses)) {
      for (const e of contact.EmailAddresses) {
        const key = e.Key || "EmailAddress1";
        const dictEntry = { __type: "EmailAddressDictionaryEntry:#Exchange", Key: key, Address: String(e.Address || ""), Name: String(e.Name || e.Address || "") };
        Updates.push(fi("contacts:EmailAddress", key, { EmailAddresses: [dictEntry] }));
      }
    }

    // phone numbers — each needs its own IndexedFieldURI
    if (Array.isArray(contact.PhoneNumbers)) {
      for (const p of contact.PhoneNumbers) {
        const key = p.Key || "OtherTelephone";
        const dictEntry = { __type: "PhoneNumberDictionaryEntry:#Exchange", Key: key, Value: String(p.Value || "") };
        Updates.push(fi("contacts:PhoneNumber", key, { PhoneNumbers: [dictEntry] }));
      }
    }

    // body/notes
    if (contact.Body != null) Updates.push(f("item:Body", { Body: contact.Body }));
    // error if nothing to update
    if (!Updates.length) throw new Error("OWA contact update: no fields to update");
    // build the ItemId object (include ChangeKey if provided)
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
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (!m || String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact update failed: " + ((m && m.ResponseCode) || "no response") + ": " + ((m && m.MessageText) || ""));
    }
    // extract the updated item
    const updated = Array.isArray(m.Items) && m.Items[0];
    return updated || contact;
  },

  // Delete a contact by ItemId via DeleteItem (hard delete).
  async deleteContact(itemId, changeKey) {
    // build the ItemId object (include ChangeKey if provided)
    const itemObj = { __type: "ItemId:#Exchange", Id: String(itemId) };
    if (changeKey) itemObj.ChangeKey = String(changeKey);
    const payload = {
      __type: "DeleteItemJsonRequest:#Exchange",
      Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
      Body: {
        __type: "DeleteItemRequest:#Exchange",
        ItemIds: [itemObj],
        // permanently delete
        DeleteType: "HardDelete",
      },
    };
    console.log("[M365OWA] deleteContact payload:", JSON.stringify(payload));
    const data = await this._postJson("DeleteItem", "People", payload);
    console.log("[M365OWA] deleteContact response:", JSON.stringify(data).slice(0, 1000));
    // navigate to the first response message
    const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
    const m = Array.isArray(msgs) && msgs[0];
    // check for success
    if (m && String(m.ResponseClass || "").toLowerCase() !== "success") {
      throw new Error("OWA contact delete failed: " + (m.ResponseCode || "unknown") + ": " + (m.MessageText || ""));
    }
  },

  // Search for people via FindPeople against the default contacts folder.
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
        // pagination settings
        IndexedPageItemView: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: maxEntries },
        // search query (null = all)
        QueryString: searchTerm || null,
        // target the default contacts folder
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
    // find the personas array (field name varies)
    const ppl = body.ResultSet || body.People || body.Personas || body.Items || body.personas;
    if (Array.isArray(ppl)) {
      console.log("[M365OWA] FindPeople returned", ppl.length, "personas (TotalInView:", body.TotalNumberOfPeopleInView, "). first keys:", ppl[0] ? Object.keys(ppl[0]).join(",").slice(0,200) : "(empty)");
      return ppl;
    }
    console.log("[M365OWA] FindPeople: no personas array. Body keys:", body && Object.keys(body));
    return [];
  },

  // List contacts via FindItem (returns real EWS Contact items with ItemId + ChangeKey).
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
        // request all default properties
        ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "Default" },
        // target the default contacts folder
        ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "contacts" }],
        // shallow traversal (no sub-folders)
        Traversal: "Shallow",
        // pagination
        Paging: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: maxEntries },
        ViewFilter: "All",
        ClutterFilter: "All",
        IsWarmUpSearch: 0,
        ShapeName: "MailListItem",
        // sort by date received, descending
        SortOrder: [{ __type: "SortResults:#Exchange", Order: "Descending", Path: { __type: "PropertyUri:#Exchange", FieldURI: "DateTimeReceived" } }],
      },
    };
    const data = await this._postJson("FindItem", "People", payload);
    // navigate to the items array: Body.ResponseMessages.Items[0].RootFolder.Items
    let items = [];
    try {
      const msgs = data.Body && data.Body.ResponseMessages && data.Body.ResponseMessages.Items;
      const m = Array.isArray(msgs) && msgs[0];
      if (m && m.RootFolder && m.RootFolder.Items) items = m.RootFolder.Items;
    } catch {}
    // ensure it's an array
    if (!Array.isArray(items)) items = [];
    console.log("[M365OWA] FindItem returned", items.length, "contacts. first keys:", items[0] ? Object.keys(items[0]).join(",").slice(0,300) : "(empty)");
    return items;
  },

  // Extract the ChangeKey from a FindItem contact's ItemId.
  contactChangeKey(item) {
    if (!item || !item.ItemId) return "";
    return String(item.ItemId.ChangeKey || "");
  },

  // Auth probe: test authentication, endpoint reachability, and calendar access with a tiny GetCalendarView window.
  async probe() {
    // use a 1-minute window around now
    const now = new Date();
    const start = new Date(now.getTime() - 30000);
    const end = new Date(now.getTime() + 30000);
    // if listEvents doesn't throw, the probe succeeded
    await this.listEvents(start.toISOString(), end.toISOString());
    return { ok: true };
  },
};
