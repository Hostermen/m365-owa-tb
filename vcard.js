// Minimal vCard 4.0 builder from an OWA "persona" (FindPeople response).
//
// Contacts are READ-ONLY here: m365-owa-cli has no contacts surface, and OWA's
// FindPeople is a people-search read endpoint. We only build vCard strings to
// populate a local Thunderbird address book; local edits are not pushed back.
var VCard = {
  esc(s) {
    if (s == null) return "";
    return String(s).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");
  },
  first(obj, keys) { for (const k of keys) if (obj && obj[k] != null && obj[k] !== "") return obj[k]; return null; },
  arr(v) { return Array.isArray(v) ? v : (v ? [v] : []); },

  // OWA persona -> vCard 4.0 string
  fromOwa(p) {
    if (!p) return null;
    const L = ["BEGIN:VCARD", "VERSION:4.0"];
    const fn = this.first(p, ["DisplayName", "displayName", "Alias", "FullName"]) ||
      [this.first(p, ["GivenName", "FirstName"]), this.first(p, ["Surname", "LastName"])].filter(Boolean).join(" ");
    if (fn) L.push("FN:" + this.esc(fn));
    const given = this.first(p, ["GivenName", "FirstName"]);
    const sur = this.first(p, ["Surname", "LastName"]);
    L.push("N:" + this.esc(sur || "") + ";" + this.esc(given || "") + ";;;");

    const company = this.first(p, ["CompanyName", "companyName", "Company"]);
    const dept = this.first(p, ["Department", "department"]);
    const title = this.first(p, ["Title", "JobTitle", "title", "jobTitle"]);
    if (company || dept) L.push("ORG:" + this.esc(company || "") + (dept ? ";" + this.esc(dept) : ""));
    if (title) L.push("TITLE:" + this.esc(title));

    // emails: OWA personas usually have EmailAddress (single) + EmailAddresses (array)
    const emails = new Set();
    for (const e of this.arr(this.first(p, ["EmailAddresses", "emailAddresses"]))) {
      const addr = typeof e === "string" ? e : (e && (e.EmailAddress || e.Address || e.address || e.email));
      if (addr) emails.add(String(addr));
    }
    const single = this.first(p, ["EmailAddress", "emailAddress", "PrimarySmtpAddress"]);
    if (single) emails.add(String(single));
    for (const e of emails) L.push("EMAIL:" + this.esc(e));

    // phones
    const mobile = this.first(p, ["MobilePhoneNumber", "MobilePhone", "mobilePhone"]);
    if (mobile) L.push("TEL;TYPE=cell:" + this.esc(mobile));
    for (const ph of this.arr(this.first(p, ["HomePhones", "HomePhoneNumber"]))) {
      const v = typeof ph === "string" ? ph : (ph && ph.Number);
      if (v) L.push("TEL;TYPE=home:" + this.esc(v));
    }
    for (const ph of this.arr(this.first(p, ["BusinessPhones", "BusinessPhoneNumber", "WorkPhone"]))) {
      const v = typeof ph === "string" ? ph : (ph && ph.Number);
      if (v) L.push("TEL;TYPE=work:" + this.esc(v));
    }

    // address (best-effort)
    const ba = this.first(p, ["BusinessAddress", "WorkAddress", "OfficeLocation"]);
    if (ba && typeof ba === "object") {
      const parts = [ba.Street || ba.StreetAddress || "", ba.City || "", ba.State || "", ba.PostalCode || "", ba.Country || ba.CountryOrRegion || ""];
      if (parts.some(Boolean)) L.push("ADR;TYPE=work:;;" + parts.map(this.esc).join(";"));
    }
    if (ba && typeof ba === "string" && ba) L.push("ADR;TYPE=work:;;" + this.esc(ba) + ";;;;");

    const note = this.first(p, ["Notes", "personalNotes", "PersonalNotes"]);
    if (note) L.push("NOTE:" + this.esc(note));

    // stable id so we can dedupe across syncs: prefer PersonaId / ExchangeGuid,
    // else synthesise one from primary email + display name.
    let pid = this.first(p, ["PersonaId", "Id", "ExchangeGuid", "GuidId"]);
    if (pid && typeof pid === "object") pid = pid.Id || pid.id || pid.GuidId || null;
    if (!pid) {
      const seed = (single && String(single)) || fn || "";
      let h = 0; for (let i = 0; i < seed.length; i++) { h = ((h << 5) - h + seed.charCodeAt(i)) | 0; }
      pid = "syn-" + (h >>> 0).toString(16);
    }
    L.push("X-M365-OWA-ID:" + this.esc(String(pid)));

    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Unfold vCard line folding (RFC 6350 §3.2): a line starting with a
  // space or tab continues the previous line.  Must be called before
  // any line-by-line processing.
  _unfold(vcard) {
    return String(vcard || "")
      .replace(/\r\n/g, "\n").replace(/\r/g, "\n")
      .replace(/\n[ \t]/g, "");
  },

  // extract our stable OWA id from a vCard (if we wrote one) for dedupe
  idFromVcard(vcard) {
    const m = this._unfold(vcard).match(/^X-M365-OWA-ID:(.+)$/m);
    return m ? m[1].trim() : null;
  },

  // Replace the X-M365-OWA-ID in a vCard with a new value (after delete+create)
  rewriteOwaId(vcard, newOwaId) {
    if (!vcard || !newOwaId) return null;
    const unfolded = this._unfold(vcard);
    if (!/^X-M365-OWA-ID:/m.test(unfolded)) return null;
    const replaced = unfolded.replace(/^X-M365-OWA-ID:.+$/m, "X-M365-OWA-ID:" + this.esc(newOwaId));
    return replaced.replace(/\n/g, "\r\n");
  },

  // Parse vCard lines into key→[values] map (case-insensitive, handles folding)
  _parse(vcard) {
    const lines = this._unfold(vcard).split("\n");
    const props = [];
    for (const line of lines) {
      if (!line || line === "BEGIN:VCARD" || line === "END:VCARD" || line.startsWith("VERSION:")) continue;
      const colon = line.indexOf(":");
      if (colon < 0) continue;
      const name = line.slice(0, colon);
      const value = line.slice(colon + 1);
      const semi = name.indexOf(";");
      const key = (semi >= 0 ? name.slice(0, semi) : name).toUpperCase();
      const params = semi >= 0 ? name.slice(semi + 1) : "";
      props.push({ key, params, value: this._unesc(value) });
    }
    return props;
  },
  _unesc(s) { return String(s || "").replace(/\\n/g, "\n").replace(/\\,/g, ",").replace(/\\;/g, ";").replace(/\\\\/g, "\\"); },

  // vCard string -> OWA Contact item for CreateItem/UpdateItem
  vcardToOwa(vcard) {
    const props = this._parse(vcard);
    const get = (key) => { const p = props.find(p => p.key === key); return p ? p.value : null; };
    const getAll = (key) => props.filter(p => p.key === key).map(p => p.value);
    const item = {};

    const fn = get("FN");
    const n = get("N");
    if (n) {
      const parts = n.split(";");
      if (parts[0]) item.Surname = parts[0];
      if (parts[1]) item.GivenName = parts[1];
    } else if (fn) {
      const parts = fn.trim().split(/\s+/);
      item.GivenName = parts[0];
      if (parts.length > 1) item.Surname = parts.slice(1).join(" ");
    }
    if (fn) item.FileAs = fn;

    const org = get("ORG");
    if (org) {
      const parts = org.split(";");
      if (parts[0]) item.CompanyName = parts[0];
      if (parts[1]) item.Department = parts[1];
    }
    const title = get("TITLE");
    if (title) item.JobTitle = title;

    // emails — EWS requires a Key (EmailAddress1/2/3) on each entry
    const emails = getAll("EMAIL");
    if (emails.length) {
      const emailKeys = ["EmailAddress1", "EmailAddress2", "EmailAddress3"];
      item.EmailAddresses = emails.map((addr, i) => ({
        __type: "EmailAddressDictionaryEntry:#Exchange",
        Key: emailKeys[i] || "EmailAddress3",
        Address: addr,
        Name: fn || addr,
      }));
    }

    // phones
    const phoneEntries = props.filter(p => p.key === "TEL");
    if (phoneEntries.length) {
      item.PhoneNumbers = phoneEntries.map(p => {
        const type = (p.params.match(/TYPE=([^;:]+)/i) || [])[1] || "";
        const key = type.toLowerCase().includes("cell") || type.toLowerCase().includes("mobile") ? "MobilePhone"
          : type.toLowerCase().includes("home") ? "HomePhone"
          : type.toLowerCase().includes("work") || type.toLowerCase().includes("business") ? "BusinessPhone"
          : "OtherTelephone";
        return { __type: "PhoneNumberDictionaryEntry:#Exchange", Key: key, Value: p.value };
      });
    }

    const note = get("NOTE");
    if (note) item.Body = { __type: "BodyContentType:#Exchange", BodyType: "Text", Value: note };

    return item;
  },

  // EWS Contact item (from FindItem) -> vCard 4.0 string.
  // Unlike fromOwa (which parses Persona objects from FindPeople),
  // this parses actual EWS Contact store items with ItemId+ChangeKey.
  fromContact(c) {
    if (!c) return null;
    const L = ["BEGIN:VCARD", "VERSION:4.0"];
    const given = this.first(c, ["GivenName", "givenName"]);
    const sur = this.first(c, ["Surname", "surname"]);
    const fn = this.first(c, ["FileAs", "DisplayName", "fileAs"]) ||
      [given, sur].filter(Boolean).join(" ");
    if (fn) L.push("FN:" + this.esc(fn));
    L.push("N:" + this.esc(sur || "") + ";" + this.esc(given || "") + ";;;");

    const company = this.first(c, ["CompanyName", "companyName"]);
    const dept = this.first(c, ["Department", "department"]);
    if (company || dept) L.push("ORG:" + this.esc(company || "") + (dept ? ";" + this.esc(dept) : ""));
    const title = this.first(c, ["JobTitle", "Title", "jobTitle"]);
    if (title) L.push("TITLE:" + this.esc(title));

    // EmailAddresses — array of {Key, Address, Name}
    for (const e of this.arr(this.first(c, ["EmailAddresses", "emailAddresses"]))) {
      const addr = typeof e === "string" ? e : (e && (e.Address || e.Value || e.EmailAddress));
      if (addr) L.push("EMAIL:" + this.esc(String(addr)));
    }

    // PhoneNumbers — array of {Key, Value}
    for (const p of this.arr(this.first(c, ["PhoneNumbers", "phoneNumbers"]))) {
      const val = typeof p === "string" ? p : (p && (p.Value || p.Number));
      if (!val) continue;
      const key = String((p && p.Key) || "").toLowerCase();
      const type = key.includes("mobile") || key.includes("cell") ? "cell"
        : key.includes("home") ? "home"
        : key.includes("business") || key.includes("work") ? "work" : "voice";
      L.push("TEL;TYPE=" + type + ":" + this.esc(String(val)));
    }

    if (c.Body && c.Body.Value) L.push("NOTE:" + this.esc(c.Body.Value));

    // Use the real EWS ItemId (not PersonaId) for dedupe
    const itemId = c.ItemId && (c.ItemId.Id || c.ItemId.id);
    if (itemId) L.push("X-M365-OWA-ID:" + this.esc(String(itemId)));

    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Build a vCard from a Thunderbird property-bag (Map or plain object).
  // Used when contacts.get() returns no vCard field (legacy property-bag format).
  fromProperties(props) {
    if (!props) return null;
    const get = (key) => {
      let v;
      if (typeof props.get === "function") v = props.get(key);
      else v = props[key];
      return (v != null && v !== "") ? String(v) : null;
    };

    // Some TB versions store the raw vCard as a property — use it if present.
    const raw = get("vCard") || get("VCard");
    if (raw && raw.indexOf("BEGIN:VCARD") >= 0) return raw;

    const L = ["BEGIN:VCARD", "VERSION:4.0"];
    const given = get("FirstName") || get("GivenName");
    const sur = get("LastName") || get("Surname");
    const fn = get("DisplayName") || [given, sur].filter(Boolean).join(" ");
    if (fn) L.push("FN:" + this.esc(fn));
    L.push("N:" + this.esc(sur || "") + ";" + this.esc(given || "") + ";;;");

    const company = get("Company") || get("CompanyName");
    const dept = get("Department");
    if (company || dept) L.push("ORG:" + this.esc(company || "") + (dept ? ";" + this.esc(dept) : ""));
    const title = get("JobTitle") || get("Title");
    if (title) L.push("TITLE:" + this.esc(title));

    const email = get("PrimaryEmail");
    if (email) L.push("EMAIL:" + this.esc(email));
    const email2 = get("SecondEmail");
    if (email2) L.push("EMAIL:" + this.esc(email2));

    const mobile = get("CellularNumber") || get("MobilePhone");
    if (mobile) L.push("TEL;TYPE=cell:" + this.esc(mobile));
    const home = get("HomePhone");
    if (home) L.push("TEL;TYPE=home:" + this.esc(home));
    const work = get("WorkPhone") || get("BusinessPhone");
    if (work) L.push("TEL;TYPE=work:" + this.esc(work));

    const note = get("Notes");
    if (note) L.push("NOTE:" + this.esc(note));

    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Canonical form for comparison: unfolds line folding, strips
  // X-M365-OWA-ID and UID (TB auto-adds UID; server doesn't have it),
  // removes empty lines, sorts lines so order doesn't matter.
  vcardCanonical(vcard) {
    const lines = this._unfold(vcard).split("\n")
      .map(l => l.trim())
      .filter(l => l && l !== "BEGIN:VCARD" && l !== "END:VCARD" && !l.startsWith("VERSION:"))
      .filter(l => {
        const u = l.toUpperCase();
        return !u.startsWith("X-M365-OWA-ID:") && !u.startsWith("UID:");
      });
    lines.sort();
    return lines.join("\n");
  },

  // extract OWA id from vCard AND return the vCard (for comparison)
  vcardNormalized(vcard) {
    return String(vcard || "").replace(/\r/g, "").trim();
  }
};
