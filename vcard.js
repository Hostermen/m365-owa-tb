// vCard 4.0 conversion helpers for OWA contact items.
var VCard = {
  // Escape special characters in a vCard value (backslash, newline, comma, semicolon).
  esc(s) {
    if (s == null) return "";
    return String(s).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");
  },
  // Return the first non-null, non-empty value found under any of the given keys.
  first(obj, keys) { for (const k of keys) if (obj && obj[k] != null && obj[k] !== "") return obj[k]; return null; },
  // Wrap a value in an array if it isn't already; null/undefined -> empty array.
  arr(v) { return Array.isArray(v) ? v : (v ? [v] : []); },

  // Convert an OWA persona object to a vCard 4.0 string.
  fromOwa(p) {
    // null -> null
    if (!p) return null;
    // start building the vCard line array
    const L = ["BEGIN:VCARD", "VERSION:4.0"];
    // display name: prefer DisplayName/Alias/FullName, else assemble from given+surname
    const fn = this.first(p, ["DisplayName", "displayName", "Alias", "FullName"]) ||
      [this.first(p, ["GivenName", "FirstName"]), this.first(p, ["Surname", "LastName"])].filter(Boolean).join(" ");
    if (fn) L.push("FN:" + this.esc(fn));
    // given name and surname for N property
    const given = this.first(p, ["GivenName", "FirstName"]);
    const sur = this.first(p, ["Surname", "LastName"]);
    // N: surname;given;;;
    L.push("N:" + this.esc(sur || "") + ";" + this.esc(given || "") + ";;;");

    // company and department for ORG
    const company = this.first(p, ["CompanyName", "companyName", "Company"]);
    const dept = this.first(p, ["Department", "department"]);
    // job title
    const title = this.first(p, ["Title", "JobTitle", "title", "jobTitle"]);
    if (company || dept) L.push("ORG:" + this.esc(company || "") + (dept ? ";" + this.esc(dept) : ""));
    if (title) L.push("TITLE:" + this.esc(title));

    // collect emails from EmailAddresses array and single EmailAddress field
    const emails = new Set();
    for (const e of this.arr(this.first(p, ["EmailAddresses", "emailAddresses"]))) {
      // extract address from string or object
      const addr = typeof e === "string" ? e : (e && (e.EmailAddress || e.Address || e.address || e.email));
      if (addr) emails.add(String(addr));
    }
    // also try the single-address field
    const single = this.first(p, ["EmailAddress", "emailAddress", "PrimarySmtpAddress"]);
    if (single) emails.add(String(single));
    // add each email as a vCard EMAIL property
    for (const e of emails) L.push("EMAIL:" + this.esc(e));

    // mobile phone
    const mobile = this.first(p, ["MobilePhoneNumber", "MobilePhone", "mobilePhone"]);
    if (mobile) L.push("TEL;TYPE=cell:" + this.esc(mobile));
    // home phones
    for (const ph of this.arr(this.first(p, ["HomePhones", "HomePhoneNumber"]))) {
      const v = typeof ph === "string" ? ph : (ph && ph.Number);
      if (v) L.push("TEL;TYPE=home:" + this.esc(v));
    }
    // business/work phones
    for (const ph of this.arr(this.first(p, ["BusinessPhones", "BusinessPhoneNumber", "WorkPhone"]))) {
      const v = typeof ph === "string" ? ph : (ph && ph.Number);
      if (v) L.push("TEL;TYPE=work:" + this.esc(v));
    }

    // business address (best-effort)
    const ba = this.first(p, ["BusinessAddress", "WorkAddress", "OfficeLocation"]);
    if (ba && typeof ba === "object") {
      // assemble address parts
      const parts = [ba.Street || ba.StreetAddress || "", ba.City || "", ba.State || "", ba.PostalCode || "", ba.Country || ba.CountryOrRegion || ""];
      if (parts.some(Boolean)) L.push("ADR;TYPE=work:;;" + parts.map(this.esc).join(";"));
    }
    // if address is a plain string
    if (ba && typeof ba === "string" && ba) L.push("ADR;TYPE=work:;;" + this.esc(ba) + ";;;;");

    // notes
    const note = this.first(p, ["Notes", "personalNotes", "PersonalNotes"]);
    if (note) L.push("NOTE:" + this.esc(note));

    // stable id for dedupe: prefer PersonaId/ExchangeGuid, else synthesise from email+name
    let pid = this.first(p, ["PersonaId", "Id", "ExchangeGuid", "GuidId"]);
    // if the id is an object, drill into it
    if (pid && typeof pid === "object") pid = pid.Id || pid.id || pid.GuidId || null;
    // if no id, synthesise a hash from the primary email or display name
    if (!pid) {
      const seed = (single && String(single)) || fn || "";
      let h = 0; for (let i = 0; i < seed.length; i++) { h = ((h << 5) - h + seed.charCodeAt(i)) | 0; }
      pid = "syn-" + (h >>> 0).toString(16);
    }
    // store the stable id as a custom property
    L.push("X-M365-OWA-ID:" + this.esc(String(pid)));

    // close the vCard
    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Unfold vCard line folding (RFC 6350 §3.2): lines starting with space/tab continue the previous line.
  _unfold(vcard) {
    return String(vcard || "")
      // normalise line endings to \n
      .replace(/\r\n/g, "\n").replace(/\r/g, "\n")
      // join continuation lines
      .replace(/\n[ \t]/g, "");
  },

  // Extract the stable X-M365-OWA-ID from a vCard string (for dedupe).
  idFromVcard(vcard) {
    // match the X-M365-OWA-ID line
    const m = this._unfold(vcard).match(/^X-M365-OWA-ID:(.+)$/m);
    return m ? m[1].trim() : null;
  },

  // Replace the X-M365-OWA-ID in a vCard with a new value (after a server-side delete+create).
  rewriteOwaId(vcard, newOwaId) {
    if (!vcard || !newOwaId) return null;
    // unfold the vCard
    const unfolded = this._unfold(vcard);
    // if no X-M365-OWA-ID line exists, return null
    if (!/^X-M365-OWA-ID:/m.test(unfolded)) return null;
    // replace the old id with the new one
    const replaced = unfolded.replace(/^X-M365-OWA-ID:.+$/m, "X-M365-OWA-ID:" + this.esc(newOwaId));
    // restore \r\n line endings
    return replaced.replace(/\n/g, "\r\n");
  },

  // Parse vCard lines into an array of { key, params, value } objects (case-insensitive, handles folding).
  _parse(vcard) {
    // unfold and split into lines
    const lines = this._unfold(vcard).split("\n");
    // result array
    const props = [];
    for (const line of lines) {
      // skip empty lines, BEGIN/END markers, and VERSION
      if (!line || line === "BEGIN:VCARD" || line === "END:VCARD" || line.startsWith("VERSION:")) continue;
      // find the first colon separating name from value
      const colon = line.indexOf(":");
      if (colon < 0) continue;
      // property name (may include params before semicolon)
      const name = line.slice(0, colon);
      // property value
      const value = line.slice(colon + 1);
      // find the first semicolon separating name from params
      const semi = name.indexOf(";");
      // property key (uppercase)
      const key = (semi >= 0 ? name.slice(0, semi) : name).toUpperCase();
      // params string (after semicolon)
      const params = semi >= 0 ? name.slice(semi + 1) : "";
      // unescape the value and push
      props.push({ key, params, value: this._unesc(value) });
    }
    return props;
  },

  // Unescape vCard value: reverse the esc() escaping.
  _unesc(s) { return String(s || "").replace(/\\n/g, "\n").replace(/\\,/g, ",").replace(/\\;/g, ";").replace(/\\\\/g, "\\"); },

  // Parse a vCard string into an OWA Contact item for CreateItem/UpdateItem.
  vcardToOwa(vcard) {
    // parse all properties
    const props = this._parse(vcard);
    // get the first value for a key
    const get = (key) => { const p = props.find(p => p.key === key); return p ? p.value : null; };
    // get all values for a key
    const getAll = (key) => props.filter(p => p.key === key).map(p => p.value);
    // build the OWA item
    const item = {};

    // FN and N for name fields
    const fn = get("FN");
    const n = get("N");
    if (n) {
      // N: surname;given;;;
      const parts = n.split(";");
      if (parts[0]) item.Surname = parts[0];
      if (parts[1]) item.GivenName = parts[1];
    } else if (fn) {
      // no N -> split FN into given + surname
      const parts = fn.trim().split(/\s+/);
      item.GivenName = parts[0];
      if (parts.length > 1) item.Surname = parts.slice(1).join(" ");
    }
    // FileAs = display name
    if (fn) item.FileAs = fn;

    // ORG for company/department
    const org = get("ORG");
    if (org) {
      const parts = org.split(";");
      if (parts[0]) item.CompanyName = parts[0];
      if (parts[1]) item.Department = parts[1];
    }
    // job title
    const title = get("TITLE");
    if (title) item.JobTitle = title;

    // emails — EWS requires a Key (EmailAddress1/2/3) on each entry
    const emails = getAll("EMAIL");
    if (emails.length) {
      const emailKeys = ["EmailAddress1", "EmailAddress2", "EmailAddress3"];
      // map each email to a dictionary entry
      item.EmailAddresses = emails.map((addr, i) => ({
        __type: "EmailAddressDictionaryEntry:#Exchange",
        Key: emailKeys[i] || "EmailAddress3",
        Address: addr,
        Name: fn || addr,
      }));
    }

    // phone numbers — each needs a Key based on TYPE param
    const phoneEntries = props.filter(p => p.key === "TEL");
    if (phoneEntries.length) {
      item.PhoneNumbers = phoneEntries.map(p => {
        // extract TYPE from params
        const type = (p.params.match(/TYPE=([^;:]+)/i) || [])[1] || "";
        // map type to EWS phone key
        const key = type.toLowerCase().includes("cell") || type.toLowerCase().includes("mobile") ? "MobilePhone"
          : type.toLowerCase().includes("home") ? "HomePhone"
          : type.toLowerCase().includes("work") || type.toLowerCase().includes("business") ? "BusinessPhone"
          : "OtherTelephone";
        return { __type: "PhoneNumberDictionaryEntry:#Exchange", Key: key, Value: p.value };
      });
    }

    // notes -> Body (text type)
    const note = get("NOTE");
    if (note) item.Body = { __type: "BodyContentType:#Exchange", BodyType: "Text", Value: note };

    return item;
  },

  // Convert an EWS Contact item (from FindItem) to a vCard 4.0 string.
  fromContact(c) {
    // null -> null
    if (!c) return null;
    // start building the vCard
    const L = ["BEGIN:VCARD", "VERSION:4.0"];
    // given name and surname
    const given = this.first(c, ["GivenName", "givenName"]);
    const sur = this.first(c, ["Surname", "surname"]);
    // display name: prefer FileAs/DisplayName, else assemble from given+surname
    const fn = this.first(c, ["FileAs", "DisplayName", "fileAs"]) ||
      [given, sur].filter(Boolean).join(" ");
    if (fn) L.push("FN:" + this.esc(fn));
    // N property
    L.push("N:" + this.esc(sur || "") + ";" + this.esc(given || "") + ";;;");

    // company and department
    const company = this.first(c, ["CompanyName", "companyName"]);
    const dept = this.first(c, ["Department", "department"]);
    if (company || dept) L.push("ORG:" + this.esc(company || "") + (dept ? ";" + this.esc(dept) : ""));
    // job title
    const title = this.first(c, ["JobTitle", "Title", "jobTitle"]);
    if (title) L.push("TITLE:" + this.esc(title));

    // email addresses (array of {Key, Address, Name})
    for (const e of this.arr(this.first(c, ["EmailAddresses", "emailAddresses"]))) {
      // extract address from string or object
      const addr = typeof e === "string" ? e : (e && (e.Address || e.Value || e.EmailAddress));
      if (addr) L.push("EMAIL:" + this.esc(String(addr)));
    }

    // phone numbers (array of {Key, Value})
    for (const p of this.arr(this.first(c, ["PhoneNumbers", "phoneNumbers"]))) {
      // extract value from string or object
      const val = typeof p === "string" ? p : (p && (p.Value || p.Number));
      if (!val) continue;
      // determine the TEL TYPE from the EWS key
      const key = String((p && p.Key) || "").toLowerCase();
      const type = key.includes("mobile") || key.includes("cell") ? "cell"
        : key.includes("home") ? "home"
        : key.includes("business") || key.includes("work") ? "work" : "voice";
      L.push("TEL;TYPE=" + type + ":" + this.esc(String(val)));
    }

    // notes from Body
    if (c.Body && c.Body.Value) L.push("NOTE:" + this.esc(c.Body.Value));

    // use the real EWS ItemId for dedupe
    const itemId = c.ItemId && (c.ItemId.Id || c.ItemId.id);
    if (itemId) L.push("X-M365-OWA-ID:" + this.esc(String(itemId)));

    // close the vCard
    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Build a vCard from a Thunderbird property-bag (Map or plain object).
  fromProperties(props) {
    if (!props) return null;
    // accessor that handles both Map and plain object
    const get = (key) => {
      let v;
      if (typeof props.get === "function") v = props.get(key);
      else v = props[key];
      return (v != null && v !== "") ? String(v) : null;
    };

    // if the property bag already contains a raw vCard, use it
    const raw = get("vCard") || get("VCard");
    if (raw && raw.indexOf("BEGIN:VCARD") >= 0) return raw;

    // build the vCard from individual properties
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

    // emails
    const email = get("PrimaryEmail");
    if (email) L.push("EMAIL:" + this.esc(email));
    const email2 = get("SecondEmail");
    if (email2) L.push("EMAIL:" + this.esc(email2));

    // phones
    const mobile = get("CellularNumber") || get("MobilePhone");
    if (mobile) L.push("TEL;TYPE=cell:" + this.esc(mobile));
    const home = get("HomePhone");
    if (home) L.push("TEL;TYPE=home:" + this.esc(home));
    const work = get("WorkPhone") || get("BusinessPhone");
    if (work) L.push("TEL;TYPE=work:" + this.esc(work));

    // notes
    const note = get("Notes");
    if (note) L.push("NOTE:" + this.esc(note));

    L.push("END:VCARD");
    return L.join("\r\n");
  },

  // Produce a canonical form for comparison: unfold, strip X-M365-OWA-ID/UID, remove empty lines, sort.
  vcardCanonical(vcard) {
    // unfold, split, trim, filter, and sort lines
    const lines = this._unfold(vcard).split("\n")
      .map(l => l.trim())
      // remove BEGIN/END/VERSION
      .filter(l => l && l !== "BEGIN:VCARD" && l !== "END:VCARD" && !l.startsWith("VERSION:"))
      // remove X-M365-OWA-ID and UID (TB adds UID; server doesn't have it)
      .filter(l => {
        const u = l.toUpperCase();
        return !u.startsWith("X-M365-OWA-ID:") && !u.startsWith("UID:");
      });
    // sort so order doesn't matter
    lines.sort();
    return lines.join("\n");
  },

  // Normalise a vCard string for comparison (strip \r, trim).
  vcardNormalized(vcard) {
    return String(vcard || "").replace(/\r/g, "").trim();
  }
};
