// jCal (RFC 7265) <-> OWA calendar item converters.
var JCal = {
  // Return the first non-null value found under any of the given keys.
  _first(obj, keys) { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; },

  // Normalise a value to a display string, handling string / object / Mailbox shapes.
  _text(v) {
    // null/undefined -> null
    if (v == null) return null;
    // already a string -> return as-is
    if (typeof v === "string") return v;
    // object: try common display-name keys, then recurse into Mailbox
    if (typeof v === "object") {
      for (const k of ["DisplayName", "displayName", "Name", "name", "EmailAddress", "emailAddress"]) if (v[k]) return String(v[k]);
      // try nested Mailbox object
      const mb = v.Mailbox || v.mailbox;
      // recurse into the Mailbox sub-object
      if (mb && typeof mb === "object") return this._text(mb);
      // last resort: stringify
      return String(v);
    }
    // numbers etc. -> stringify
    return String(v);
  },

  // Extract body text and type from an OWA item's Body field.
  _body(owa) {
    // find the Body field (camelCase or PascalCase)
    const b = this._first(owa, ["Body", "body"]);
    // if no Body, fall back to a preview field as plain text
    if (!b) return { text: this._first(owa, ["BodyPreview", "bodyPreview", "Preview", "preview"]), type: "text" };
    // if Body is a plain string, return it as text
    if (typeof b === "string") return { text: b, type: "text" };
    // determine the content type (html or text)
    const type = String(this._first(b, ["BodyType", "bodyType", "contentType", "type"]) || "text").toLowerCase();
    // extract the actual body value
    const text = this._first(b, ["Value", "value", "content", "text"]) || "";
    // normalise type to "html" or "text"
    return { text: String(text), type: type.indexOf("html") >= 0 ? "html" : "text" };
  },

  // Normalise an OWA datetime string to a UTC ISO string.
  _iso(v) {
    // null -> null
    if (v == null) return null;
    // Date object -> ISO string
    if (v instanceof Date) return v.toISOString();
    // coerce to string
    const s = String(v);
    // empty -> null
    if (!s) return null;
    // if it contains a time component
    if (s.includes("T")) {
      // already has UTC or offset -> return as-is
      if (s.endsWith("Z") || /[+-]\d{2}:?\d{2}$/.test(s)) return s;
      // naive datetime -> append Z (we send TimeZoneContext=UTC so it's UTC)
      return s + "Z";
    }
    // date-only -> add midnight UTC
    return s + "T00:00:00Z";
  },

  // Extract the string ItemId from an OWA event item (handles object and string shapes).
  _idStr(item) {
    // null item -> null
    if (!item) return null;
    // ItemId is the primary field
    const fid = item.ItemId;
    // object form: { Id: "..." }
    if (fid && typeof fid === "object" && fid.Id) return String(fid.Id);
    // string form
    if (typeof fid === "string") return fid;
    // fallback: try alternate keys
    const v = this._first(item, ["Id", "id", "UID", "Uid"]);
    return v != null ? String(v) : null;
  },

  // Convert an OWA event item to a jCal vcalendar array.
  eventToJcal(item) {
    // jCal property list for the vevent
    const props = [];
    // extract the event id
    const id = this._idStr(item);
    // UID property (generate a UUID if absent)
    props.push(["uid", {}, "text", id || _uuid()]);
    // summary / subject
    const subject = this._first(item, ["Subject", "subject", "title", "Title"]);
    if (subject) props.push(["summary", {}, "text", String(subject)]);
    // body / description
    const body = this._body(item);
    // strip HTML if needed
    if (body.text) props.push(["description", {}, "text", body.type === "html" ? this._stripHtml(body.text) : body.text]);
    // start/end raw values
    const startRaw = this._first(item, ["Start", "start"]);
    const endRaw = this._first(item, ["End", "end"]);
    // all-day flag
    const allDay = !!this._first(item, ["IsAllDayEvent", "isAllDay", "IsAllDay"]);
    // all-day events: use date-only (no timezone conversion)
    if (allDay) {
      // extract date part before "T"
      const sd = startRaw ? String(startRaw).split("T")[0] : null;
      const ed = endRaw ? String(endRaw).split("T")[0] : null;
      if (sd) props.push(["dtstart", {}, "date", sd]);
      if (ed) props.push(["dtend", {}, "date", ed]);
    } else {
      // timed events: normalise to UTC ISO strings
      const start = this._iso(startRaw);
      const end = this._iso(endRaw);
      if (start) props.push(["dtstart", {}, "date-time", start]);
      if (end) props.push(["dtend", {}, "date-time", end]);
    }
    // location
    const loc = this._text(this._first(item, ["Location", "location"]));
    if (loc) props.push(["location", {}, "text", loc]);
    // categories
    const cats = this._first(item, ["Categories", "categories"]);
    if (Array.isArray(cats) && cats.length) props.push(["categories", {}, "text", cats.map(String)]);
    // cancelled / confirmed status
    const cancelled = this._first(item, ["IsCancelled", "isCancelled"]);
    props.push(["status", {}, "text", cancelled ? "CANCELLED" : "CONFIRMED"]);
    // sensitivity / classification
    const sens = String(this._first(item, ["Sensitivity", "sensitivity"]) || "").toLowerCase();
    if (sens === "private") props.push(["class", {}, "text", "PRIVATE"]);
    else if (sens === "confidential") props.push(["class", {}, "text", "CONFIDENTIAL"]);
    else props.push(["class", {}, "text", "PUBLIC"]);
    // organizer
    const org = this._text(this._first(item, ["Organizer", "organizer"]));
    // only add if it looks like an email
    if (org && /@/.test(org)) props.push(["organizer", { cn: org }, "cal-address", "mailto:" + org.replace(/^mailto:/i, "")]);
    // attendees
    const atts = this._first(item, ["Attendees", "attendees"]);
    // iterate each attendee
    if (Array.isArray(atts)) for (const a of atts) {
      // extract email address
      const ea = this._text(a) || "";
      // skip if no email
      if (!ea || !/@/.test(ea)) continue;
      // map response type to partstat
      const ps = String(this._first(a, ["ResponseType", "response"]) || "").toLowerCase();
      const partstat = ({ accepted: "ACCEPTED", declined: "DECLINED", tentative: "TENTATIVE" }[ps]) || "NEEDS-ACTION";
      // add attendee property
      props.push(["attendee", { cn: this._text(a) || ea, partstat }, "cal-address", "mailto:" + ea.replace(/^mailto:/i, "")]);
    }
    // online meeting link
    const ml = this._first(item, ["meeting_link", "meetingLink"]) || (item.onlineMeeting && (item.onlineMeeting.joinUrl || item.onlineMeeting.meetingLink));
    if (ml) props.push(["x-m365-meeting-link", {}, "text", String(ml)]);
    // assemble the full vcalendar
    return ["vcalendar", [["version", {}, "text", "2.0"], ["prodid", {}, "text", "-//M365OWASync//EN"]], [["vevent", props, []]]];
  },

  // Convert a jCal vcalendar (from the experiment) to an OWA CreateItem item shape.
  jcalToOwa(jcal) {
    // wrap the raw jCal array in a _JCalComp accessor
    const comp = Array.isArray(jcal) && jcal[0] === "vcalendar" ? new _JCalComp(jcal) : null;
    // not a valid vcalendar -> null
    if (!comp) return null;
    // find the vevent sub-component
    const ve = comp.sub("vevent");
    // no vevent -> null
    if (!ve) return null;
    // build the OWA item object
    const item = {};
    // UID -> ItemId
    const uid = ve.prop("uid"); if (uid) item.Id = String(uid);
    // summary -> Subject
    const summary = ve.prop("summary"); if (summary != null) item.Subject = String(summary);
    // description -> Body (text type)
    const desc = ve.prop("description");
    if (desc != null) item.Body = { __type: "BodyContentType:#Exchange", BodyType: "Text", Value: String(desc) };
    // location
    const loc = ve.prop("location"); if (loc != null) item.Location = String(loc);
    // start/end properties
    const startProp = ve.allProps("dtstart")[0];
    const endProp = ve.allProps("dtend")[0];
    // format dates for OWA
    if (startProp) item.Start = this._fmtDate(startProp);
    if (endProp) item.End = this._fmtDate(endProp);
    // all-day if dtstart is a date (not date-time)
    if (startProp && startProp[2] === "date") item.IsAllDayEvent = true;
    // categories
    const cats = ve.prop("categories");
    if (cats) item.Categories = String(cats).split(",").map(s => s.trim()).filter(Boolean);
    return item;
  },

  // Format a jCal date property value as an OWA datetime string (UTC, with milliseconds).
  _fmtDate(prop) {
    // the raw value string
    const v = String(prop[3]);
    // date-only -> midnight with milliseconds
    if (prop[2] === "date") return v + "T00:00:00.000";
    // already UTC (Z suffix) -> strip Z and add .000
    if (/Z$/.test(v)) return v.replace(/Z$/, "").split(".")[0] + ".000";
    // explicit timezone offset -> parse and convert to UTC
    if (/[+-]\d{2}:?\d{2}$/.test(v)) {
      const d = new Date(v);
      if (!isNaN(d)) return d.toISOString().replace(/Z$/, "").split(".")[0] + ".000";
    }
    // naive datetime -> interpret as local time, convert to UTC
    const d = new Date(v);
    if (!isNaN(d)) return d.toISOString().replace(/Z$/, "").split(".")[0] + ".000";
    // fallback: strip any offset or Z
    return v.split("+")[0].replace(/Z$/, "").split(".")[0] + ".000";
  },

  // Strip HTML tags and decode common entities, returning plain text.
  _stripHtml(h) {
    // quote characters for entity replacement
    const q = String.fromCharCode(34), sq = String.fromCharCode(39);
    return String(h || "")
      // remove tags
      .replace(/<[^>]*>/g, "")
      // decode entities
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&#0?39;/g, sq)
      .replace(/&quot;/g, q)
      .trim();
  },

  // Generate a RFC 4122 v4 UUID (uses crypto.randomUUID if available, else manual fallback).
  _uuid() { return crypto.randomUUID ? crypto.randomUUID() : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => { const r = crypto.getRandomValues(new Uint8Array(1))[0] % 16; return (c === "x" ? r : (r & 0x3) | 0x8).toString(16); }); }
};

// Minimal jCal component accessor shared across sync modules.
class _JCalComp {
  // Parse a raw jCal array into name, props, and sub-components.
  constructor(jcal) { this.name = jcal[0]; this.props = jcal[1] || []; this.subs = (jcal[2] || []).map(s => new _JCalComp(s)); }
  // Find the first sub-component with the given name.
  sub(name) { return this.subs.find(s => s.name === name); }
  // Get a property value by name (case-insensitive).
  prop(name) { const p = this.props.find(p => p[0].toLowerCase() === name.toLowerCase()); return p ? p[3] : null; }
  // Get all property entries matching the given name (case-insensitive).
  allProps(name) { return this.props.filter(p => p[0].toLowerCase() === name.toLowerCase()); }
}
// Module-level UUID delegate so eventToJcal can call _uuid() without JCal prefix.
function _uuid() { return JCal._uuid(); }
