// jCal (RFC 7265 JSON form of iCalendar) <-> OWA calendar item converters.
//
// OWA event fields (as returned by /owa/service.svc GetCalendarView) are flat:
// Start/End are ISO strings, Subject is a string, Body is {BodyType, Value},
// Location/Organizer may be strings or {DisplayName}/{Mailbox:{...}}. This
// mirrors the normalisation in m365-owa-cli's owa/normalize.py.
var JCal = {
  _first(obj, keys) { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; },
  // Local IANA timezone (e.g. "Europe/Berlin") — matches OWA's mailbox timezone
  // for a user sitting at their own machine.
  _localTz() {
    try { return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"; }
    catch { return "UTC"; }
  },
  _text(v) {
    if (v == null) return null;
    if (typeof v === "string") return v;
    if (typeof v === "object") {
      for (const k of ["DisplayName", "displayName", "Name", "name", "EmailAddress", "emailAddress"]) if (v[k]) return String(v[k]);
      const mb = v.Mailbox || v.mailbox;
      if (mb && typeof mb === "object") return this._text(mb);
      return String(v);
    }
    return String(v);
  },
  _body(owa) {
    const b = this._first(owa, ["Body", "body"]);
    if (!b) return { text: this._first(owa, ["BodyPreview", "bodyPreview", "Preview", "preview"]), type: "text" };
    if (typeof b === "string") return { text: b, type: "text" };
    const type = String(this._first(b, ["BodyType", "bodyType", "contentType", "type"]) || "text").toLowerCase();
    const text = this._first(b, ["Value", "value", "content", "text"]) || "";
    return { text: String(text), type: type.indexOf("html") >= 0 ? "html" : "text" };
  },
  _iso(v) {
    if (v == null) return null;
    if (v instanceof Date) {
      // Format as naive local time (no Z) — OWA returns times in its timezone
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      const hh = String(v.getHours()).padStart(2, "0");
      const mm = String(v.getMinutes()).padStart(2, "0");
      const ss = String(v.getSeconds()).padStart(2, "0");
      return `${y}-${m}-${d}T${hh}:${mm}:${ss}`;
    }
    const s = String(v);
    if (!s) return null;
    // OWA returns naive datetime strings (in the TimeZoneContext timezone).
    // Strip any Z/offset so Thunderbird uses the tzid we attach in eventToJcal.
    if (s.includes("T")) return s.replace(/Z$/, "").split("+")[0];
    return s + "T00:00:00";
  },

  _idStr(item) {
    if (!item) return null;
    const fid = item.ItemId;
    if (fid && typeof fid === "object" && fid.Id) return String(fid.Id);
    if (typeof fid === "string") return fid;
    const v = this._first(item, ["Id", "id", "UID", "Uid"]);
    return v != null ? String(v) : null;
  },

  // OWA item -> jCal vcalendar
  eventToJcal(item) {
    const props = [];
    const id = this._idStr(item);
    props.push(["uid", {}, "text", id || _uuid()]);
    const subject = this._first(item, ["Subject", "subject", "title", "Title"]);
    if (subject) props.push(["summary", {}, "text", String(subject)]);
    const body = this._body(item);
    if (body.text) props.push(["description", {}, "text", body.type === "html" ? this._stripHtml(body.text) : body.text]);
    const startRaw = this._first(item, ["Start", "start"]);
    const endRaw = this._first(item, ["End", "end"]);
    const allDay = !!this._first(item, ["IsAllDayEvent", "isAllDay", "IsAllDay"]);
    // For all-day events, extract the date part directly (no timezone conversion)
    if (allDay) {
      const sd = startRaw ? String(startRaw).split("T")[0] : null;
      const ed = endRaw ? String(endRaw).split("T")[0] : null;
      if (sd) props.push(["dtstart", {}, "date", sd]);
      if (ed) props.push(["dtend", {}, "date", ed]);
    } else {
      const start = this._iso(startRaw);
      const end = this._iso(endRaw);
      const tz = this._localTz();
      if (start) props.push(["dtstart", { tzid: tz }, "date-time", start]);
      if (end) props.push(["dtend", { tzid: tz }, "date-time", end]);
    }
    const loc = this._text(this._first(item, ["Location", "location"]));
    if (loc) props.push(["location", {}, "text", loc]);
    const cats = this._first(item, ["Categories", "categories"]);
    if (Array.isArray(cats) && cats.length) props.push(["categories", {}, "text", cats.map(String)]);
    const cancelled = this._first(item, ["IsCancelled", "isCancelled"]);
    props.push(["status", {}, "text", cancelled ? "CANCELLED" : "CONFIRMED"]);
    const sens = String(this._first(item, ["Sensitivity", "sensitivity"]) || "").toLowerCase();
    if (sens === "private") props.push(["class", {}, "text", "PRIVATE"]);
    else if (sens === "confidential") props.push(["class", {}, "text", "CONFIDENTIAL"]);
    else props.push(["class", {}, "text", "PUBLIC"]);
    const org = this._text(this._first(item, ["Organizer", "organizer"]));
    if (org && /@/.test(org)) props.push(["organizer", { cn: org }, "cal-address", "mailto:" + org.replace(/^mailto:/i, "")]);
    const atts = this._first(item, ["Attendees", "attendees"]);
    if (Array.isArray(atts)) for (const a of atts) {
      const ea = this._text(a) || "";
      if (!ea || !/@/.test(ea)) continue;
      const ps = String(this._first(a, ["ResponseType", "response"]) || "").toLowerCase();
      const partstat = ({ accepted: "ACCEPTED", declined: "DECLINED", tentative: "TENTATIVE" }[ps]) || "NEEDS-ACTION";
      props.push(["attendee", { cn: this._text(a) || ea, partstat }, "cal-address", "mailto:" + ea.replace(/^mailto:/i, "")]);
    }
    const ml = this._first(item, ["meeting_link", "meetingLink"]) || (item.onlineMeeting && (item.onlineMeeting.joinUrl || item.onlineMeeting.meetingLink));
    if (ml) props.push(["x-m365-meeting-link", {}, "text", String(ml)]);
    return ["vcalendar", [["version", {}, "text", "2.0"], ["prodid", {}, "text", "-//M365OWASync//EN"]], [["vevent", props, []]]];
  },

  // jCal (from the experiment) -> OWA CreateItem item shape
  jcalToOwa(jcal) {
    const comp = Array.isArray(jcal) && jcal[0] === "vcalendar" ? new _JCalComp(jcal) : null;
    if (!comp) return null;
    const ve = comp.sub("vevent");
    if (!ve) return null;
    const item = {};
    const uid = ve.prop("uid"); if (uid) item.Id = String(uid);
    const summary = ve.prop("summary"); if (summary != null) item.Subject = String(summary);
    const desc = ve.prop("description");
    if (desc != null) item.Body = { __type: "BodyContentType:#Exchange", BodyType: "Text", Value: String(desc) };
    const loc = ve.prop("location"); if (loc != null) item.Location = String(loc);
    const startProp = ve.allProps("dtstart")[0];
    const endProp = ve.allProps("dtend")[0];
    if (startProp) item.Start = this._fmtDate(startProp);
    if (endProp) item.End = this._fmtDate(endProp);
    if (startProp && startProp[2] === "date") item.IsAllDayEvent = true;
    const cats = ve.prop("categories");
    if (cats) item.Categories = String(cats).split(",").map(s => s.trim()).filter(Boolean);
    return item;
  },

  _fmtDate(prop) {
    const v = String(prop[3]);
    if (prop[2] === "date") return v + "T00:00:00.000";
    // Send naive datetime as-is. OWA interprets it in the TimeZoneContext
    // timezone (the user's local TZ), so no conversion needed.
    return v.replace(/Z$/, "").split("+")[0].split(".")[0] + ".000";
  },
  _stripHtml(h) {
    const q = String.fromCharCode(34), sq = String.fromCharCode(39);
    return String(h || "")
      .replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&#0?39;/g, sq)
      .replace(/&quot;/g, q)
      .trim();
  },

  _uuid() { return crypto.randomUUID ? crypto.randomUUID() : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => { const r = crypto.getRandomValues(new Uint8Array(1))[0] % 16; return (c === "x" ? r : (r & 0x3) | 0x8).toString(16); }); }
};

// minimal jCal accessor (shared with the Graph variant)
class _JCalComp {
  constructor(jcal) { this.name = jcal[0]; this.props = jcal[1] || []; this.subs = (jcal[2] || []).map(s => new _JCalComp(s)); }
  sub(name) { return this.subs.find(s => s.name === name); }
  prop(name) { const p = this.props.find(p => p[0].toLowerCase() === name.toLowerCase()); return p ? p[3] : null; }
  allProps(name) { return this.props.filter(p => p[0].toLowerCase() === name.toLowerCase()); }
}
function _uuid() { return JCal._uuid(); }
