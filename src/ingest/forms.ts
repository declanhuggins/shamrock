// @ts-nocheck
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

Shamrock.onFormSubmit = function (e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName().toLowerCase();
  if (sheetName.includes("cadet")) {
    handleCadetForm(e);
  } else if (sheetName.includes("event")) {
    handleEventForm(e);
  } else if (sheetName.includes("excusal")) {
    handleExcusalForm(e);
  } else {
    // Unknown form; ignore
  }
};

function handleCadetForm(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const data = e.namedValues || {};
    const email = pick(data, ["University Email", "Email", "cadet_email", "University Email Address"]);
    if (!email) throw new Error("Cadet form missing email");

    const phoneRaw = pick(data, ["Phone", "phone"]) || "";
    const phoneDigits = phoneRaw.replace(/\D+/g, "").slice(0, 10);

    const stateRaw = pick(data, ["Home State", "home_state"]) || "";
    const stateFull = normalizeState(stateRaw);

    const cipBroadRaw = pick(data, ["CIP Broad", "cip_broad"]) || "";
    const cipBroad = cipBroadRaw.split(" - ")[0].trim() ? cipBroadRaw : cipBroadRaw;
    const cipCodeRaw = pick(data, ["CIP Code", "cip_code"]) || "";
    const cipCode = cipCodeRaw.split(" - ")[0].trim() || cipCodeRaw;

    const fpsRaw = pick(data, ["Flight Path Status", "flight_path_status"]) || "";
    const fps = normalizeFlightPathStatus(fpsRaw);

    const record = {
      cadet_email: email.trim().toLowerCase(),
      last_name: pick(data, ["Last Name", "last_name"]) || "",
      first_name: pick(data, ["First Name", "first_name"]) || "",
      as_year: pick(data, ["AS Year", "as_year"]) || "",
      graduation_year: pick(data, ["Graduation Year", "Class/Graduation Year", "graduation_year"]) || "",
      flight: pick(data, ["Flight", "flight"]) || "",
      squadron: pick(data, ["Squadron", "squadron"]) || "",
      university: pick(data, ["University", "university"]) || "",
      dorm: pick(data, ["Dorm", "dorm"]) || "",
      home_town: pick(data, ["Home Town", "home_town"]) || "",
      home_state: stateFull,
      phone: phoneDigits,
      dob: pick(data, ["DOB", "Birthday", "dob"]) || "",
      cip_broad: cipBroad,
      cip_code: cipCode,
      afsc: pick(data, ["AFSC", "afsc"]) || "",
      flight_path_status: fps,
      status: pick(data, ["Status", "status"]) || "Active",
      photo_url: pick(data, ["Photo Link", "photo_url", "Photo URL"]) || "",
      notes: pick(data, ["Notes", "notes"]) || "",
      created_at: Shamrock.nowIso(),
      updated_at: Shamrock.nowIso(),
    } as any;
    Shamrock.upsertCadet(record);
    Shamrock.logAudit({
      action: "directory.upsert",
      target_table: "Directory Backend",
      target_key: record.cadet_email,
      new_value: JSON.stringify(record),
      source: "form",
    });

    // Propagate to public views so Directory/Attendance pick up new cadets immediately
    try {
      Shamrock.rebuildDirectory();
      Shamrock.rebuildAttendance();
    } catch (err) {
      // If frontend is not configured yet, skip quietly
    }
  });
}

function handleEventForm(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const data = e.namedValues || {};
    const eventId = pick(data, ["event_id", "Event ID", "Event Key"]) || buildEventId(data);
    const record = {
      event_id: eventId,
      event_name: pick(data, ["Event Name", "event_name", "Display Name"]) || eventId,
      event_type: pick(data, ["Event Type", "event_type"]) || "",
      training_week: pick(data, ["Training Week", "training_week"]) || "",
      event_date: pick(data, ["Event Date", "Date", "event_date"]) || "",
      event_status: pick(data, ["Status", "event_status"]) || "Published",
      affects_attendance: pick(data, ["Affects Attendance", "affects_attendance"]) || "TRUE",
      attendance_label: pick(data, ["Attendance Label", "attendance_label"]) || pick(data, ["Event Name", "event_name"]) || eventId,
      expected_group: pick(data, ["Expected Group", "expected_group"]) || "",
      flight_scope: pick(data, ["Flight Scope", "flight_scope"]) || "All",
      location: pick(data, ["Location", "location"]) || "",
      notes: pick(data, ["Notes", "notes"]) || "",
      created_at: Shamrock.nowIso(),
      updated_at: Shamrock.nowIso(),
    } as any;
    Shamrock.upsertEvent(record);
    Shamrock.logAudit({
      action: "events.upsert",
      target_table: "Events Backend",
      target_key: record.event_id,
      new_value: JSON.stringify(record),
      source: "form",
    });
  });
}

function handleExcusalForm(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const data = e.namedValues || {};
    const cadetEmail = (pick(data, ["Email", "University Email", "cadet_email"]) || "").toLowerCase();
    const eventId = pick(data, ["Event", "Event ID", "event_id"]) || "";
    if (!cadetEmail || !eventId) throw new Error("Excusal form missing email or event");
    const excusalId = `EXC-${Utilities.getUuid()}`;
    const attendanceEffect = pick(data, ["Attendance Effect", "attendance_effect"]) || "Set ER";
    const decision = pick(data, ["Decision", "decision"]) || "";
    const record = {
      excusal_id: excusalId,
      cadet_email: cadetEmail,
      event_id: eventId,
      request_timestamp: Shamrock.nowIso(),
      reason: pick(data, ["Reason", "reason"]) || "",
      decision,
      decision_by: pick(data, ["Decision By", "decision_by"]) || "",
      decision_timestamp: decision ? Shamrock.nowIso() : "",
      attendance_effect: attendanceEffect,
      source: "form",
    } as any;
    Shamrock.appendExcusal(record);
    applyAttendanceEffect(cadetEmail, eventId, attendanceEffect);
    Shamrock.logAudit({
      action: "excusals.submit",
      target_table: "Excusals Backend",
      target_key: excusalId,
      new_value: JSON.stringify(record),
      source: "form",
    });
  });
}

function applyAttendanceEffect(cadetEmail: string, eventId: string, effect: string): void {
  const normalized = effect.toLowerCase();
  let code = "";
  if (normalized.includes("er")) code = "ER";
  else if (normalized === "set e" || normalized.includes(" approved")) code = "E";
  else if (normalized.includes("ed")) code = "ED";
  if (!code) return;
  Shamrock.setAttendance({
    cadet_email: cadetEmail,
    event_id: eventId,
    attendance_code: code,
    source: "form",
    updated_at: Shamrock.nowIso(),
  } as any);
  Shamrock.logAudit({
    action: "attendance.set_code",
    target_table: "Attendance Backend",
    target_key: `${cadetEmail}|${eventId}`,
    event_id: eventId,
    new_value: code,
    source: "form",
  });
}

function normalizeState(input: string): string {
  const s = String(input || "").trim();
  if (!s) return "";
  const upper = s.toUpperCase();
  const STATES: Record<string, string> = {
    AL: "Alabama", AK: "Alaska", AZ: "Arizona", AR: "Arkansas", CA: "California", CO: "Colorado", CT: "Connecticut", DE: "Delaware", FL: "Florida", GA: "Georgia", HI: "Hawaii", ID: "Idaho", IL: "Illinois", IN: "Indiana", IA: "Iowa", KS: "Kansas", KY: "Kentucky", LA: "Louisiana", ME: "Maine", MD: "Maryland", MA: "Massachusetts", MI: "Michigan", MN: "Minnesota", MS: "Mississippi", MO: "Missouri", MT: "Montana", NE: "Nebraska", NV: "Nevada", NH: "New Hampshire", NJ: "New Jersey", NM: "New Mexico", NY: "New York", NC: "North Carolina", ND: "North Dakota", OH: "Ohio", OK: "Oklahoma", OR: "Oregon", PA: "Pennsylvania", RI: "Rhode Island", SC: "South Carolina", SD: "South Dakota", TN: "Tennessee", TX: "Texas", UT: "Utah", VT: "Vermont", VA: "Virginia", WA: "Washington", WV: "West Virginia", WI: "Wisconsin", WY: "Wyoming", DC: "District of Columbia",
  };
  if (STATES[upper]) return STATES[upper];
  const normalized = s.replace(/\./g, "").toLowerCase();
  const match = Object.values(STATES).find(name => name.toLowerCase() === normalized);
  return match || s;
}

function normalizeFlightPathStatus(input: string): string {
  const val = (input || "").toLowerCase();
  if (val.includes("ready")) return "Ready 4/4";
  if (val.includes("active")) return "Active 3/4";
  if (val.includes("enrolled")) return "Enrolled 2/4";
  if (val.includes("participating")) return "Participating 1/4";
  return input || "";
}

function pick(namedValues: { [key: string]: any }, keys: string[]): string | undefined {
  for (const key of keys) {
    const match = namedValues[key];
    if (match == null) continue;
    const value = Array.isArray(match) ? match[0] : match;
    if (value === undefined || value === null) continue;
    const s = String(value).trim();
    if (s) return s;
  }
  // try case-insensitive lookup
  const lower = Object.keys(namedValues).reduce<Record<string, any>>((acc, k) => {
    acc[k.toLowerCase()] = namedValues[k];
    return acc;
  }, {});
  for (const key of keys) {
    const match = lower[key.toLowerCase()];
    if (match == null) continue;
    const value = Array.isArray(match) ? match[0] : match;
    if (value === undefined || value === null) continue;
    const s = String(value).trim();
    if (s) return s;
  }
  return undefined;
}

function buildEventId(namedValues: { [key: string]: any }): string {
  const term = pick(namedValues, ["Term", "term"]) || "";
  const tw = pick(namedValues, ["Training Week", "training_week"]) || "";
  const type = pick(namedValues, ["Event Type", "event_type"]) || "";
  const parts = [term, tw, type].filter(Boolean);
  if (parts.length) return parts.join("-").replace(/\s+/g, "");
  return `EVT-${Utilities.getUuid()}`;
}
