// @ts-nocheck
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

Shamrock.onFormSubmit = function (e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo("form.submit", "dispatching form submission", { sheet: e && e.range ? e.range.getSheet().getName() : "unknown" });
  }
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName().toLowerCase();
  if (isAttendanceFormSheet(e.range.getSheet())) {
    handleAttendanceForm(e);
  } else if (sheetName.includes("cadet")) {
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
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo("form.cadet", "processing cadet form submission");
  }
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const data = e.namedValues || {};
    const email = pick(data, ["Email Address", "Email", "email", "cadet_email", "University Email", "University Email Address"]);
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
    if (typeof Shamrock.logInfo === "function") {
      Shamrock.logInfo("form.cadet", "upserted cadet record", { email: record.cadet_email });
    }
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
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo("form.event", "processing event form submission");
  }
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
    if (typeof Shamrock.logInfo === "function") {
      Shamrock.logInfo("form.event", "upserted event", { event_id: record.event_id });
    }
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
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo("form.excusal", "processing excusal submission");
  }
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const data = e.namedValues || {};
    const cadetEmail = (pick(data, ["Email Address", "Email", "email", "University Email", "cadet_email"]) || "").toLowerCase();
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
    if (typeof Shamrock.logInfo === "function") {
      Shamrock.logInfo("form.excusal", "recorded excusal", { excusal_id: excusalId, event_id: eventId });
    }
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

function handleAttendanceForm(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo("form.attendance", "processing attendance submission");
  }
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();

    const sheet = e.range.getSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || ""));
    const row = sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    const namedValues = e.namedValues || {};

    const trainingWeekRaw = pick(namedValues, ["Training Week (Format as TW-00):", "Training Week", "training_week", "Training Week (Format as TW-00)"]) || readHeaderValue(headers, row, ["Training Week (Format as TW-00):", "Training Week", "training_week"]);
    const eventRaw = pick(namedValues, ["Event:", "Event", "event", "Event Name", "Event Type"]) || readHeaderValue(headers, row, ["Event:", "Event", "Event Name", "Event Type"]);
    const eventIdRaw = pick(namedValues, ["Event ID", "Event Key", "event_id"]) || readHeaderValue(headers, row, ["Event ID", "Event Key", "event_id"]);
    const flightRaw = pick(namedValues, ["Flight:", "Flight", "flight"]) || readHeaderValue(headers, row, ["Flight:", "Flight", "flight"]);

    const normalizedWeek = normalizeTrainingWeekKey(trainingWeekRaw || "");
    const eventType = normalizeAttendanceEventType(eventRaw || "");
    const eventId = resolveAttendanceEventId(eventIdRaw || "", normalizedWeek, eventType);

    const tokens = collectAttendanceTokens(headers, row, eventType, flightRaw || "");
    if (!tokens.length) throw new Error("Attendance form contained no attendees to record.");

    const mapped = mapTokensToCadetEmails(tokens);
    if (!mapped.emails.length) throw new Error("No cadet emails matched this attendance submission.");

    mapped.emails.forEach(email => {
      Shamrock.setAttendance({
        cadet_email: email,
        event_id: eventId,
        attendance_code: "P",
        source: "form",
        updated_at: Shamrock.nowIso(),
      } as any);
      if (typeof Shamrock.logInfo === "function") {
        Shamrock.logInfo("form.attendance", "marked present", { email, event_id: eventId, event_type: eventType || "" });
      }
      Shamrock.logAudit({
        action: "attendance.present",
        target_table: "Attendance Backend",
        target_key: `${email}|${eventId}`,
        event_id: eventId,
        new_value: "P",
        source: "form",
        notes: `attendance_form event=${eventType || ""} week=${normalizedWeek || ""}`,
      });
    });

    if (mapped.missing.length) {
      Shamrock.logAudit({
        action: "attendance.present.missing",
        target_table: "Attendance Backend",
        target_key: eventId,
        event_id: eventId,
        new_value: "",
        result: "partial",
        notes: `Unmatched attendees: ${mapped.missing.join("; ")}`,
        source: "form",
      });
    }

    try {
      Shamrock.rebuildAttendance();
    } catch (err) {
      // If frontend not configured yet, skip quietly
    }
  });
}

function isAttendanceFormSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
  try {
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(v => String(v || ""));
    const markers = ["Training Week (Format as TW-00):", "Event:", "Flight:"];
    return markers.some(marker => headers.some(h => normalizeHeaderText(h) === normalizeHeaderText(marker)));
  } catch (err) {
    return false;
  }
}

function readHeaderValue(headers: string[], row: any[], candidates: string[]): string | undefined {
  for (let i = 0; i < headers.length; i++) {
    const header = normalizeHeaderText(headers[i]);
    if (!header) continue;
    if (candidates.some(c => normalizeHeaderText(c) === header)) {
      const val = row[i];
      if (val == null) continue;
      const s = String(Array.isArray(val) ? val[0] : val).trim();
      if (s) return s;
    }
  }
  return undefined;
}

function collectAttendanceTokens(headers: string[], row: any[], eventType: string, flightRaw: string): string[] {
  const metaHeaders = new Set([
    normalizeHeaderText("timestamp"),
    normalizeHeaderText("name"),
    normalizeHeaderText("Training Week (Format as TW-00):"),
    normalizeHeaderText("Training Week"),
    normalizeHeaderText("Event:"),
    normalizeHeaderText("Event"),
    normalizeHeaderText("Flight:"),
    normalizeHeaderText("Flight"),
    normalizeHeaderText("Event ID"),
    normalizeHeaderText("Event Key"),
  ]);

  const tokens: string[] = [];
  headers.forEach((header, idx) => {
    const norm = normalizeHeaderText(header);
    if (!norm || metaHeaders.has(norm)) return;
    tokens.push(...parseNamesOrEmails(row[idx]));
  });

  // If a specific flight was selected and appears as its own column, prefer that column's values
  if (flightRaw) {
    const targetFlight = normalizeHeaderText(flightRaw);
    headers.forEach((header, idx) => {
      if (normalizeHeaderText(header) === normalizeHeaderText(`${flightRaw} Flight`)) {
        tokens.push(...parseNamesOrEmails(row[idx]));
      }
      if (targetFlight && normalizeHeaderText(header) === targetFlight) {
        tokens.push(...parseNamesOrEmails(row[idx]));
      }
    });
  }

  // If event type implies cross-flight (e.g., Secondary), include any per-flight columns explicitly
  if (eventType === "secondary") {
    headers.forEach((header, idx) => {
      if (normalizeHeaderText(header).endsWith("flight")) {
        tokens.push(...parseNamesOrEmails(row[idx]));
      }
    });
  }

  return Array.from(new Set(tokens.filter(Boolean)));
}

function parseNamesOrEmails(val: any): string[] {
  if (val == null) return [];
  const raw = Array.isArray(val) ? val.join(" ") : String(val || "");
  return raw
    .split(/[\n;,]/)
    .map(s => s.trim())
    .filter(Boolean);
}

function mapTokensToCadetEmails(tokens: string[]): { emails: string[]; missing: string[] } {
  const cadets = Shamrock.listCadets();
  const emailSet = new Set<string>();
  const nameToEmail = new Map<string, string | null>();

  cadets.forEach(c => {
    const email = String((c as any).cadet_email || "").toLowerCase();
    if (email) emailSet.add(email);
    const last = String((c as any).last_name || "");
    const first = String((c as any).first_name || "");
    const key = normalizeNameKey(last, first);
    if (!key) return;
    if (nameToEmail.has(key) && nameToEmail.get(key) !== email) {
      nameToEmail.set(key, null); // ambiguous name
    } else {
      nameToEmail.set(key, email);
    }
  });

  const emails: string[] = [];
  const missing: string[] = [];

  tokens.forEach(token => {
    const lower = token.toLowerCase();
    if (lower.includes("@")) {
      const email = lower.trim();
      if (emailSet.has(email)) {
        emails.push(email);
      } else {
        missing.push(token);
      }
      return;
    }

    const parsed = parseNameToken(token);
    if (!parsed) {
      missing.push(token);
      return;
    }
    const key = normalizeNameKey(parsed.last, parsed.first);
    const email = key ? nameToEmail.get(key) : null;
    if (email) {
      emails.push(email);
    } else {
      missing.push(token);
    }
  });

  return { emails: Array.from(new Set(emails)), missing };
}

function normalizeNameKey(last: string, first: string): string {
  const l = String(last || "").trim();
  const f = String(first || "").trim();
  if (!l || !f) return "";
  return `${l.toLowerCase()},${f.toLowerCase()}`;
}

function parseNameToken(token: string): { first: string; last: string } | null {
  const s = String(token || "").trim();
  if (!s) return null;
  if (s.includes(",")) {
    const parts = s.split(",");
    const last = parts[0] || "";
    const first = parts.slice(1).join(" ") || "";
    if (last.trim() && first.trim()) return { first: first.trim(), last: last.trim() };
  }
  const pieces = s.split(/\s+/);
  if (pieces.length >= 2) {
    const first = pieces[0];
    const last = pieces[pieces.length - 1];
    if (first && last) return { first, last };
  }
  return null;
}

function normalizeHeaderText(val: string): string {
  return String(val || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeTrainingWeekKey(raw: string): string {
  const s = String(raw || "").trim();
  if (!s) return "";
  const match = s.match(/tw[-\s]?(\d{2})/i);
  if (match) return `TW-${match[1]}`;
  return s.toUpperCase();
}

function normalizeAttendanceEventType(raw: string): string {
  const s = String(raw || "").trim().toLowerCase();
  if (!s) return "";
  if (s.includes("llab")) return "llab";
  if (s.includes("mando")) return "mando";
  if (s.includes("secondary")) return "secondary";
  return s;
}

function resolveAttendanceEventId(eventIdRaw: string, trainingWeek: string, eventType: string): string {
  const direct = String(eventIdRaw || "").trim();
  if (direct) return direct;

  const events = Shamrock.listEvents();
  const candidates = events.filter(ev => {
    const status = String((ev as any).event_status || "").toLowerCase();
    if (status !== "published") return false;
    const affects = String((ev as any).affects_attendance || "true").toLowerCase();
    if (affects === "false") return false;
    const twMatch = trainingWeek ? normalizeTrainingWeekKey((ev as any).training_week || "") === trainingWeek : true;
    const typeMatch = eventType ? normalizeAttendanceEventType((ev as any).event_type || (ev as any).attendance_label || (ev as any).event_name || "") === eventType : true;
    return twMatch && typeMatch;
  });

  if (candidates.length === 1) return (candidates[0] as any).event_id;
  if (candidates.length > 1) {
    const exact = candidates.find(ev => normalizeTrainingWeekKey((ev as any).training_week || "") === trainingWeek && normalizeAttendanceEventType((ev as any).event_type || "") === eventType);
    if (exact) return (exact as any).event_id;
    return (candidates[0] as any).event_id;
  }

  throw new Error("No matching published event found for this attendance submission. Ensure Training Week and Event match an existing event.");
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
