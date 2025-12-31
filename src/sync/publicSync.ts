// @ts-nocheck
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

Shamrock.syncAllPublicViews = function (): void {
  const frontend = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId());
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("sync.publicViews", () => {
      Shamrock.rebuildDirectory(frontend);
      Shamrock.rebuildEvents(frontend);
      Shamrock.rebuildAttendance(frontend);
      Shamrock.rebuildExcusals(frontend);
      Shamrock.rebuildAudit(frontend);
      Shamrock.rebuildDataLegend(frontend);
    }, { frontendId: frontend.getId() });
  } else {
    Shamrock.rebuildDirectory(frontend);
    Shamrock.rebuildEvents(frontend);
    Shamrock.rebuildAttendance(frontend);
    Shamrock.rebuildExcusals(frontend);
    Shamrock.rebuildAudit(frontend);
    Shamrock.rebuildDataLegend(frontend);
  }
};

Shamrock.rebuildDirectory = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const { machine: machineHeaders, human: humanHeaders } = Shamrock.PUBLIC_DIRECTORY_HEADERS;
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, Shamrock.PUBLIC_SHEET_NAMES.directory, machineHeaders, humanHeaders);
  const cadets = Shamrock.listCadets();
  const rows = cadets
    .slice()
    .sort((a, b) => {
      const asA = parseAsYear(a.as_year);
      const asB = parseAsYear(b.as_year);
      if (asA !== asB) return asB - asA; // Z -> A (higher AS first)
      const lastA = (a.last_name || "").toLowerCase();
      const lastB = (b.last_name || "").toLowerCase();
      if (lastA < lastB) return -1;
      if (lastA > lastB) return 1;
      return 0;
    })
    .map(cadet => [
      cadet.last_name || "",
      cadet.first_name || "",
      cadet.as_year || "",
      cadet.graduation_year || "",
      cadet.flight || "",
      cadet.squadron || "",
      cadet.university || "",
      cadet.cadet_email || "",
      cadet.phone || "",
      cadet.dorm || "",
      cadet.home_town || "",
      cadet.home_state || "",
      cadet.dob || "",
      cadet.cip_broad || "",
      cadet.cip_code || "",
      cadet.afsc || "",
      cadet.flight_path_status || "",
      cadet.status || "",
      cadet.photo_url || "",
      cadet.notes || "",
      cadet.updated_at || cadet.created_at || "",
      "backend",
    ]);
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.directory", () => writeTable(sheet, machineHeaders, humanHeaders, rows), { rows: rows.length });
  } else {
    writeTable(sheet, machineHeaders, humanHeaders, rows);
  }

  try {
    sheet.showColumns(1, machineHeaders.length);
    [6, 15, 21, 22].forEach(col => sheet.hideColumns(col));
  } catch (err) {
    // If sheet operations fail, skip hiding to avoid breaking rebuild
  }
};

function parseAsYear(asYear: string | undefined): number {
  const match = String(asYear || "").match(/AS(\d+)/i);
  if (match && match[1]) return parseInt(match[1], 10);
  return 0;
}

Shamrock.rebuildEvents = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const { machine: machineHeaders, human: humanHeaders } = Shamrock.PUBLIC_EVENTS_HEADERS;
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, Shamrock.PUBLIC_SHEET_NAMES.events, machineHeaders, humanHeaders);
  const events = Shamrock.listEvents();
  const rows = events.map(ev => [
    ev.event_id || "",
    deriveTerm(ev.training_week, ev.event_date) || "",
    ev.training_week || "",
    ev.event_type || "",
    ev.event_name || "",
    ev.attendance_label || ev.event_name || ev.event_id,
    ev.expected_group || "",
    ev.flight_scope || "",
    ev.event_status || "",
    ev.event_date || "",
    "",
    ev.location || "",
    ev.notes || "",
    ev.created_at || "",
    (ev as any).created_by || "System",
  ]);
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.events", () => writeTable(sheet, machineHeaders, humanHeaders, rows), { rows: rows.length });
  } else {
    writeTable(sheet, machineHeaders, humanHeaders, rows);
  }
};

Shamrock.rebuildAttendance = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const cadets = Shamrock.listCadets();
  const events = Shamrock.listEvents().filter(ev => String((ev as any).affects_attendance).toLowerCase() === "true" || (ev as any).affects_attendance === true);
  const publishedEvents = events.filter(ev => ev.event_status === "Published" || ev.event_status === "Cancelled" || ev.event_status === "Archived");
  const attendance = Shamrock.listAttendance();
  const baseMachine = Shamrock.PUBLIC_ATTENDANCE_BASE_HEADERS.machine;
  const baseHuman = Shamrock.PUBLIC_ATTENDANCE_BASE_HEADERS.human;

  const machineHeaders = baseMachine.concat(publishedEvents.map(ev => ev.event_id));
  const humanHeaders = baseHuman.concat(publishedEvents.map(ev => (ev as any).attendance_label || ev.event_name || ev.event_id));
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, Shamrock.PUBLIC_SHEET_NAMES.attendance, machineHeaders, humanHeaders);

  const attendanceMap = new Map<string, string>();
  attendance.forEach(row => {
    const key = `${row.cadet_email}|${row.event_id}`;
    attendanceMap.set(key, row.attendance_code || "");
  });

  const rows = cadets.map(cadet => {
    const row = [
      cadet.cadet_email || "",
      cadet.status || "",
      cadet.last_name || "",
      cadet.first_name || "",
      cadet.as_year || "",
      cadet.flight || "",
      cadet.squadron || "",
    ];
    for (const ev of publishedEvents) {
      const key = `${cadet.cadet_email}|${ev.event_id}`;
      if (ev.event_status === "Cancelled") {
        row.push("N/A");
      } else {
        row.push(attendanceMap.get(key) || "");
      }
    }
    return row;
  });

  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.attendance", () => writeTable(sheet, machineHeaders, humanHeaders, rows), { rows: rows.length, events: publishedEvents.length });
  } else {
    writeTable(sheet, machineHeaders, humanHeaders, rows);
  }
};

Shamrock.rebuildExcusals = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const { machine: machineHeaders, human: humanHeaders } = Shamrock.PUBLIC_EXCUSALS_HEADERS;
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, Shamrock.PUBLIC_SHEET_NAMES.excusals, machineHeaders, humanHeaders);

  const cadetMap = new Map<string, { last: string; first: string; flight: string; squadron: string }>();
  Shamrock.listCadets().forEach(c => {
    cadetMap.set((c.cadet_email || "").toLowerCase(), {
      last: c.last_name || "",
      first: c.first_name || "",
      flight: c.flight || "",
      squadron: c.squadron || "",
    });
  });

  const rows = Shamrock.listExcusals().map(ex => {
    const cadet = cadetMap.get((ex.cadet_email || "").toLowerCase()) || { last: "", first: "", flight: "", squadron: "" };
    return [
      ex.excusal_id || "",
      ex.event_id || "",
      ex.cadet_email || "",
      cadet.last,
      cadet.first,
      cadet.flight,
      cadet.squadron,
      ex.decision ? ex.decision : "Submitted",
      ex.decision || "",
      ex.decision_by || "",
      ex.decision_timestamp || "",
      ex.attendance_effect || "",
      ex.request_timestamp || "",
      Shamrock.nowIso(),
      "",
    ];
  });

  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.excusals", () => writeTable(sheet, machineHeaders, humanHeaders, rows), { rows: rows.length });
  } else {
    writeTable(sheet, machineHeaders, humanHeaders, rows);
  }
};

Shamrock.rebuildDataLegend = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const backendId = typeof Shamrock.getBackendSpreadsheetIdSafe === "function" ? Shamrock.getBackendSpreadsheetIdSafe() : null;
  if (!backendId || backendId === "SHAMROCK_BACKEND_SPREADSHEET_ID") return;

  let backend: GoogleAppsScript.Spreadsheet.Spreadsheet;
  try {
    backend = SpreadsheetApp.openById(backendId);
  } catch (err) {
    return;
  }

  const source = backend.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.dataLegend);
  if (!source) return;
  const values = source.getDataRange().getValues();
  if (values.length < 2) return;

  const machineHeaders = values[0];
  const humanHeaders = values[1];
  const body = values.slice(2);
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, "Data Legend", machineHeaders, humanHeaders);
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.dataLegend", () => writeTable(sheet, machineHeaders, humanHeaders, body), { rows: body.length });
  } else {
    writeTable(sheet, machineHeaders, humanHeaders, body);
  }
};

Shamrock.rebuildAudit = function (frontend: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(Shamrock.getFrontendSpreadsheetId())): void {
  const backend = SpreadsheetApp.getActive();
  const source = backend.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.audit);
  if (!source) return;
  const lastRow = source.getLastRow();
  if (lastRow < 2) return;
  const values = source.getDataRange().getValues();
  const sheet = Shamrock.ensureSheetWithHeaders(frontend, Shamrock.PUBLIC_SHEET_NAMES.audit, values[0], values[1]);
  const body = values.slice(2);
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuild.audit", () => writeTable(sheet, values[0], values[1], body), { rows: body.length });
  } else {
    writeTable(sheet, values[0], values[1], body);
  }
};

function writeTable(sheet: GoogleAppsScript.Spreadsheet.Sheet, machineHeaders: any[], humanHeaders: any[], rows: any[][]): void {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, machineHeaders.length).setValues([machineHeaders]);
  sheet.getRange(2, 1, 1, humanHeaders.length).setValues([humanHeaders]);
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, machineHeaders.length).setValues(rows);
  }
}

function deriveTerm(trainingWeek: string | undefined, eventDate: string | Date | undefined): string {
  if (trainingWeek && /^(20\d{2}[SF])$/i.test(String(trainingWeek))) return String(trainingWeek);
  if (!eventDate) return "";
  const date = typeof eventDate === "string" ? new Date(eventDate) : eventDate;
  const year = date.getFullYear();
  const month = date.getMonth();
  return month < 6 ? `${year}S` : `${year}F`;
}
