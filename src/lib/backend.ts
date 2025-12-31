// @ts-nocheck
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

function getBackendSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  // Safely resolve the backend spreadsheet ID from globals without throwing when `this` is undefined.
  const idFromConfig = typeof Shamrock.getBackendSpreadsheetIdSafe === "function" ? Shamrock.getBackendSpreadsheetIdSafe() : null;
  const globalObj: any = typeof globalThis !== "undefined" ? (globalThis as any) : (typeof this !== "undefined" ? (this as any) : {});
  const idFromGlobal = globalObj && typeof globalObj.SHAMROCK_BACKEND_ID === "string" ? globalObj.SHAMROCK_BACKEND_ID : globalObj?.SHAMROCK_BACKEND_SPREADSHEET_ID || null;
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore
  const idFromVar = typeof SHAMROCK_BACKEND_ID !== "undefined" ? (SHAMROCK_BACKEND_ID as string) : (typeof SHAMROCK_BACKEND_SPREADSHEET_ID !== "undefined" ? (SHAMROCK_BACKEND_SPREADSHEET_ID as string) : null);
  const idFromShamrock = Shamrock && typeof Shamrock.BACKEND_ID === "string" ? Shamrock.BACKEND_ID : null;
  const backendId = idFromConfig || idFromGlobal || idFromVar || idFromShamrock;

  if (!backendId) return SpreadsheetApp.getActive();
  try {
    return SpreadsheetApp.openById(backendId);
  } catch (err) {
    return SpreadsheetApp.getActive();
  }
}

type CadetRecord = Record<typeof Shamrock.CADET_FIELDS[number], any>;
type EventRecord = Record<typeof Shamrock.EVENT_FIELDS[number], any>;
type AttendanceRecord = Record<typeof Shamrock.ATTENDANCE_FIELDS[number], any>;
type ExcusalRecord = Record<typeof Shamrock.EXCUSAL_FIELDS[number], any>;
type AdminActionRecord = Record<typeof Shamrock.ADMIN_ACTION_FIELDS[number], any>;

Shamrock.ensureBackendSheets = function (): void {
  const ss = getBackendSpreadsheet();
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.cadets, Array.from(Shamrock.CADET_FIELDS));
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.events, Array.from(Shamrock.EVENT_FIELDS));
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.attendance, Array.from(Shamrock.ATTENDANCE_FIELDS));
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.excusals, Array.from(Shamrock.EXCUSAL_FIELDS));
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.adminActions, Array.from(Shamrock.ADMIN_ACTION_FIELDS));
  Shamrock.ensureSheetWithHeaders(ss, Shamrock.BACKEND_SHEET_NAMES.audit, Array.from(Shamrock.AUDIT_FIELDS), Array.from(Shamrock.AUDIT_HEADERS_HUMAN));

  ensureDataLegendSheet(ss);
};

function ensureDataLegendSheet(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const columns = getDataLegendColumns();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.dataLegend) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.dataLegend);
  const machineHeaders = columns.map(c => c.name);
  const humanHeaders = columns.map(c => c.label);
  sheet.getRange(1, 1, 1, machineHeaders.length).setValues([machineHeaders]);
  sheet.getRange(2, 1, 1, humanHeaders.length).setValues([humanHeaders]);
  const maxLen = Math.max(...columns.map(c => c.values.length));
  const body: any[][] = [];
  for (let r = 0; r < maxLen; r++) {
    body[r] = columns.map(c => c.values[r] || "");
  }
  if (maxLen > 0) {
    sheet.getRange(3, 1, maxLen, machineHeaders.length).setValues(body);
  } else {
    const maxRows = sheet.getMaxRows();
    if (maxRows > 2) sheet.deleteRows(3, maxRows - 2);
  }
  const maxCols = sheet.getMaxColumns();
  if (maxCols > machineHeaders.length) {
    sheet.deleteColumns(machineHeaders.length + 1, maxCols - machineHeaders.length);
  }
  sheet.hideRows(1); // hide machine headers
}

type DataLegendColumn = { name: string; label: string; values: string[] };

function getDataLegendColumns(): DataLegendColumn[] {
  const opts = Shamrock.DATA_LEGEND_OPTIONS;
  return [
    { name: "as_year_options", label: "AS Year Options", values: Array.from(opts.as_year_options) },
    { name: "flight_options", label: "Flight Options", values: Array.from(opts.flight_options) },
    { name: "squadron_options", label: "Squadron Options", values: Array.from(opts.squadron_options) },
    { name: "university_options", label: "University Options", values: Array.from(opts.university_options) },
    { name: "dorm_options", label: "Dorm Options", values: Array.from(opts.dorm_options) },
    { name: "home_state_options", label: "Home State Options", values: Array.from(opts.home_state_options) },
    { name: "afsc_options", label: "AFSC Options", values: Array.from(opts.afsc_options) },
    { name: "flight_path_status_options", label: "Flight Path Status Options", values: Array.from(opts.flight_path_status_options) },
    { name: "status_options", label: "Status Options", values: Array.from(opts.status_options) },
    { name: "cip_broad_options", label: "CIP Broad Options", values: Array.from(opts.cip_broad_options) },
  ];
}

Shamrock.listCadets = function (): CadetRecord[] {
  const sheet = getBackendSpreadsheet().getSheetByName(Shamrock.BACKEND_SHEET_NAMES.cadets);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.CADET_FIELDS.length);
  if (lastRow < dataStart) return [];
  const values = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, Shamrock.CADET_FIELDS.length).getValues();
  return values.map(row => rowToRecord<CadetRecord>(Array.from(Shamrock.CADET_FIELDS), row));
};

Shamrock.listEvents = function (): EventRecord[] {
  const sheet = getBackendSpreadsheet().getSheetByName(Shamrock.BACKEND_SHEET_NAMES.events);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.EVENT_FIELDS.length);
  if (lastRow < dataStart) return [];
  const values = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, Shamrock.EVENT_FIELDS.length).getValues();
  return values.map(row => rowToRecord<EventRecord>(Array.from(Shamrock.EVENT_FIELDS), row));
};

Shamrock.listAttendance = function (): AttendanceRecord[] {
  const sheet = getBackendSpreadsheet().getSheetByName(Shamrock.BACKEND_SHEET_NAMES.attendance);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.ATTENDANCE_FIELDS.length);
  if (lastRow < dataStart) return [];
  const values = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, Shamrock.ATTENDANCE_FIELDS.length).getValues();
  return values.map(row => rowToRecord<AttendanceRecord>(Array.from(Shamrock.ATTENDANCE_FIELDS), row));
};

Shamrock.listExcusals = function (): ExcusalRecord[] {
  const sheet = getBackendSpreadsheet().getSheetByName(Shamrock.BACKEND_SHEET_NAMES.excusals);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.EXCUSAL_FIELDS.length);
  if (lastRow < dataStart) return [];
  const values = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, Shamrock.EXCUSAL_FIELDS.length).getValues();
  return values.map(row => rowToRecord<ExcusalRecord>(Array.from(Shamrock.EXCUSAL_FIELDS), row));
};

Shamrock.upsertCadet = function (record: Partial<CadetRecord> & { cadet_email: string }): number {
  const ss = getBackendSpreadsheet();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.cadets) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.cadets);
  const row = Array.from(Shamrock.CADET_FIELDS).map(field => {
    if (field === "updated_at") return Shamrock.nowIso();
    if (field === "created_at") return (record as any).created_at || Shamrock.nowIso();
    return (record as any)[field] || "";
  });
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.CADET_FIELDS.length);
  return Shamrock.upsertRowByKey(sheet, 1, Shamrock.CADET_FIELDS.indexOf("cadet_email") + 1, record.cadet_email, row, dataStart);
};

Shamrock.upsertEvent = function (record: Partial<EventRecord> & { event_id: string }): number {
  const ss = getBackendSpreadsheet();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.events) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.events);
  const row = Array.from(Shamrock.EVENT_FIELDS).map(field => {
    if (field === "updated_at") return Shamrock.nowIso();
    if (field === "created_at") return (record as any).created_at || Shamrock.nowIso();
    return (record as any)[field] || "";
  });
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.EVENT_FIELDS.length);
  return Shamrock.upsertRowByKey(sheet, 1, Shamrock.EVENT_FIELDS.indexOf("event_id") + 1, record.event_id, row, dataStart);
};

Shamrock.setAttendance = function (record: AttendanceRecord): number {
  const ss = getBackendSpreadsheet();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.attendance) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.attendance);
  const lastRow = sheet.getLastRow();
  const headerRow = 1;
  const dataStart = Shamrock.dataStartRow(sheet, Shamrock.ATTENDANCE_FIELDS.length);
  const targetRow = findAttendanceRow(sheet, headerRow, dataStart, record.cadet_email, record.event_id);
  const rowValues = Array.from(Shamrock.ATTENDANCE_FIELDS).map(field => {
    if (field === "updated_at") return Shamrock.nowIso();
    return (record as any)[field] || "";
  });
  const destRow = targetRow || Math.max(lastRow + 1, dataStart);
  sheet.getRange(destRow, 1, 1, rowValues.length).setValues([rowValues]);
  return destRow;
};

function findAttendanceRow(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headerRow: number,
  dataStartRow: number,
  cadetEmail: string,
  eventId: string,
): number | null {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return null;
  const data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, Shamrock.ATTENDANCE_FIELDS.length).getValues();
  const emailIdx = Shamrock.ATTENDANCE_FIELDS.indexOf("cadet_email");
  const eventIdx = Shamrock.ATTENDANCE_FIELDS.indexOf("event_id");
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (Shamrock.normalizeHeader(row[emailIdx]) === Shamrock.normalizeHeader(cadetEmail) && Shamrock.normalizeHeader(row[eventIdx]) === Shamrock.normalizeHeader(eventId)) {
      return dataStartRow + i;
    }
  }
  return null;
}

Shamrock.appendExcusal = function (record: ExcusalRecord): number {
  const ss = getBackendSpreadsheet();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.excusals) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.excusals);
  const row = Array.from(Shamrock.EXCUSAL_FIELDS).map(field => (record as any)[field] || "");
  sheet.appendRow(row);
  return sheet.getLastRow();
};

Shamrock.appendAdminAction = function (record: AdminActionRecord): number {
  const ss = getBackendSpreadsheet();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.adminActions) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.adminActions);
  const row = Array.from(Shamrock.ADMIN_ACTION_FIELDS).map(field => (record as any)[field] || "");
  sheet.appendRow(row);
  return sheet.getLastRow();
};

function rowToRecord<T extends Record<string, any>>(headers: string[], row: any[]): T {
  const obj: Record<string, any> = {};
  headers.forEach((h, idx) => {
    obj[h] = row[idx];
  });
  return obj as T;
}
