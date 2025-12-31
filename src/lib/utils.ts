// Utility helpers for SHAMROCK backend.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

type HeaderMap = { [key: string]: number };

Shamrock.nowIso = function (): string {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
};

Shamrock.withLock = function <T>(fn: () => T): T {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
};

Shamrock.normalizeHeader = function (header: unknown): string {
  return String(header || "").trim();
};

Shamrock.buildHeaderMap = function (values: any[]): HeaderMap {
  const map: HeaderMap = {};
  values.forEach((v, idx) => {
    const key = Shamrock.normalizeHeader(v);
    if (key) map[key] = idx;
  });
  return map;
};

Shamrock.toObjectRow = function (headers: HeaderMap, row: any[]): Record<string, any> {
  const out: Record<string, any> = {};
  for (const key of Object.keys(headers)) {
    out[key] = row[headers[key]];
  }
  return out;
};

Shamrock.dataStartRow = function (sheet: GoogleAppsScript.Spreadsheet.Sheet, fieldCount: number): number {
  // If row 2 has any content (e.g., human headers), treat data as starting on row 3
  if (sheet.getMaxRows() >= 2) {
    const secondRow = sheet.getRange(2, 1, 1, fieldCount).getValues()[0];
    const hasContent = secondRow.some(cell => String(cell || "").trim().length > 0);
    if (hasContent) return 3;
  }
  return 2;
};

Shamrock.upsertRowByKey = function (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headerRow: number,
  keyColumn: number,
  keyValue: string,
  rowValues: any[],
  dataStartRow?: number,
): number {
  const startRow = dataStartRow || headerRow + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    sheet.getRange(startRow, 1, 1, rowValues.length).setValues([rowValues]);
    return startRow;
  }
  const keyRange = sheet.getRange(startRow, keyColumn, lastRow - startRow + 1, 1);
  const keys = keyRange.getValues().map(r => Shamrock.normalizeHeader(r[0]));
  const idx = keys.findIndex(k => k === Shamrock.normalizeHeader(keyValue));
  const targetRow = idx === -1 ? lastRow + 1 : startRow + idx;
  sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  return targetRow;
};

Shamrock.ensureSheetWithHeaders = function (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  name: string,
  machineHeaders: readonly string[],
  humanHeaders?: readonly string[],
): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  const machine = Array.from(machineHeaders);
  const human = humanHeaders ? Array.from(humanHeaders) : undefined;
  const maxCols = Math.max(machine.length, human?.length || 0);
  if (sheet.getMaxColumns() < maxCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, machine.length).setValues([machine]);
  if (human && human.length) {
    sheet.getRange(2, 1, 1, human.length).setValues([human]);
  }
  return sheet;
};

Shamrock.appendAuditRow = function (
  auditSheet: GoogleAppsScript.Spreadsheet.Sheet,
  entry: Record<typeof Shamrock.AUDIT_FIELDS[number], any>,
): void {
  const row = (Shamrock.AUDIT_FIELDS as string[]).map(key => entry[key]);
  auditSheet.appendRow(row);
};
