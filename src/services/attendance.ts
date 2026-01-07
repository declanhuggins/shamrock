// Attendance matrix builder: projects directory + events + attendance log into backend and frontend matrices.

namespace AttendanceService {
  interface CadetKey {
    last: string;
    first: string;
  }

  interface EventDef {
    name: string; // display_name
    eventId: string;
    eventType: string;
  }

  const BASE_HEADERS = ['last_name', 'first_name', 'as_year', 'flight', 'squadron'];
  const SUMMARY_HEADERS = ['overall_attendance_pct', 'llab_attendance_pct'];
  const CREDIT_CODES = new Set(['P', 'E', 'ES', 'MU', 'MRS']);
  const CREDIT_PATTERNS = ['P*', 'E', 'ES*', 'MU*', 'MRS*'];
  const TOTAL_PATTERNS = ['P*', 'E', 'ES*', 'ER*', 'ED*', 'T*', 'UR*', 'U', 'MU*', 'MRS*'];

  function getIds() {
    return {
      backendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.BACKEND_SHEET_ID) || '',
      frontendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.FRONTEND_SHEET_ID) || '',
    };
  }

  function ensureMatrixSheet(spreadsheetId: string, name: string): GoogleAppsScript.Spreadsheet.Sheet | null {
    if (!spreadsheetId) return null;
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    return sheet;
  }

  function readDirectory(): any[] {
    const { backendId } = getIds();
    if (!backendId) return [];
    const sheet = SheetUtils.getSheet(backendId, 'Directory Backend');
    if (!sheet) return [];
    return SheetUtils.readTable(sheet).rows;
  }

  function readEvents(): EventDef[] {
    const { backendId } = getIds();
    if (!backendId) return [];
    const sheet = SheetUtils.getSheet(backendId, 'Events Backend');
    if (!sheet) return [];
    return SheetUtils.readTable(sheet)
      .rows
      .map((r) => ({
        name: r['display_name'] || r['attendance_column_label'] || r['event_id'] || '',
        eventId: r['event_id'] || r['display_name'] || '',
        eventType: String(r['event_type'] || '').toLowerCase(),
      }))
      .filter((e) => e.name);
  }

  function readAttendanceLog(): any[] {
    const { backendId } = getIds();
    if (!backendId) return [];
    const sheet = SheetUtils.getSheet(backendId, 'Attendance Backend');
    if (!sheet) return [];
    return SheetUtils.readTable(sheet).rows;
  }

  function colToLetter(col: number): string {
    let n = col;
    let s = '';
    while (n > 0) {
      const rem = ((n - 1) % 26) + 1;
      s = String.fromCharCode(64 + rem) + s;
      n = Math.floor((n - rem) / 26);
    }
    return s;
  }

  function applyAttendanceFormulas(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rowsCount: number,
    eventsStartCol: number,
    eventsEndCol: number,
  ) {
    if (!sheet || rowsCount <= 0 || eventsEndCol < eventsStartCol) return;

    const overallCol = BASE_HEADERS.length + 1; // column for overall_attendance_pct (1-indexed)
    const llabCol = BASE_HEADERS.length + 2; // column for llab_attendance_pct
    const startRow = 3;

    const eventsHeaderRange = `${colToLetter(eventsStartCol)}1:${colToLetter(eventsEndCol)}1`;
    const eventsDataRange = `${colToLetter(eventsStartCol)}${startRow}:${colToLetter(eventsEndCol)}`;

    const overallFormula =
      `=ARRAYFORMULA(IF(ROW(${colToLetter(overallCol)}$${startRow}:${colToLetter(overallCol)})<${startRow},"",` +
      `BYROW(${eventsDataRange},LAMBDA(r,` +
      `LET(` +
      `cred,BYCOL(r,LAMBDA(c,IF(SUM(COUNTIF(c,{"${CREDIT_PATTERNS.join('","')}"}))>0,1,0))),` +
      `tot,BYCOL(r,LAMBDA(c,IF(SUM(COUNTIF(c,{"${TOTAL_PATTERNS.join('","')}"}))>0,1,0))),` +
      `num,SUM(cred),` +
      `den,SUM(tot),` +
      `IF(den=0,"",num/den)` +
      `))))`;

    const llabFormula =
      `=ARRAYFORMULA(IF(ROW(${colToLetter(llabCol)}$${startRow}:${colToLetter(llabCol)})<${startRow},"",` +
      `BYROW(${eventsDataRange},LAMBDA(r,` +
      `LET(h,${eventsHeaderRange},` +
      `mask,BYCOL(h,LAMBDA(hd,IF(REGEXMATCH(hd,"(?i)llab"),1,0))),` +
      `cred,BYCOL(r,LAMBDA(c,IF(SUM(COUNTIF(c,{"${CREDIT_PATTERNS.join('","')}"}))>0,1,0))),` +
      `tot,BYCOL(r,LAMBDA(c,IF(SUM(COUNTIF(c,{"${TOTAL_PATTERNS.join('","')}"}))>0,1,0))),` +
      `num,SUM(mask*cred),` +
      `den,SUM(mask*tot),` +
      `IF(den=0,"",num/den)` +
      `))))`;

    // Clear existing values in summary columns and apply formulas
    sheet.getRange(startRow, overallCol, rowsCount, 1).clearContent();
    sheet.getRange(startRow, llabCol, rowsCount, 1).clearContent();
    sheet.getRange(startRow, overallCol).setFormula(overallFormula);
    sheet.getRange(startRow, llabCol).setFormula(llabFormula);
  }

  function normalizeName(part: string): string {
    return String(part || '').trim().toLowerCase();
  }

  function cadetKey(cadet: any): string {
    return `${normalizeName(cadet.last_name)}|${normalizeName(cadet.first_name)}`;
  }

  function parseCadetEntries(cadetField: string): CadetKey[] {
    if (!cadetField) return [];
    return cadetField
      .split(';')
      .map((s) => s.trim())
      .filter(Boolean)
      .map((entry) => {
        // Accept "Last, First" or "Last, First=Code" or "Last, First (AS ...)=Code".
        const [namePart] = entry.split('=');
        const cleaned = namePart.replace(/\(AS[^)]*\)/gi, '').trim();
        const [last, first] = cleaned.split(',').map((p) => p.trim());
        return { last: last || '', first: first || '' };
      })
      .filter((k) => k.last || k.first);
  }

  function buildMatrixRows(directory: any[], events: EventDef[], logRows: any[]) {
    const rows = directory.map((d) => ({
      last_name: d['last_name'] || '',
      first_name: d['first_name'] || '',
      as_year: d['as_year'] || '',
      flight: d['flight'] || '',
      squadron: d['squadron'] || '',
      overall_attendance_pct: '',
      llab_attendance_pct: '',
    }));

    const keyToIndex = new Map<string, number>();
    rows.forEach((r, idx) => keyToIndex.set(cadetKey(r), idx));

    // Initialize event columns with ''
    rows.forEach((r) => {
      events.forEach((ev) => {
        (r as any)[ev.name] = '';
      });
    });

    logRows.forEach((entry) => {
      const evName = entry['event'] || entry['display_name'] || '';
      if (!evName) return;
      const cadets = parseCadetEntries(entry['cadets'] || '');
      cadets.forEach((c) => {
        const idx = keyToIndex.get(`${normalizeName(c.last)}|${normalizeName(c.first)}`);
        if (idx === undefined) return;
        const row = rows[idx] as any;
        if (evName in row) {
          row[evName] = 'P';
        }
      });
    });

    return rows;
  }

  function writeMatrix(sheet: GoogleAppsScript.Spreadsheet.Sheet, events: EventDef[], rows: any[]) {
    const machineHeaders = [...BASE_HEADERS, ...SUMMARY_HEADERS, ...events.map((e) => e.name)];
    const displayHeaders = [
      'Last Name',
      'First Name',
      'AS Year',
      'Flight',
      'Squadron',
      'Overall Attendance %',
      'LLAB Attendance %',
      ...events.map((e) => e.name),
    ];
    sheet.clear();
    if (machineHeaders.length) sheet.getRange(1, 1, 1, machineHeaders.length).setValues([machineHeaders]);
    if (displayHeaders.length) sheet.getRange(2, 1, 1, displayHeaders.length).setValues([displayHeaders]);
    if (rows.length) {
      const data = rows.map((r) => machineHeaders.map((h) => (r as any)[h] ?? ''));
      sheet.getRange(3, 1, data.length, machineHeaders.length).setValues(data);
      const eventsStartCol = BASE_HEADERS.length + SUMMARY_HEADERS.length + 1;
      const eventsEndCol = machineHeaders.length;
      applyAttendanceFormulas(sheet, rows.length, eventsStartCol, eventsEndCol);
    }
  }

  export function rebuildMatrix() {
    const { backendId, frontendId } = getIds();
    const directory = readDirectory();
    const events = readEvents();
    const logRows = readAttendanceLog();
    const matrixRows = buildMatrixRows(directory, events, logRows);

    const backendSheet = ensureMatrixSheet(backendId, 'Attendance Matrix Backend');
    const frontendSheet = frontendId ? SheetUtils.getSheet(frontendId, 'Attendance') : null;

    if (backendSheet) writeMatrix(backendSheet, events, matrixRows);
    if (frontendSheet) writeMatrix(frontendSheet, events, matrixRows);
  }
}