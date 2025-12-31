Do not directly use code from these snippets, but rather use them as reference of what the old, slow system worked as, may be helpful for logic or explanations of how the system should act.

```SyncCadets.gs
function onFormSubmit(e) {
  const lock = LockService.getDocumentLock();
  lock.tryLock(30000); // avoid races on bursts of submissions
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Form Responses') return;

    // Clean the newly appended row in place
    const row = e.range.getRow();
    const lastCol = sh.getLastColumn();
    const rng = sh.getRange(row, 1, 1, lastCol);
    const values = rng.getValues();
    const cleaned = values.map(r => r.map(normalizeText));
    rng.setValues(cleaned);

    // Then run existing pipeline (which sorts + syncs)
    syncNewCadets();
  } finally {
    lock.releaseLock();
  }
}

function syncNewCadets() {
  const ss = SpreadsheetApp.getActive();
  const formSheet = ss.getSheetByName('Form Responses');
  const dirSheet = ss.getSheetByName('Cadet Directory');

  formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).sort({column: 1, ascending: false});

  const formData = formSheet.getRange(2, 4, formSheet.getLastRow() - 1, 2).getValues(); // D (Last), E (First)
  const dirData = dirSheet.getRange(5, 3, dirSheet.getLastRow() - 4, 2).getValues();    // C (Last), D (First)

  const existingNames = new Set(dirData.map(([last, first]) =>
    `${last?.toString().trim().toLowerCase()}|${first?.toString().trim().toLowerCase()}`
  ));

  let newRows = [];

  formData.forEach(([lastName, firstName], i) => {
    const key = `${lastName?.toString().trim().toLowerCase()}|${firstName?.toString().trim().toLowerCase()}`;
    if (!existingNames.has(key)) {
      const fullRow = formSheet.getRange(i + 2, 1, 1, formSheet.getLastColumn()).getValues()[0];
      const [timestamp, year] = fullRow;

      const row = new Array(18).fill(""); // Columns B to S
      row[0] = year;      // Column B: Year
      row[1] = lastName;  // Column C: Last Name
      row[2] = firstName; // Column D: First Name
      // Columns P (14), Q (15), R (16) = Photo, Squadron, Flight → left as blank

      newRows.push(row);
    }
  });

  if (newRows.length > 0) {
    const insertStartRow = dirSheet.getLastRow() + 1;
    const startCol = 2; // Column B

    // Insert core values: B to S (18 columns)
    dirSheet.getRange(insertStartRow, startCol, newRows.length, 18).setValues(newRows);

    // Copy formulas from the row above for columns E to O (columns 5–15 relative to B = cols 6–16 absolute)
    const lastRowWithFormulas = insertStartRow - 1;
    const formulaRange = dirSheet.getRange(lastRowWithFormulas, startCol + 3, 1, 11); // E–O
    const formulas = formulaRange.getFormulasR1C1()[0];

    for (let i = 0; i < newRows.length; i++) {
      const targetRange = dirSheet.getRange(insertStartRow + i, startCol + 3, 1, 11); // E–O
      targetRange.setFormulasR1C1([formulas]);
    }

    // Copy DOB formula from column S (index 18 relative to B = column 19 absolute)
    const dobFormula = dirSheet.getRange(lastRowWithFormulas, startCol + 17).getFormulaR1C1();
    for (let i = 0; i < newRows.length; i++) {
      dirSheet.getRange(insertStartRow + i, startCol + 17).setFormulaR1C1(dobFormula);
    }

    Logger.log(`${newRows.length} new cadet(s) added.`);

    // Sort by: Year (B, desc), then Last Name (C, asc), then First Name (D, asc)
    const totalRows = dirSheet.getLastRow();
    dirSheet.getRange(5, 2, totalRows - 4, 18).sort([
      { column: 2, ascending: false }, // B: Year, Z→A
      { column: 3, ascending: true },  // C: Last Name, A→Z
      { column: 4, ascending: true }   // D: First Name, A→Z
    ]);
  } else {
    Logger.log("No new cadets to add.");
  }
  syncToAttendanceTracker();
}
```

```HelperFunctions.gs
// "Dev" menu for quick testing
function onOpen() {
  // Active user email can be empty in some domains; normalize to lower-case
  const userEmail = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();

  // Keep menu private to these users (case-insensitive)
  const allowed = new Set([
    'dhuggin2@nd.edu',
    'nspecht@nd.edu',
    'kander44@nd.edu',
    'atriplet@nd.edu',
    'nallen3@nd.edu'
  ].map(s => s.toLowerCase()));

  if (!userEmail || !allowed.has(userEmail)) {
    // Optional: quiet exit for non-allowed users
    return;
  }

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Dev')
    .addItem('Test submit response', 'testOnFormSubmitLast')
    .addItem('Clean all form responses', 'cleanAllFormResponses')
    .addItem('Sync to attendance tracker (run now)', 'syncToAttendanceTracker')
    .addSeparator()
    .addItem('Install daily trigger (00:00)', 'installDailySyncTrigger')
    .addItem('Remove daily trigger(s)', 'removeDailySyncTriggers')
    .addToUi();
}

/**
 * Install (or re-install) a single daily trigger at 00:00 for syncToAttendanceTracker.
 * Run this once, or call from a menu.
 */
function installDailySyncTrigger() {
  const FN = 'syncToAttendanceTracker';

  // 1) Remove any existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FN)
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 2) Create the daily 00:00 trigger (script timezone)
  ScriptApp.newTrigger(FN)
    .timeBased()
    .everyDays(1)
    .atHour(0)        // 00:00 local to the script's timezone
    .nearMinute(0)    // as close to :00 as Apps Script allows
    .create();

  Logger.log('Installed daily trigger for %s at ~00:00.', FN);
}

/**
 * Optional: quick remover if you need to clear triggers for this function.
 */
function removeDailySyncTriggers() {
  const FN = 'syncToAttendanceTracker';
  const toDelete = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FN);
  toDelete.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Removed %s trigger(s) for %s.', toDelete.length, FN);
}

// Test helper: simulate onFormSubmit for the last row in 'Form Responses'
function testOnFormSubmitLast() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Form Responses');
  const row = sh.getLastRow();
  const e = { range: sh.getRange(row, 1, 1, sh.getLastColumn()) };
  onFormSubmit(e);
}

function cleanAllFormResponses() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Form Responses');
  if (!sh) throw new Error('Sheet "Form Responses" not found');

  const firstDataRow = 2; // keep row 1 as header
  const lastRow = sh.getLastRow();
  if (lastRow < firstDataRow) {
    Logger.log('No data rows to clean.');
    return;
  }

  const lastCol = sh.getLastColumn();
  const totalRows = lastRow - firstDataRow + 1;

  const lock = LockService.getDocumentLock();
  lock.tryLock(30000);
  try {
    const BATCH = 400; // adjust if needed
    let cleanedCount = 0;

    for (let start = firstDataRow; start <= lastRow; start += BATCH) {
      const num = Math.min(BATCH, (lastRow - start + 1));
      const range = sh.getRange(start, 1, num, lastCol);
      const values = range.getValues();

      const cleaned = values.map(row =>
        row.map(cell => (typeof cell === 'string' ? normalizeText(cell) : cell))
      );

      range.setValues(cleaned);
      cleanedCount += num;
      SpreadsheetApp.flush();
      // Optional: brief pause to be polite with large sheets
      Utilities.sleep(50);
    }

    Logger.log(`Cleaned ${cleanedCount} rows in "Form Responses".`);
  } finally {
    lock.releaseLock();
  }
}

function normalizeText(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return v; // keep dates/numbers/timestamps intact
  // Convert non-breaking and odd Unicode spaces to normal spaces
  let s = v.replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
  // Collapse runs of whitespace and trim ends
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function getSpreadsheetByIdOrUrl(idOrUrl) {
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('Invalid Spreadsheet ID/URL: ' + idOrUrl);
  const id = m[0];
  // permission check
  DriveApp.getFileById(id);
  return SpreadsheetApp.openById(id);
}

/* ensure sheet has at least N rows/cols */
function ensureRows_(sh, minRows) {
  const mr = sh.getMaxRows();
  if (mr < minRows) sh.insertRowsAfter(mr, minRows - mr);
}

function ensureCols_(sh, minCols) {
  const mc = sh.getMaxColumns();
  if (mc < minCols) sh.insertColumnsAfter(mc, minCols - mc);
}

function ensureRowsWithFormat_(sh, minRows, formatTemplateRow) {
  const mr = sh.getMaxRows();
  if (mr < minRows) {
    const rowsToAdd = minRows - mr;
    sh.insertRowsAfter(mr, rowsToAdd);

    // Copy format from template row into each new row
    const templateRange = sh.getRange(formatTemplateRow, 1, 1, sh.getMaxColumns());
    for (let r = mr + 1; r <= minRows; r++) {
      templateRange.copyTo(sh.getRange(r, 1, 1, sh.getMaxColumns()), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }
}
```

```SyncToAttendanceTracker.gs
function syncToAttendanceTracker() {
  const SOURCE_SHEET_NAME = 'Cadet Directory';
  const TARGET_SPREADSHEET_ID_OR_URL = 'https://docs.google.com/spreadsheets/d/1MaPFLz5N8ngR399IrJHsqDMDr4p1mxO8l6OWcg4rKYU';
  const TARGET_SHEET_NAME = 'Attendance';

  // Source (header row 4; data row 5+). Mirror B–E and R -> target A:E (5 cols).
  const SRC_HEADER_ROW = 4, SRC_FIRST_DATA_ROW = 5;
  const SRC_BLOCKS = [{ col: 2, width: 4 }, { col: 18, width: 1 }]; // B–E, R

  // Target bands
  const LEFT_START_COL = 1, LEFT_COLS = 5; // A:E
  const RIGHT_START_COL = 8;               // H..end
  const RED = '#f8d7da';                   // removed highlight

  // Whitelist of fills to keep with a cadet (your “Light blue 3” + variants)
  const PRESERVE_COLORS = new Set([
    '#cfe2f3', // Light blue 3 (standard palette)
    '#c9daf8', // variant often seen in some themes
    '#cfe2ff',  // very light blue variant
    '#D9D9D9', // light gray
    '#5884E1' // MRS blue
  ].map(c => c.toLowerCase()));

  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  const src = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) throw new Error('Source sheet not found: ' + SOURCE_SHEET_NAME);

  const targetSS = getSpreadsheetByIdOrUrl(TARGET_SPREADSHEET_ID_OR_URL);
  const dst = targetSS.getSheetByName(TARGET_SHEET_NAME) || targetSS.insertSheet(TARGET_SHEET_NAME);

  // Build A:E headers from source (B–E + R)
  const headerParts = SRC_BLOCKS.map(b => src.getRange(SRC_HEADER_ROW, b.col, 1, b.width).getValues()[0]);
  const rosterHeader = headerParts.flat().map(v => String(v || '').trim());
  if (rosterHeader.length !== LEFT_COLS) throw new Error('Expected 5 roster header columns for A:E');

  // Write ONLY A:E headers
  ensureRowsWithFormat_(dst, 1, 1);
  ensureCols_(dst, Math.max(dst.getLastColumn(), LEFT_START_COL + LEFT_COLS - 1));
  dst.getRange(1, LEFT_START_COL, 1, LEFT_COLS).setValues([rosterHeader]);

  // Key selection (prefer Email; fallback Last|First)
  const upperHdr = rosterHeader.map(h => h.toUpperCase());
  const idxEmail = (() => {
    let i = upperHdr.indexOf('UNIVERSITY EMAIL ADDRESS');
    if (i === -1) i = upperHdr.indexOf('EMAIL');
    return i;
  })();
  const idxLast  = upperHdr.indexOf('LAST NAME');
  const idxFirst = upperHdr.indexOf('FIRST NAME');

  const makeKey = (leftRow) => {
    const email = (idxEmail >= 0 ? String(leftRow[idxEmail] || '').trim() : '');
    if (email) return email.toUpperCase();
    const last  = (idxLast  >= 0 ? String(leftRow[idxLast]  || '').trim() : '');
    const first = (idxFirst >= 0 ? String(leftRow[idxFirst] || '').trim() : '');
    if (last || first) return (last + '|' + first).toUpperCase();
    return '';
  };

  // Read source left rows (A:E built from B–E + R)
  const lastSrcRow = src.getLastRow();
  const srcLeftRows = [];
  if (lastSrcRow >= SRC_FIRST_DATA_ROW) {
    const bodyParts = SRC_BLOCKS.map(b =>
      src.getRange(SRC_FIRST_DATA_ROW, b.col, lastSrcRow - (SRC_FIRST_DATA_ROW - 1), b.width).getValues()
    );
    const numRows = bodyParts[0].length;
    for (let r = 0; r < numRows; r++) {
      const left = bodyParts[0][r];
      const extra = bodyParts.slice(1).map(block => block[r]).flat();
      const row = left.concat(extra).map(v => (v == null ? '' : String(v).trim()));
      if (row.some(x => x !== '')) srcLeftRows.push(row);
    }
  }

  // Source maps
  const srcKeys = [];
  const srcLeftMap = new Map();
  for (const L of srcLeftRows) {
    const k = makeKey(L);
    if (!k) continue;
    srcKeys.push(k);
    srcLeftMap.set(k, L);
  }
  const srcKeySet = new Set(srcKeys);

  // Determine H..end width
  const dstLastCol = Math.max(dst.getLastColumn(), RIGHT_START_COL - 1);
  const rightCols = Math.max(0, dstLastCol - RIGHT_START_COL + 1);

  // Read existing target A:E and H..end (preserve right-side, detect removed)
  const dstLastRow = dst.getLastRow();
  const existingLeft = (dstLastRow >= 2)
    ? dst.getRange(2, LEFT_START_COL, dstLastRow - 1, LEFT_COLS).getValues()
    : [];
  const existingRight = (rightCols > 0 && dstLastRow >= 2)
    ? dst.getRange(2, RIGHT_START_COL, dstLastRow - 1, rightCols).getValues()
    : [];

  // Also capture existing explicit backgrounds to preserve per cadet (like blue)
  const lastColAll = dst.getLastColumn();
  const existingBg = (dstLastRow >= 2)
    ? dst.getRange(2, 1, dstLastRow - 1, lastColAll).getBackgrounds()
    : [];

  const existingKeys = [];
  const existingLeftMap = new Map();
  const existingRightMap = new Map();
  const existingPreserveBgMap = new Map(); // key -> [color|null] per column

  for (let i = 0; i < existingLeft.length; i++) {
    const L = existingLeft[i].map(v => String(v || '').trim());
    const k = makeKey(L);
    if (!k) continue;
    existingKeys.push(k);
    existingLeftMap.set(k, L);
    if (rightCols > 0) existingRightMap.set(k, existingRight[i] || []);

    const rowBg = existingBg[i] || [];
    const keepBg = rowBg.map(color => {
      const c = (color || '').toLowerCase();
      return PRESERVE_COLORS.has(c) ? color : null; // keep only whitelisted colors
    });
    existingPreserveBgMap.set(k, keepBg);
  }

  // Removed = in existing but not in source (keep original order)
  const removedKeys = existingKeys.filter(k => !srcKeySet.has(k));

  // Final order = actives (source order) + removed (existing order)
  const finalKeys = srcKeys.concat(removedKeys);

  // Build final bodies
  const finalLeftBody  = [];
  const finalRightBody = [];
  for (const k of finalKeys) {
    const L = srcLeftMap.get(k) || existingLeftMap.get(k) || Array(LEFT_COLS).fill('');
    finalLeftBody.push(L);
    if (rightCols > 0) {
      const R = existingRightMap.get(k) || Array(rightCols).fill('');
      finalRightBody.push(R);
    }
  }

  const bodyRows = finalLeftBody.length;
  const totalRows = 1 + bodyRows;

  // Ensure capacity and clear only what we rewrite
  ensureRowsWithFormat_(dst, totalRows, 2);
  ensureCols_(dst, Math.max(dst.getLastColumn(), RIGHT_START_COL + Math.max(0, rightCols) - 1));
  dst.getRange(1, LEFT_START_COL, totalRows, LEFT_COLS).clearContent(); // A:E
  if (rightCols > 0) {
    dst.getRange(2, RIGHT_START_COL, Math.max(bodyRows, 0), rightCols).clearContent(); // H..end body
  }

  // Write header (A:E only)
  dst.getRange(1, LEFT_START_COL, 1, LEFT_COLS).setValues([rosterHeader]);

  // Write body (A:E and H..end)
  if (bodyRows > 0) {
    dst.getRange(2, LEFT_START_COL, bodyRows, LEFT_COLS).setValues(finalLeftBody);
    if (rightCols > 0) {
      dst.getRange(2, RIGHT_START_COL, bodyRows, rightCols).setValues(finalRightBody);
    }
  }

  // Highlight removed rows (whole row)
  const activeCount = srcKeys.length;
  if (removedKeys.length > 0) {
    const removedStart = 2 + activeCount;
    dst.getRange(removedStart, 1, removedKeys.length, dst.getLastColumn())
       .setBackground(RED);
  }

  // --- CLEAN COLORS IN ACTIVE BLOCK, THEN REAPPLY PER-CADET ---

  if (activeCount > 0) {
    const activeRange = dst.getRange(2, 1, activeCount, dst.getLastColumn());
    const bg = activeRange.getBackgrounds();

    // 1) Clear only red from active rows
    for (let r = 0; r < bg.length; r++) {
      for (let c = 0; c < bg[r].length; c++) {
        const col = (bg[r][c] || '').toLowerCase();
        if (col === RED) {
          activeRange.getCell(r + 1, c + 1).setBackground(null);
        }
      }
    }

    // 2) Clear any preserved colors (e.g., blue) from active rows
    const bg2 = activeRange.getBackgrounds();
    for (let r = 0; r < bg2.length; r++) {
      for (let c = 0; c < bg2[r].length; c++) {
        const col = (bg2[r][c] || '').toLowerCase();
        if (PRESERVE_COLORS.has(col)) {
          activeRange.getCell(r + 1, c + 1).setBackground(null);
        }
      }
    }

    // 3) Reapply preserved colors to the cadet now in each row
    const lastColNow = dst.getLastColumn();
    for (let r = 0; r < activeCount; r++) {
      const key = finalKeys[r]; // r=0 -> row 2
      const keepColors = existingPreserveBgMap.get(key);
      if (!keepColors) continue;
      for (let c = 0; c < Math.min(keepColors.length, lastColNow); c++) {
        const colColor = keepColors[c];
        if (colColor) {
          dst.getRange(2 + r, 1 + c).setBackground(colColor);
        }
      }
    }
  }
}
```

```AttendanceSubmission.gs
/***** CONFIG *****/
const RESPONSES_SHEET_NAME = 'Attendance Form'; // Form responses tab
const ROSTER_SHEET_NAME    = 'Attendance';      // Master roster/output tab

// Roster columns (1-based)
const COL_YEAR   = 1; // A
const COL_LAST   = 2; // B
const COL_FIRST  = 3; // C
const COL_EMAIL  = 4; // D
const COL_FLIGHT = 5; // E

// Attendance headers start at column H
const ATTENDANCE_HEADERS_START_COL = 8; // H

// Known flights and AS groups (as they appear in Form headers)
const FLIGHT_NAMES = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot'];
const AS_GROUP_TITLES = ['AS 100/150', 'AS 200/250', 'AS 300', 'AS 400'];

// Treat these as "Cross-Town" (case/space/punct insensitive)
const CROSSTOWN_ALIASES = new Set(['cross-town', 'crosstown', 'cross town']);
function isCrossTownFlight(s) { return CROSSTOWN_ALIASES.has(norm(s)); }

// Ignore these flights (case-insensitive)
const IGNORE_FLIGHTS = new Set(['abroad']);

// Privileged users who can parse ANY flight (and “All Flights”)
const SUPER_USERS = new Set([
  'dhuggin2@nd.edu',
  'nspecht@nd.edu',
  'kander44@nd.edu',
  'atriplet@nd.edu'
]);

// Map each flight to its lead’s email (who may parse ONLY their flight)
const FLIGHT_LEADS = {
  'Alpha': 'msavidge@nd.edu',
  'Bravo': 'njacob@nd.edu',
  'Charlie': 'smouran2@nd.edu',
  'Delta': 'cglenn2@nd.edu',
  'Echo': 'bwhitela@nd.edu',
  'Foxtrot': 'sleavitt@nd.edu',
};

// Attendance Form exact headers
const ATT_HDR_TW     = 'Training Week (Format as TW-00):';
const ATT_HDR_EVENT  = 'Event:';
const ATT_HDR_FLIGHT = 'Flight:';

/***** MENU *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
    .addItem('Parse submissions (prompt)…', 'uiPromptAndParse')
    .addSeparator()
    .addItem('Clean all form responses', 'cleanAllFormResponses')
    .addItem('Process last response (dev)', 'processLastResponseManually')
    .addToUi();
}

/***** INSTALL TRIGGER (AUTO ON SUBMIT) *****/
function installOnFormSubmit() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

/***** AUTO: runs on every form submission *****/
function onFormSubmit(e) {
  try {
    // Clean responses first
    cleanResponsesSheet_(RESPONSES_SHEET_NAME);

    const headers = getHeadersFromSheet(RESPONSES_SHEET_NAME);

    // Safely resolve the submitted row
    let row = null;
    if (e && e.values && Array.isArray(e.values)) {
      row = e.values;
    } else if (e && e.namedValues) {
      row = mapNamedValuesToRow(headers, e.namedValues);
    } else {
      // Fallback: last row from the responses sheet (covers manual runs)
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(RESPONSES_SHEET_NAME);
      const lastRow = sh.getLastRow();
      if (lastRow < 2) throw new Error('No response rows available.');
      row = sh.getRange(lastRow, 1, 1, sh.getLastColumn()).getValues()[0];
      Logger.log('[auto] No event payload; using last row fallback.');
    }

    const result = handleSingleSubmission(headers, row);
    Logger.log(`[auto] Updated "P": ${result.updated}, Missing: ${result.missing}, Notes: ${result.notes}`);
  } catch (err) {
    Logger.log('onFormSubmit error: ' + (err && err.stack || err));
    throw err;
  }
}

/***** UI FLOW (MANUAL PARSE) *****/
function uiPromptAndParse() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  // Choose Flight
  const flightChoices = [...FLIGHT_NAMES];
  if (SUPER_USERS.has(userEmail)) flightChoices.unshift('All Flights');

  const flightResp = ui.prompt('Select Flight', `Enter one of: ${flightChoices.join(', ')}`, ui.ButtonSet.OK_CANCEL);
  if (flightResp.getSelectedButton() !== ui.Button.OK) return;
  const flightInput = String(flightResp.getResponseText() || '').trim();

  // Permission check
  if (!SUPER_USERS.has(userEmail)) {
    if (!FLIGHT_NAMES.includes(flightInput)) {
      ui.alert('Permission denied', 'Only super users may select "All Flights".', ui.ButtonSet.OK);
      return;
    }
    const leadEmail = FLIGHT_LEADS[flightInput];
    if (!leadEmail || leadEmail.toLowerCase() !== (userEmail || '').toLowerCase()) {
      ui.alert('Permission denied', `You are not the lead for ${flightInput} Flight.`, ui.ButtonSet.OK);
      return;
    }
  } else {
    if (!(flightInput === 'All Flights' || FLIGHT_NAMES.includes(flightInput))) {
      ui.alert('Invalid flight', `Please enter one of: ${flightChoices.join(', ')}`, ui.ButtonSet.OK);
      return;
    }
  }

  // Training Week (TW-01..TW-15)
  const twResp = ui.prompt('Training Week', 'Enter like "TW-01", "TW-02", … "TW-15".', ui.ButtonSet.OK_CANCEL);
  if (twResp.getSelectedButton() !== ui.Button.OK) return;
  const weekKey = String(twResp.getResponseText() || '').trim();
  if (!/^TW-\d{2}$/.test(weekKey)) {
    ui.alert('Invalid Training Week', 'Use format "TW-01", "TW-02", … "TW-15".', ui.ButtonSet.OK);
    return;
  }

  // Event
  const evResp = ui.prompt('Event', 'Enter one of: Mando, LLAB, Secondary', ui.ButtonSet.OK_CANCEL);
  if (evResp.getSelectedButton() !== ui.Button.OK) return;
  const eventNorm = norm(evResp.getResponseText());

  if (!['mando', 'llab', 'secondary'].includes(eventNorm)) {
    ui.alert('Invalid Event', 'Enter exactly: Mando, LLAB, or Secondary', ui.ButtonSet.OK);
    return;
  }

  const result = parseSubmissionsBatch({ flightSelection: flightInput, weekKey, eventNorm, userEmail });
  ui.alert(
    'Parse Complete',
    `Matched: ${result.matched}\nUpdated "P": ${result.updated}\nMissing on roster: ${result.missing}\nNotes: ${result.notes}`,
    ui.ButtonSet.OK
  );
}

/***** SINGLE ROW PROCESSOR (used by auto on submit) *****/
function handleSingleSubmission(headers, row) {
  const idx = indexHeaders(headers);

  // Read exact headers you provided
  const weekRaw   = String(readCell(idx, row, ATT_HDR_TW) || '').trim();
  const eventRaw  = String(readCell(idx, row, ATT_HDR_EVENT) || '').trim();
  const flightRaw = String(readCell(idx, row, ATT_HDR_FLIGHT) || '').trim();

  // Normalize for matching columns & branching
  const weekKey   = normalizeTWKeyAttendance_(weekRaw);
  const eventName = normalizeEventTitle_(eventRaw);  // "Secondary" | "Mando" | "LLAB"
  const eventNorm = norm(eventName);                 // "secondary" | "mando" | "llab"

  if (!weekKey || !eventName) {
    return { matched: 0, updated: 0, missing: 0, notes: 'Missing week or event' };
  }

  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!roster) throw new Error(`Roster sheet "${ROSTER_SHEET_NAME}" not found`);

  // Roster map
  const rosterLastRow = roster.getLastRow();
  const rosterLastCol = roster.getLastColumn();
  const rosterData = roster.getRange(1, 1, rosterLastRow, rosterLastCol).getValues();
  const rosterHeader = rosterData[0];
  const nameToRow = new Map();
  for (let i = 1; i < rosterData.length; i++) {
    const last = String(rosterData[i][COL_LAST - 1] || '').trim();
    const first = String(rosterData[i][COL_FIRST - 1] || '').trim();
    if (last && first) nameToRow.set(norm(`${last}, ${first}`), i + 1);
  }

  // Target column uses canonical title
  const targetHeader = `${weekKey} ${eventName}`;
  const attendanceCol = findAttendanceColumn(rosterHeader, targetHeader);
  if (!attendanceCol) return { matched: 0, updated: 0, missing: 0, notes: `Column not found: ${targetHeader}` };

  // Gather attendees from this one row
  const attendees = [];

  if (eventNorm === 'secondary') {
    // Pull from each "<Flight> Flight" column
    FLIGHT_NAMES.forEach(f => {
      const colTitle = `${f} Flight`;
      const val = readCell(idx, row, colTitle);
      attendees.push(...parseNamesCsv(val));
    });

  } else if (eventNorm === 'mando' || eventNorm === 'llab') {
    // Use the chosen flight if provided; otherwise infer from which section has AS selections
    let chosenFlight = flightRaw;
    if (!chosenFlight) chosenFlight = detectFlightFromRow(headers, row);

    // Cross-Town for Mando → parse like Secondary (across flights), but still write to the Mando column.
    if (eventNorm === 'mando' && isCrossTownFlight(chosenFlight)) {
      FLIGHT_NAMES.forEach(f => {
        const colTitle = `${f} Flight`;
        const val = readCell(idx, row, colTitle);
        attendees.push(...parseNamesCsv(val));
      });

    } else {
      if (!chosenFlight) {
        return { matched: 0, updated: 0, missing: 0, notes: 'No Flight selected and none detected' };
      }
      if (IGNORE_FLIGHTS.has(norm(chosenFlight))) {
        return { matched: 0, updated: 0, missing: 0, notes: 'Ignored abroad' };
      }

      const layoutAll = buildResponseLayoutByIndex(headers);
      const lay = layoutAll.get(chosenFlight);
      if (!lay) return { matched: 0, updated: 0, missing: 0, notes: `AS columns not found for flight ${chosenFlight}` };

      AS_GROUP_TITLES.forEach(as => {
        const colIndex = lay.asCols.get(as); // 1-based
        if (!colIndex) return;
        const cellVal = row[colIndex - 1];
        attendees.push(...parseNamesCsv(cellVal));
      });
    }
  } // ← this closes the else-if (mando/llab) block

  // Dedup + write
  const uniq = Array.from(new Set(attendees.map(n => norm(n))));
  const colIndex = attendanceCol.index;
  const colVals = roster.getRange(1, colIndex, rosterLastRow, 1).getValues();
  let updated = 0, missing = 0;

  uniq.forEach(key => {
    const r = nameToRow.get(key);
    if (!r) { missing++; return; }
    const current = String(colVals[r - 1][0] || '').trim();
    if (current !== 'P') {
      colVals[r - 1][0] = 'P';
      updated++;
    }
  });

  if (updated > 0) roster.getRange(1, colIndex, rosterLastRow, 1).setValues(colVals);

  // Add a compact note for logs
  const noteFlight = (eventNorm === 'secondary')
    ? ''
    : (flightRaw || `(detected ${detectFlightFromRow(headers, row) || '?'})`);

  return {
    matched: uniq.length,
    updated,
    missing,
    notes: `Auto ${weekKey} ${eventName}${noteFlight ? ` / ${noteFlight}` : ''}`
  };
}

/***** BATCH PARSER (used by menu) *****/
function parseSubmissionsBatch({ flightSelection, weekKey, eventNorm, userEmail }) {
  const ss = SpreadsheetApp.getActive();
  const responses = ss.getSheetByName(RESPONSES_SHEET_NAME);
  if (!responses) throw new Error(`Responses sheet "${RESPONSES_SHEET_NAME}" not found`);
  const roster = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!roster) throw new Error(`Roster sheet "${ROSTER_SHEET_NAME}" not found`);

  const respLastRow = responses.getLastRow();
  if (respLastRow < 2) return { matched: 0, updated: 0, missing: 0, notes: 'No responses found.' };

  const headers = responses.getRange(1, 1, 1, responses.getLastColumn()).getValues()[0];
  const rows = responses.getRange(2, 1, respLastRow - 1, responses.getLastColumn()).getValues();

  // Roster header + map
  const rosterLastRow = roster.getLastRow();
  const rosterLastCol = roster.getLastColumn();
  const rosterData = roster.getRange(1, 1, rosterLastRow, rosterLastCol).getValues();
  const rosterHeader = rosterData[0];
  const nameToRow = new Map();
  for (let i = 1; i < rosterData.length; i++) {
    const last = String(rosterData[i][COL_LAST - 1] || '').trim();
    const first = String(rosterData[i][COL_FIRST - 1] || '').trim();
    if (last && first) nameToRow.set(norm(`${last}, ${first}`), i + 1);
  }

  // Target attendance column
  const attendanceHeaderTarget = `${weekKey} ${capitalize(eventNorm)}`; // e.g., TW-02 Secondary
  const attendanceCol = findAttendanceColumn(rosterHeader, attendanceHeaderTarget);
  if (!attendanceCol) return { matched: 0, updated: 0, missing: 0, notes: `Column "${attendanceHeaderTarget}" not found.` };

  const idx = indexHeaders(headers);
  const attendees = [];
  let matched = 0;

  for (const row of rows) {
    const rowWeek  = normalizeTWKeyAttendance_(String(readCell(idx, row, ATT_HDR_TW) || ''));
    const rowEvent = normalizeEventTitle_(String(readCell(idx, row, ATT_HDR_EVENT) || '')).toLowerCase();
    if (rowWeek !== weekKey || rowEvent !== eventNorm) continue;

    if (eventNorm === 'secondary') {
      const canParseAll = SUPER_USERS.has(userEmail) || flightSelection === 'All Flights';
      const flightsToUse = canParseAll ? FLIGHT_NAMES : [flightSelection];
      flightsToUse.forEach(f => {
        const colTitle = `${f} Flight`;
        const parsed = parseNamesCsv(readCell(idx, row, colTitle));
        if (parsed.length) { attendees.push(...parsed); matched += parsed.length; }
      });
      continue;
    }

    // 'mando' or 'llab'
    const chosenFlight = String(readCell(idx, row, ATT_HDR_FLIGHT) || '').trim();

    // Cross-Town mando → parse like Secondary
    if (eventNorm === 'mando' && isCrossTownFlight(chosenFlight)) {
      const canParseAll = SUPER_USERS.has(userEmail) || flightSelection === 'All Flights';
      const flightsToUse = canParseAll ? FLIGHT_NAMES : [flightSelection];
      flightsToUse.forEach(f => {
        const colTitle = `${f} Flight`;
        const parsed = parseNamesCsv(readCell(idx, row, colTitle));
        if (parsed.length) { attendees.push(...parsed); matched += parsed.length; }
      });
      continue;
    }

    if (!chosenFlight) continue;
    if (IGNORE_FLIGHTS.has(norm(chosenFlight))) continue;
    if (flightSelection !== 'All Flights' && chosenFlight !== flightSelection) continue;

    const layout = buildResponseLayoutByIndex(headers);
    const lay = layout.get(chosenFlight);
    if (!lay) return { matched: 0, updated: 0, missing: 0, notes: `AS columns not found for flight ${chosenFlight}` };

    AS_GROUP_TITLES.forEach(as => {
      const colIndex = lay.asCols.get(as); // 1-based
      if (!colIndex) return;
      const cellVal = row[colIndex - 1];
      const parsed = parseNamesCsv(cellVal);
      if (parsed.length) { attendees.push(...parsed); matched += parsed.length; }
    });
  }

  // Dedup + write
  const uniq = Array.from(new Set(attendees.map(n => norm(n))));
  const colIndex = attendanceCol.index;
  const colVals = roster.getRange(1, colIndex, rosterLastRow, 1).getValues();
  let updated = 0, missing = 0;

  uniq.forEach(key => {
    const r = nameToRow.get(key);
    if (!r) { missing++; return; }
    const current = String(colVals[r - 1][0] || '').trim();
    if (current !== 'P') {
      colVals[r - 1][0] = 'P';
      updated++;
    }
  });
  if (updated > 0) roster.getRange(1, colIndex, rosterLastRow, 1).setValues(colVals);

  return { matched, updated, missing, notes: `Flight=${flightSelection}` };
}

/***** HELPERS *****/
function getHeadersFromSheet(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found`);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function indexHeaders(headers) {
  const map = new Map();
  headers.forEach((h, i) => map.set(String(h || '').trim(), i + 1));
  return map;
}

function readCell(idxMap, row, headerTitle) {
  const col = idxMap.get(headerTitle);
  if (!col) return '';
  return row[col - 1];
}

// Remove Forms' disambiguation suffixes like " (1)" or " (Alpha Flight)"
function canonicalizeASTitle(h) {
  let t = String(h || '').trim();
  // Strip one trailing (...) group
  t = t.replace(/\s*\([^)]*\)\s*$/, '');
  return t;
}

// Convert e.namedValues into a row array ordered by `headers`
function mapNamedValuesToRow(headers, namedValues) {
  if (!namedValues) return [];
  return headers.map(h => {
    const v = namedValues[h];
    if (v == null) return '';
    return Array.isArray(v) ? v.join(', ') : v;
  });
}

// Case-insensitive equals with whitespace normalized
function equalsCI(a, b) { return norm(a) === norm(b); }
function norm(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function capitalize(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

// "<Last, First, Last, First, ...>" -> ["Last, First", ...]
function parseNamesCsv(val) {
  const s = String(val || '').trim();
  if (!s || equalsCI(s, 'n/a')) return [];
  const parts = s.split(',').map(x => x.trim()).filter(Boolean);
  const names = [];
  for (let i = 0; i < parts.length - 1; i += 2) {
    names.push(`${parts[i]}, ${parts[i + 1]}`);
  }
  return names;
}

// Look through the row's AS checkboxes and infer which flight section was used.
function detectFlightFromRow(headers, row) {
  const layout = buildResponseLayoutByIndex(headers); // Map<flight, {asCols: Map<AS, colIndex>}>
  const nonEmptyFlights = [];

  for (const flight of FLIGHT_NAMES) {
    const lay = layout.get(flight);
    if (!lay) continue;

    let hasAny = false;
    for (const as of AS_GROUP_TITLES) {
      const colIndex = lay.asCols.get(as);
      if (!colIndex) continue;
      const val = row[colIndex - 1];
      const s = String(val || '').trim();
      if (s && !/^n\/?a$/i.test(s)) {
        hasAny = true;
        break;
      }
    }
    if (hasAny) nonEmptyFlights.push(flight);
  }

  return (nonEmptyFlights.length === 1) ? nonEmptyFlights[0] : '';
}

function processLastResponseManually() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  if (!sheet) throw new Error(`Responses sheet "${RESPONSES_SHEET_NAME}" not found`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No response rows to process.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowVals  = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const res = handleSingleSubmission(headers, rowVals);
  Logger.log(JSON.stringify(res));
  SpreadsheetApp.getUi().alert(
    `Matched: ${res.matched}\nUpdated "P": ${res.updated}\nMissing: ${res.missing}\nNotes: ${res.notes}`
  );
}

// Normalize "TW-1", "tw01", "Tw 15" → "TW-01" … "TW-15"
function normalizeTWKeyAttendance_(tw) {
  const m = String(tw || '').trim().match(/tw[-\s]?(\d{1,2})/i);
  if (!m) return '';
  const n = Number(m[1]);
  if (!isFinite(n) || n < 1 || n > 16) return '';
  return `TW-${String(n).padStart(2, '0')}`;
}

// Normalize event text → "Secondary" | "LLAB" | "Mando" (with nice casing)
function normalizeEventTitle_(ev) {
  const e = String(ev || '').toLowerCase().trim();
  if (e.startsWith('sec')) return 'Secondary';
  if (e === 'llab')       return 'LLAB';
  if (e.startsWith('mand')) return 'Mando';
  return e ? e.charAt(0).toUpperCase() + e.slice(1) : '';
}

// Find the column index for a given attendance header in the roster sheet.
// Returns { index: <1-based column>, header: <string> } or null if not found.
function findAttendanceColumn(rosterHeaderRow, targetHeader) {
  const targetNorm = norm(targetHeader);
  for (let c = ATTENDANCE_HEADERS_START_COL - 1; c < rosterHeaderRow.length; c++) {
    const h = String(rosterHeaderRow[c] || '');
    if (norm(h) === targetNorm) {
      return { index: c + 1, header: h };
    }
  }
  return null;
}

// Map the responses header row into per-flight AS column indexes.
// Returns: Map<flight, { asCols: Map<AS title, 1-based colIndex> }>
function buildResponseLayoutByIndex(headers) {
  const layout = new Map();
  const flightHeaderPositions = [];

  // Record all "<Flight> Flight" header positions (0-based indexes)
  for (let i = 0; i < headers.length; i++) {
    const title = String(headers[i] || '').trim();
    const flightMatch = FLIGHT_NAMES.find(f => equalsCI(title, `${f} Flight`));
    if (flightMatch) {
      flightHeaderPositions.push({ flight: flightMatch, index: i });
    }
  }

  // Quick check: is a given header index itself a "<Flight> Flight" header?
  function isFlightHeaderAt(j) {
    const t = String(headers[j] || '').trim();
    return FLIGHT_NAMES.some(f => equalsCI(t, `${f} Flight`));
  }

  // For each flight section, bind the following AS columns until next flight header
  for (let k = 0; k < flightHeaderPositions.length; k++) {
    const { flight, index: startIdx } = flightHeaderPositions[k];
    const endIdx = (k + 1 < flightHeaderPositions.length)
      ? flightHeaderPositions[k + 1].index
      : headers.length;

    const asMap = new Map(); // AS title -> 1-based col index
    for (let j = startIdx + 1; j < endIdx; j++) {
      if (isFlightHeaderAt(j)) break; // safety
      const h = String(headers[j] || '').trim();
      const canon = canonicalizeASTitle(h); // e.g., strip "(Alpha Flight)" or "(1)"
      if (AS_GROUP_TITLES.includes(canon) && !asMap.has(canon)) {
        asMap.set(canon, j + 1); // store 1-based
      }
      if (asMap.size === AS_GROUP_TITLES.length) break; // found all four AS groups
    }

    layout.set(flight, { asCols: asMap });
  }

  return layout;
}
```

```CleanWhitespace.gs
// --- Add this helper (uses the same logic as cleanAllFormResponses, but for any sheet name)
function cleanResponsesSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found`);

  const firstDataRow = 2; // keep row 1 as header
  const lastRow = sh.getLastRow();
  if (lastRow < firstDataRow) {
    Logger.log(`No data rows to clean in "${sheetName}".`);
    return;
  }

  const lastCol = sh.getLastColumn();
  const lock = LockService.getDocumentLock();
  lock.tryLock(30000);
  try {
    const BATCH = 400;
    for (let start = firstDataRow; start <= lastRow; start += BATCH) {
      const num = Math.min(BATCH, (lastRow - start + 1));
      const range = sh.getRange(start, 1, num, lastCol);
      const values = range.getValues();
      const cleaned = values.map(row =>
        row.map(cell => (typeof cell === 'string' ? normalizeText(cell) : cell))
      );
      range.setValues(cleaned);
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    }
    Logger.log(`Cleaned rows ${firstDataRow}-${lastRow} in "${sheetName}".`);
  } finally {
    lock.releaseLock();
  }
}

// keep your original menu action working exactly the same
function cleanAllFormResponses() {
  cleanResponsesSheet_('Attendance Form'); // or change to RESPONSES_SHEET_NAME if you prefer
}

// unchanged
function normalizeText(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return v;
  let s = v.replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}
```

```ExcusalEmails.gs
/***** CONFIG *****/
const EXCUSAL_SHEET_NAME = 'Excusal Requests';
const CMD_INFO_SHEET_NAME = 'Squadron Commander Info';

// Attendance workbook (separate file)
const ATTENDANCE_FILE_ID_OR_URL = 'https://docs.google.com/spreadsheets/d/1MaPFLz5N8ngR399IrJHsqDMDr4p1mxO8l6OWcg4rKYU';
const ATTENDANCE_SHEET_NAME = 'Attendance';

// === Excusal header titles (exact, including punctuation/colon) ===
const HDR_DECISION  = 'Approve/Denied';
const HDR_EMAIL     = 'Email Address';
const HDR_LAST      = 'Last Name';
const HDR_FIRST     = 'First Name';
const HDR_TW        = 'Training Week (Format as TW-00) of Your Absence:';
const HDR_EVENT     = 'Event';
const HDR_REASON    = 'Reason for Absence';
const HDR_COMMANDER = 'Select Your Squadron Commander';
const HDR_MFR = 'Upload filled out MFR found here. You can find comments on how to fill it out here.';

// Roster columns (1-based) in Attendance
const COL_YEAR   = 1; // A
const COL_LAST   = 2; // B
const COL_FIRST  = 3; // C
const COL_EMAIL  = 4; // D
const COL_FLIGHT = 5; // E

// Attendance headers start at column H (1-based)
const ATTENDANCE_HEADERS_START_COL = 8; // H

// Allowed existing marks we do NOT overwrite on initial submit
const ALLOWED_CODES = new Set(['P', 'E', 'ES', 'T', 'UR', 'U', 'MU', 'MRS']);

// Values used by decisions
const CODE_ER = 'ER';
const CODE_E  = 'E';
const CODE_ED = 'ED';

// Allowed existing marks we CAN overwrite when making a decision
const DECISION_OVERWRITABLE_CODES = new Set(['ER', 'ED']);

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Send Email')
    .addItem('Send decision email for active row (from me)', 'menuSendDecisionEmailForActiveRow')
    .addToUi();
}

/***** TRIGGERS *****/
function installExcusalOnFormSubmit() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onExcusalFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onExcusalFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

function installExcusalOnEdit() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onExcusalEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onExcusalEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/***** HELPERS YOU PROVIDED *****/
function getSpreadsheetByIdOrUrl(idOrUrl) {
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('Invalid Spreadsheet ID/URL: ' + idOrUrl);
  const id = m[0];
  DriveApp.getFileById(id); // permission check
  return SpreadsheetApp.openById(id);
}

// Normalize text (trim, collapse whitespace, replace funky spaces)
function normalizeText(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return v;
  let s = v.replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// Clean a sheet’s string fields in place (optional to call before reads)
function cleanResponsesSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found`);

  const firstDataRow = 2;
  const lastRow = sh.getLastRow();
  if (lastRow < firstDataRow) return;

  const lastCol = sh.getLastColumn();
  const lock = LockService.getDocumentLock();
  lock.tryLock(30000);
  try {
    const BATCH = 400;
    for (let start = firstDataRow; start <= lastRow; start += BATCH) {
      const num = Math.min(BATCH, (lastRow - start + 1));
      const range = sh.getRange(start, 1, num, lastCol);
      const values = range.getValues();
      const cleaned = values.map(row =>
        row.map(cell => (typeof cell === 'string' ? normalizeText(cell) : cell))
      );
      range.setValues(cleaned);
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    }
  } finally {
    lock.releaseLock();
  }
}

/***** CORE: ATTENDANCE UPDATE *****/
function findAttendanceColumn_(headerRow, tw, event) {
  const target = norm(`${tw} ${event}`);
  for (let c = ATTENDANCE_HEADERS_START_COL - 1; c < headerRow.length; c++) {
    if (norm(headerRow[c]) === target) return c + 1; // 1-based
  }
  return null;
}

function buildRosterNameMap_(attendanceSheet) {
  const lastRow = attendanceSheet.getLastRow();
  const lastCol = attendanceSheet.getLastColumn();
  const data = attendanceSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0];
  const body = data.slice(1);
  const nameToRow = new Map(); // "last, first" (norm) -> row#
  for (let i = 0; i < body.length; i++) {
    const last = String(body[i][COL_LAST - 1] || '').trim();
    const first = String(body[i][COL_FIRST - 1] || '').trim();
    if (last && first) nameToRow.set(norm(`${last}, ${first}`), i + 2);
  }
  return { header, nameToRow, lastRow, lastCol };
}

function setAttendanceCode_(attendanceSheet, nameToRow, header, tw, event, last, first, desiredCode, onlyReplaceUR, protectExisting) {
  const col = findAttendanceColumn_(header, tw, event);
  if (!col) throw new Error(`Attendance column not found: "${tw} ${event}"`);
  const row = nameToRow.get(norm(`${last}, ${first}`));
  if (!row) return { updated: false, reason: 'name not in roster' };

  const current = String(attendanceSheet.getRange(row, col).getValue() || '').trim();

  // Initial submit: set UR only if current not in ALLOWED_CODES (protect existing marks)
  if (protectExisting && !ALLOWED_CODES.has(current)) {
    attendanceSheet.getRange(row, col).setValue(CODE_ER);
    return { updated: true, newVal: CODE_ER, prev: current };
  }

  // Decision step: change UR→E or UR→UD; if onlyReplaceUR true, ignore other current values
  if (onlyReplaceUR) {
    if (current === CODE_ER && (desiredCode === CODE_E || desiredCode === CODE_ED)) {
      attendanceSheet.getRange(row, col).setValue(desiredCode);
      return { updated: true, newVal: desiredCode, prev: CODE_ER };
    }
    return { updated: false, reason: `current "${current}" is not UR` };
  }

  // Fallback generic set
  attendanceSheet.getRange(row, col).setValue(desiredCode);
  return { updated: true, newVal: desiredCode, prev: current };
}

// Safely extract a scalar string from namedValues (arrays -> joined)
function nvGet_(nv, key) {
  if (!nv || !(key in nv)) return '';
  const v = nv[key];
  if (Array.isArray(v)) return v.join(', ');
  return String(v ?? '');
}

// 1) On initial form submit to Excusal Requests → mark UR in Attendance (if needed)
function onExcusalFormSubmit(e) {
  try {
    cleanResponsesSheet_(EXCUSAL_SHEET_NAME);

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(EXCUSAL_SHEET_NAME);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    let lastRaw = '', firstRaw = '', twRaw = '', eventRaw = '', reasonRaw = '', commanderSel = '', cadetEmail = '';
    let mfrLinkRaw = '';

    if (e && e.namedValues) {
      lastRaw      = nvGet_(e.namedValues, HDR_LAST);
      firstRaw     = nvGet_(e.namedValues, HDR_FIRST);
      twRaw        = nvGet_(e.namedValues, HDR_TW);
      eventRaw     = nvGet_(e.namedValues, HDR_EVENT);
      reasonRaw    = nvGet_(e.namedValues, HDR_REASON);
      commanderSel = nvGet_(e.namedValues, HDR_COMMANDER);
      cadetEmail   = nvGet_(e.namedValues, HDR_EMAIL);
      mfrLinkRaw   = nvGet_(e.namedValues, HDR_MFR);
      Logger.log(`[EAR submit] using namedValues`);
    } else if (e && e.values) {
      const row = e.values;
      const idx = indexHeaders_(headers);
      lastRaw      = readByHeader_(idx, row, HDR_LAST);
      firstRaw     = readByHeader_(idx, row, HDR_FIRST);
      twRaw        = readByHeader_(idx, row, HDR_TW);
      eventRaw     = readByHeader_(idx, row, HDR_EVENT);
      reasonRaw    = readByHeader_(idx, row, HDR_REASON);
      commanderSel = readByHeader_(idx, row, HDR_COMMANDER);
      cadetEmail   = readByHeader_(idx, row, HDR_EMAIL);
      mfrLinkRaw   = readByHeader_(idx, row, HDR_MFR);
      Logger.log(`[EAR submit] using values[] + header map`);
    } else {
      const lastRowNum = sh.getLastRow();
      if (lastRowNum < 2) { Logger.log('[EAR submit] No rows.'); return; }
      const row = sh.getRange(lastRowNum, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = indexHeaders_(headers);
      lastRaw      = readByHeader_(idx, row, HDR_LAST);
      firstRaw     = readByHeader_(idx, row, HDR_FIRST);
      twRaw        = readByHeader_(idx, row, HDR_TW);
      eventRaw     = readByHeader_(idx, row, HDR_EVENT);
      reasonRaw    = readByHeader_(idx, row, HDR_REASON);
      commanderSel = readByHeader_(idx, row, HDR_COMMANDER);
      cadetEmail   = readByHeader_(idx, row, HDR_EMAIL);
      mfrLinkRaw   = readByHeader_(idx, row, HDR_MFR);
      Logger.log(`[EAR submit] using last-row fallback`);
    }

    // Normalize fields
    const last   = normalizeText(lastRaw);
    const first  = normalizeText(firstRaw);
    const twKey  = normalizeTWKey_(twRaw);
    const event  = normalizeEvent_(eventRaw); // now includes "Other"
    const reason = String(reasonRaw || 'No reason provided');
    const cadet  = cadetEmail || '';

    Logger.log(`[EAR submit] Extracted: last="${last}", first="${first}", twRaw="${twRaw}"→"${twKey}", eventRaw="${eventRaw}"→"${event}", commanderSel="${commanderSel}"`);

    if (!last || !first || !twKey || !event) {
      Logger.log('[EAR submit] Missing one of required fields; aborting.');
      return;
    }

    // === Attendance marking (skip when "Other") ===
    if (event !== 'Other') {
      const attSS = getSpreadsheetByIdOrUrl(ATTENDANCE_FILE_ID_OR_URL);
      const attSh = attSS.getSheetByName(ATTENDANCE_SHEET_NAME);
      if (!attSh) throw new Error(`Attendance sheet "${ATTENDANCE_SHEET_NAME}" not found`);

      const { header, nameToRow } = buildRosterNameMap_(attSh);
      const col = findAttendanceColumnFlexible_(header, twKey, event);
      if (!col) {
        Logger.log(`[EAR submit] Column not found for "${twKey} ${event}". Headers from H: ` +
                   header.slice(ATTENDANCE_HEADERS_START_COL - 1).join(' | '));
      } else {
        const res = setAttendanceCodeAt_(attSh, nameToRow, col, last, first, CODE_ER, /*protectExisting*/true);
        Logger.log(`[EAR submit] ${res.updated ? 'Set' : 'Skipped'} UR for ${last}, ${first} at col=${col} (prev=${res.prev || ''}, reason=${res.reason || 'ok'})`);
        SpreadsheetApp.flush();
      }
    } else {
      Logger.log('[EAR submit] Event is "Other" → skipping attendance update.');
    }

    // === Email the relevant Squadron Commander (always, including "Other") ===
    const dir = buildCommanderDirectory_();
    let info = dir.get(norm(commanderSel));
    if (!info) {
      const { last: cLast, first: cFirst } = parseLastFirst_(commanderSel);
      const key = norm(`${cLast}, ${cFirst}`);
      info = dir.get(key);
    }
    const commanderEmail = (info && info.email) ? info.email : '';
    const commanderName  = (info && info.displayName) ? info.displayName : getDisplayNameSafe_(commanderEmail);

    if (!commanderEmail) {
      Logger.log(`[EAR submit] Could not resolve commander email from selection "${commanderSel}". Not sending commander notification.`);
      return;
    }

    // Greeting: Good morning/afternoon/evening C/Lastname
    const { last: fcLast } = parseLastFirst_(commanderSel);
    const greeting = `${getGreeting_()} C/${fcLast}`;

    const subjectFC = `New EAR Submitted: ${last}, ${first} — ${twKey} ${event}`;
    const { attachments, linkText } = buildAttachmentOrLink_(mfrLinkRaw);

    const bodyFC =
`${greeting},

You have received a new Excused Absence Request (EAR) from Cadet ${first} ${last}.

Details:
• Cadet: ${first} ${last} ${cadet ? `(${cadet})` : ''}
• TW/Event: ${twKey} ${event}
• Reason: ${reason}

Review & take action here:
${REQUESTS_SPREADSHEET_URL}

${linkText ? `MFR (PDF) provided by cadet:\n${linkText}\n` : ''}
Very Respectfully,
EAR Automations`;

    const mailOpts = {
      name: 'EAR Automations',
      replyTo: cadet || undefined,
      cc: cadet || undefined,
      attachments: attachments.length ? attachments : undefined
    };

    GmailApp.sendEmail(commanderEmail, subjectFC, bodyFC, mailOpts);
    Logger.log(`[EAR submit] Commander notification sent to ${commanderEmail}`);

  } catch (err) {
    Logger.log('onExcusalFormSubmit error: ' + (err && err.stack || err));
    throw err;
  }
}

// 2) On edit of Approve/Denied column in Excusal Requests → change UR→E or UR→UD and send email
function onExcusalEdit(e) {
  try {
    const range = e && e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (sheet.getName() !== EXCUSAL_SHEET_NAME) return;

    // Build header map to identify the Approve/Denied column by title
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = indexHeaders_(headers);
    const colDecision = idx.get(HDR_DECISION);
    if (!colDecision) {
      Logger.log(`[EAR edit] Could not find "${HDR_DECISION}" header.`);
      return;
    }

    // Only act when the Approve/Denied cell is edited
    if (range.getColumn() !== colDecision || range.getRow() === 1) return;

    const rowNum = range.getRow();
    const rowVals = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Extract fields by header (robust to column re-ordering)
    const cadetEmail   = normalizeText(readByHeader_(idx, rowVals, HDR_EMAIL));
    const lastNameRaw  = readByHeader_(idx, rowVals, HDR_LAST);
    const firstNameRaw = readByHeader_(idx, rowVals, HDR_FIRST);
    const twRaw        = readByHeader_(idx, rowVals, HDR_TW);
    const eventRaw     = readByHeader_(idx, rowVals, HDR_EVENT);
    const reasonRaw    = readByHeader_(idx, rowVals, HDR_REASON);
    const commanderSel = normalizeText(readByHeader_(idx, rowVals, HDR_COMMANDER));
    const statusRaw    = normalizeText(readByHeader_(idx, rowVals, HDR_DECISION));

    const lastName = normalizeText(lastNameRaw);
    const firstName = normalizeText(firstNameRaw);
    const reason = String(reasonRaw || 'No reason provided');

    // Normalize TW + Event
    const twKey = normalizeTWKey_(twRaw);       // -> "TW-01"
    const event = normalizeEvent_(eventRaw);    // -> "Mando" | "LLAB" | "Secondary"

    // If event is "Other", skip Attendance write (still send email to cadet)
    const isOther = (event === 'Other');

    if (!lastName || !firstName || !twKey || !event || !statusRaw) {
      Logger.log(`[EAR edit] Missing required fields. last="${lastName}", first="${firstName}", tw="${twRaw}", event="${eventRaw}", status="${statusRaw}"`);
      return;
    }

    // Canonical decision
    const decision =
      (/^approved$/i.test(statusRaw) || /^approve$/i.test(statusRaw)) ? 'Approved' :
      (/^denied$/i.test(statusRaw)   || /^deny$/i.test(statusRaw))    ? 'Denied'   :
      statusRaw;

    if (!isOther) {
      // Open Attendance and locate the cell
      const attSS = getSpreadsheetByIdOrUrl(ATTENDANCE_FILE_ID_OR_URL);
      const attSh = attSS.getSheetByName(ATTENDANCE_SHEET_NAME);
      if (!attSh) throw new Error(`Attendance sheet "${ATTENDANCE_SHEET_NAME}" not found`);

      const { header, nameToRow } = buildRosterNameMap_(attSh);
      const col = findAttendanceColumnFlexible_(header, twKey, event);
      if (!col) {
        Logger.log(`[EAR edit] Attendance column not found for "${twKey} ${event}".`);
        return;
      }
      const rosterRow = nameToRow.get(norm(`${lastName}, ${firstName}`));
      if (!rosterRow) {
        Logger.log(`[EAR edit] Cadet not found on roster: "${lastName}, ${firstName}"`);
        return;
      }

      // Only change UR → E or UR → UD (leave other codes alone)
      const cell = attSh.getRange(rosterRow, col);
      const current = String(cell.getValue() || '').trim();
      if (!DECISION_OVERWRITABLE_CODES.has(current)) {
        Logger.log(`[EAR edit] Skip update; current="${current}" (only change ${Array.from(DECISION_OVERWRITABLE_CODES).join(', ')}).`);
      } else {
        if (/^approved$/i.test(decision)) {
          cell.setValue(CODE_E);
          Logger.log(`[EAR edit] Set E for ${lastName}, ${firstName} (${twKey} ${event}).`);
        } else if (/^denied$/i.test(decision)) {
          cell.setValue(CODE_ED);
          Logger.log(`[EAR edit] Set UD for ${lastName}, ${firstName} (${twKey} ${event}).`);
        } else {
          Logger.log(`[EAR edit] Unknown decision "${decision}" — no change.`);
        }
      }
    } else {
      Logger.log('[EAR edit] Event is "Other" → skipping attendance update on decision.');
    }

    // Commander directory (reply-to + signature)
    const dir = buildCommanderDirectory_();
    const info = dir.get(norm(commanderSel)) || { email: '', signature: '', displayName: '' };
    const approverEmail = info.email || Session.getActiveUser().getEmail();
    const approverName  = info.displayName || (Session.getActiveUser().getEmail() || 'Approver');

    // Email body with time-of-day greeting
    const greeting = `${getGreeting_()} Cadet ${firstName} ${lastName},`;
    const subject  = `EAR ${decision}: ${lastName}, ${firstName} — ${twKey} ${event}`;
    const signatureBlock = info.signature ? info.signature : '';

    const body =
`${greeting}

Your Excused Absence Request for ${twKey} ${event} has been ${decision.toUpperCase()}.

Your reason provided: "${reason}"

If you have questions or would like to appeal, contact your flight commander through the squadron chain of command.

Very Respectfully,
${signatureBlock}`;

    // Send (auto-trigger path): from script identity; replies go to approver
    GmailApp.sendEmail(cadetEmail || approverEmail, subject, body, {
      replyTo: approverEmail,
      name: approverName,
      cc: approverEmail
    });

  } catch (err) {
    Logger.log('onExcusalEdit error: ' + (err && err.stack || err));
    throw err;
  }
}

/***** MENU ACTION: send from current user *****/
// Select a cell in the row, then run this.
// This fakes an edit event so we can reuse onExcusalEdit() logic.
function menuSendDecisionEmailForActiveRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== EXCUSAL_SHEET_NAME) {
    SpreadsheetApp.getUi().alert(`Switch to "${EXCUSAL_SHEET_NAME}" first.`);
    return;
  }
  const r = sh.getActiveRange();
  if (!r) return;
  const row = r.getRow();
  if (row === 1) return;

  // Build a fake event that mimics an edit of the Approve/Denied column
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = indexHeaders_(headers);
  const colDecision = idx.get(HDR_DECISION);
  if (!colDecision) {
    SpreadsheetApp.getUi().alert(`Could not find "${HDR_DECISION}" header.`);
    return;
  }

  const decisionCell = sh.getRange(row, colDecision);
  const fakeEvent = {
    range: decisionCell,
    value: decisionCell.getValue(),
    oldValue: null, // optional, not needed here
    source: SpreadsheetApp.getActive()
  };

  // Call the main handler
  onExcusalEdit(fakeEvent);
}

/***** SMALL UTILS *****/
function indexHeaders_(headers) {
  const map = new Map();
  headers.forEach((h, i) => map.set(String(h || '').trim(), i + 1));
  return map;
}
function readByHeader_(idx, row, title) {
  const col = idx.get(title);
  return col ? row[col - 1] : '';
}
function norm(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

// Try to get a pretty name from the approver’s account (best-effort)
function getDisplayNameSafe_(email) {
  try {
    // Basic: use part before '@' capitalized
    if (!email) return '';
    const namePart = String(email).split('@')[0].replace(/[._]/g, ' ');
    return namePart.replace(/\b\w/g, c => c.toUpperCase());
  } catch (e) { return ''; }
}

// "Good morning/afternoon/evening"
function getGreeting_() {
  // Uses project time zone
  const hour = new Date().getHours();
  if (hour < 12) return 'Good morning';
  if (hour < 17) return 'Good afternoon';
  return 'Good evening';
}

// Build a map: "last, first" (norm) -> { email, signature, displayName }
function buildCommanderDirectory_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CMD_INFO_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${CMD_INFO_SHEET_NAME}" not found`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Map();

  // Ensure we grab at least 4 columns (A:D)
  const data = sh.getRange(2, 1, lastRow - 1, Math.max(4, sh.getLastColumn())).getValues();
  const dir = new Map();
  for (const row of data) {
    const last = String(row[0] || '').trim();        // A: Last Name
    const first = String(row[1] || '').trim();       // B: First Name
    const email = String(row[2] || '').trim();       // C: Email
    const signature = String(row[3] || '').replace(/\r\n/g, '\n'); // D: Signature Block
    if (!last || !first) continue;
    const key = norm(`${last}, ${first}`);           // normalized "Last, First"
    const displayName = `${first} ${last}`;
    dir.set(key, { email, signature, displayName });
  }
  return dir;
}

// Parse "Last, First" -> {last, first}
function parseLastFirst_(s) {
  const p = String(s || '').split(',').map(x => x.trim());
  return { last: p[0] || '', first: p[1] || '' };
}
function norm(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

// Normalize TW input to "TW-##" (zero-padded)
function normalizeTWKey_(tw) {
  const s = normalizeText(tw);
  // Accept "TW-1", "tw01", "TW- 1", etc.
  const m = s.match(/tw[-\s]?(\d{1,2})/i);
  if (!m) return '';
  const n = parseInt(m[1], 10);
  if (isNaN(n) || n < 1 || n > 15) return '';
  return `TW-${String(n).padStart(2, '0')}`;
}

// Canonicalize event to exactly: "Mando" | "LLAB" | "Secondary" | "Other"
function normalizeEvent_(ev) {
  const s = normalizeText(ev).replace(/\./g, '');
  if (/(^|[^a-z])llab($|[^a-z])/i.test(s)) return 'LLAB';
  if (/sec(ondary)?/i.test(s)) return 'Secondary';
  if (/mando/i.test(s)) return 'Mando';
  if (/other/i.test(s)) return 'Other';
  return '';
}

// Flexible finder for the attendance column.
// Tries exact "TW-## Event" match (case/space-insensitive), then regex fallbacks like TW-# vs TW-0#
function findAttendanceColumnFlexible_(headerRow, twKey, event) {
  if (!twKey || !event) return null;
  const targetNorm = norm(`${twKey} ${event}`);
  for (let c = ATTENDANCE_HEADERS_START_COL - 1; c < headerRow.length; c++) {
    if (norm(headerRow[c]) === targetNorm) return c + 1;
  }
  // Fallback: accept TW-# (no zero) and extra spaces
  const n = parseInt(twKey.replace(/[^0-9]/g, ''), 10);
  if (!isNaN(n)) {
    const regex = new RegExp(`^\\s*tw-\\s*0?${n}\\s+${event}\\s*$`, 'i');
    for (let c = ATTENDANCE_HEADERS_START_COL - 1; c < headerRow.length; c++) {
      const h = String(headerRow[c] || '');
      if (regex.test(h)) return c + 1;
    }
  }
  return null;
}

// Direct set when we already know the target column index (1-based)
function setAttendanceCodeAt_(attendanceSheet, nameToRow, colIndex, last, first, desiredCode, protectExisting) {
  const row = nameToRow.get(norm(`${last}, ${first}`));
  if (!row) return { updated: false, reason: 'name not in roster' };

  const cell = attendanceSheet.getRange(row, colIndex);
  const current = String(cell.getValue() || '').trim();

  if (protectExisting && ALLOWED_CODES.has(current)) {
    return { updated: false, reason: `existing protected code "${current}"` };
  }

  cell.setValue(desiredCode);
  return { updated: true, newVal: desiredCode, prev: current };
}
```

```FlightCommander.gs
/***** CONFIG *****/
// Already declared const CMD_INFO_SHEET_NAME = 'Flight Command Info'; // sheet with Last/First/etc.
const FORM_ID_OR_URL = 'https://docs.google.com/forms/d/1kuIHzELTwxqgp90sDTYyuaf4vpA2aBf6WRZq1N04rAI';
const FORM_DROPDOWN_TITLE = 'Select Your Squadron Commander';

// If the list would otherwise be empty, we write a placeholder to avoid Form API errors
const WRITE_PLACEHOLDER_WHEN_EMPTY = true;
const EMPTY_PLACEHOLDER_LABEL = '(none)';

/***** HELPERS *****/
function getFormByIdOrUrl(idOrUrl) {
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('Invalid Form ID/URL: ' + idOrUrl);
  const id = m[0];
  DriveApp.getFileById(id); // permission check
  return FormApp.openById(id);
}

function normalizeText(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return v;
  let s = v.replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// Build ["Last, First", ...] from Flight Command Info sheet
function getCommanderOptions_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CMD_INFO_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${CMD_INFO_SHEET_NAME}" not found`);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, Math.max(2, sh.getLastColumn())).getValues();
  const names = [];
  for (const row of values) {
    const last = normalizeText(row[0]); // Col A
    const first = normalizeText(row[1]); // Col B
    if (last && first) names.push(`${last}, ${first}`);
  }
  // de-dup + sort
  const unique = Array.from(new Set(names));
  unique.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  return unique;
}

// Safe writer to a Form dropdown (LIST) item
function setListChoicesSafe_(listItem, values, ctx) {
  if (Array.isArray(values) && values.length > 0) {
    listItem.setChoiceValues(values);
    Logger.log(`[OK] ${ctx}: wrote ${values.length} choices.`);
  } else if (WRITE_PLACEHOLDER_WHEN_EMPTY) {
    listItem.setChoiceValues([EMPTY_PLACEHOLDER_LABEL]);
    Logger.log(`[OK] ${ctx}: no options; wrote placeholder "${EMPTY_PLACEHOLDER_LABEL}".`);
  } else {
    Logger.log(`[SKIP] ${ctx}: no options; left existing values unchanged.`);
  }
}

// Find a LIST item by exact title (case-insensitive)
function findListItemByTitle_(form, title) {
  const wanted = String(title || '').trim().toLowerCase();
  const items = form.getItems(FormApp.ItemType.LIST);
  for (const it of items) {
    const li = it.asListItem();
    if ((li.getTitle() || '').trim().toLowerCase() === wanted) return li;
  }
  return null;
}

/***** CORE SYNC *****/
function syncFlightCommandersToForm() {
  const form = getFormByIdOrUrl(FORM_ID_OR_URL);
  const listItem = findListItemByTitle_(form, FORM_DROPDOWN_TITLE);
  if (!listItem) {
    Logger.log(`Dropdown "${FORM_DROPDOWN_TITLE}" not found in the form.`);
    return;
  }
  const options = getCommanderOptions_();
  setListChoicesSafe_(listItem, options, `"${FORM_DROPDOWN_TITLE}"`);
}

/***** TRIGGERS *****/
// Simple trigger: fires when editing Last/First in Flight Command Info
function onEdit(e) {
  try {
    const range = e && e.range;
    if (!range) return;
    const sh = range.getSheet();
    if (sh.getName() !== CMD_INFO_SHEET_NAME) return;

    const col = range.getColumn();
    // Only react to edits in columns A (Last Name) or B (First Name)
    if (col !== 1 && col !== 2) return;

    // Optional: clean just the edited cells (keeps names tidy)
    const val = range.getValue();
    if (typeof val === 'string') {
      const cleaned = normalizeText(val);
      if (cleaned !== val) range.setValue(cleaned);
    }

    syncFlightCommandersToForm();
  } catch (err) {
    Logger.log('onEdit error: ' + (err && err.stack || err));
    throw err;
  }
}

// Optional reliability: run daily in case someone bulk-imports without edits
function installTimeTrigger_Daily() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFlightCommandersToForm')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncFlightCommandersToForm')
    .timeBased()
    .everyDays(1)
    .create();
}
```

```NotificationUtils.gs
const REQUESTS_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1hTw8DVStCcQac_PnZh8Ph6wlxYH8DCtZ2s7jD7etWpY';

// Extract first Google Drive file ID from a URL-ish string (works for open?id=, /d/<id>/, sharing links)
function getDriveFileIdFromUrl_(s) {
  if (!s) return '';
  const str = String(s);
  // /file/d/<id>/view ...
  let m = str.match(/\/d\/([-\w]{25,})/);
  if (m) return m[1];
  // open?id=<id>
  m = str.match(/[?&]id=([-\w]{25,})/);
  if (m) return m[1];
  // direct id-ish
  m = str.match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

// If a cell contains multiple links (comma/newline separated), grab the first non-empty
function firstToken_(s) {
  if (!s) return '';
  return String(s).split(/[\n,;]/).map(x => x.trim()).filter(Boolean)[0] || '';
}

// Try to attach a Drive file given a URL-ish value; returns {attachments, linkText}
function buildAttachmentOrLink_(maybeUrl) {
  const linkText = firstToken_((maybeUrl || '').toString());
  if (!linkText) return { attachments: [], linkText: '' };

  try {
    const id = getDriveFileIdFromUrl_(linkText);
    if (!id) return { attachments: [], linkText };
    const file = DriveApp.getFileById(id); // will throw if no access
    const blob = file.getBlob();
    return { attachments: [blob], linkText };
  } catch (e) {
    // No access or not a Drive file; fall back to link-only
    return { attachments: [], linkText };
  }
}
```

```GrabCadets.gs
/***** CONFIG *****/
const FORM_ID_OR_URL  = 'https://docs.google.com/forms/d/1Qzxs5rtXOu9vIgbL9RWe-6JtmHvC89jhRCKFLMkIAxU';
const SHEET_ID_OR_URL = 'https://docs.google.com/spreadsheets/d/1MaPFLz5N8ngR399IrJHsqDMDr4p1mxO8l6OWcg4rKYU';
const SHEET_NAME = 'Attendance';

// Column indexes (1-based)
const COL_YEAR  = 1; // A
const COL_LAST  = 2; // B
const COL_FIRST = 3; // C
const COL_FLIGHT= 5; // E

// Map Sheet year values -> Form group titles
const YEAR_TO_GROUP = {
  'as100': 'AS 100/150',
  'as150': 'AS 100/150',
  'as200': 'AS 200/250',
  'as250': 'AS 200/250',
  'as300': 'AS 300',
  'as400': 'AS 400',
};

// Four **checkbox** question titles expected per-flight section
const AS_GROUP_TITLES = ['AS 100/150', 'AS 200/250', 'AS 300', 'AS 400'];

// Ignore these flights entirely (case-insensitive)
const IGNORE_FLIGHTS = new Set(['abroad']);

// When a checkbox would have zero choices, either show a placeholder or skip writing.
// Option A (default): write a placeholder so the section is visibly empty.
const WRITE_PLACEHOLDER_WHEN_EMPTY = true;
const EMPTY_PLACEHOLDER_LABEL = 'N/A'; // change if you prefer a different label

/***** MAIN *****/
function syncFormCheckboxesFromSheet() {
  const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

  const form = getFormByIdOrUrl(FORM_ID_OR_URL);
  const ss   = getSpreadsheetByIdOrUrl(SHEET_ID_OR_URL);
  const sheet= ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const lastRow = sheet.getLastRow();
  Logger.log(`Found lastRow=${lastRow} on sheet "${SHEET_NAME}".`);
  if (lastRow < 2) {
    Logger.log('No data rows found. Exiting.');
    return;
  }

  const rng = sheet.getRange(2, 1, lastRow - 1, COL_FLIGHT);
  const values = rng.getValues();

  // Build structures:
  // byFlight[flight][groupTitle] = [ "Last, First", ... ]
  // byFlightAll[flight] = [ "Last, First", ... ]
  const byFlight = {};
  const byFlightAll = {};

  let inRows = 0, ignoredRows = 0, missingYearRows = 0;
  values.forEach((row, idx) => {
    const year   = norm(row[COL_YEAR  - 1]);
    const last   = String(row[COL_LAST  - 1] || '').trim();
    const first  = String(row[COL_FIRST - 1] || '').trim();
    const flightRaw = String(row[COL_FLIGHT - 1] || '').trim();
    const flight = norm(flightRaw);

    // Ignore incomplete
    if (!year || !last || !first || !flight) return;

    // Ignore flights
    if (IGNORE_FLIGHTS.has(flight)) { ignoredRows++; return; }

    // Year to group
    const group = YEAR_TO_GROUP[year];
    if (!group) { missingYearRows++; return; }

    const name = formatName(last, first);

    if (!byFlight[flightRaw]) byFlight[flightRaw] = {};
    if (!byFlight[flightRaw][group]) byFlight[flightRaw][group] = [];
    byFlight[flightRaw][group].push(name);

    if (!byFlightAll[flightRaw]) byFlightAll[flightRaw] = [];
    byFlightAll[flightRaw].push(name);

    inRows++;
  });

  // De-dup + sort
  Object.keys(byFlight).forEach(f => {
    Object.keys(byFlight[f]).forEach(g => {
      byFlight[f][g] = Array.from(new Set(byFlight[f][g]))
        .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    });
  });
  Object.keys(byFlightAll).forEach(f => {
    byFlightAll[f] = Array.from(new Set(byFlightAll[f]))
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  });

  Logger.log(`Rows processed: included=${inRows}, ignoredByFlight=${ignoredRows}, missingYearMap=${missingYearRows}`);
  Object.keys(byFlightAll).forEach(f => {
    const parts = AS_GROUP_TITLES.map(t => `${t}:${(byFlight[f][t] || []).length}`).join(', ');
    Logger.log(`[DATA] Flight "${f}" — total=${byFlightAll[f].length}; groups { ${parts} }`);
  });

  // Collect sections (page breaks)
  const items = form.getItems();
  Logger.log(`Total form items: ${items.length}`);
  const sections = [];
  items.forEach((it, idx) => {
    if (it.getType() === FormApp.ItemType.PAGE_BREAK) {
      const t = it.asPageBreakItem().getTitle() || '';
      sections.push({ index: idx, title: t });
    }
  });
  Logger.log(`Found ${sections.length} sections: ${sections.map(s => `"${s.title}"@${s.index}`).join(' | ')}`);

  // Helper: list items inside a section (for debug)
  function debugSectionContents(sectionIndex) {
    const start = sectionIndex + 1;
    const end = sections.find(s => s.index > sectionIndex)?.index ?? items.length;
    Logger.log(`  Items in section range [${start}, ${end - 1}]`);
    for (let i = start; i < end; i++) {
      const it = items[i];
      const type = it.getType();
      let title = '';
      try {
        if (type === FormApp.ItemType.CHECKBOX) title = it.asCheckboxItem().getTitle();
        else if (type === FormApp.ItemType.LIST) title = it.asListItem().getTitle();
        else if (type === FormApp.ItemType.MULTIPLE_CHOICE) title = it.asMultipleChoiceItem().getTitle();
        else if (type === FormApp.ItemType.TEXT) title = it.asTextItem().getTitle();
        else if (type === FormApp.ItemType.PARAGRAPH_TEXT) title = it.asParagraphTextItem().getTitle();
        else if (type === FormApp.ItemType.SCALE) title = it.asScaleItem().getTitle();
        else if (type === FormApp.ItemType.TIME) title = it.asTimeItem().getTitle();
        else if (type === FormApp.ItemType.DATE) title = it.asDateItem().getTitle();
        else if (type === FormApp.ItemType.DATETIME) title = it.asDateTimeItem().getTitle();
        else if (type === FormApp.ItemType.SECTION_HEADER) title = it.asSectionHeaderItem().getTitle();
      } catch (e) {}
      Logger.log(`    [${i}] type=${type} title="${title}"`);
    }
  }

  // Helper: find **CHECKBOX** item in a section by title (case-insensitive)
  function findCheckboxInSection(sectionIndex, titleWanted) {
    const start = sectionIndex + 1;
    const end = sections.find(s => s.index > sectionIndex)?.index ?? items.length;
    const wanted = norm(titleWanted);
    for (let i = start; i < end; i++) {
      const it = items[i];
      if (it.getType() === FormApp.ItemType.CHECKBOX) {
        const t = (it.asCheckboxItem().getTitle() || '');
        if (norm(t) === wanted) return it.asCheckboxItem();
      }
    }
    return null;
  }

  // 1) Update per-flight sections’ four **checkbox** questions
  const flightsInForm = new Set();
  sections.forEach(sec => {
    const secTitle = sec.title || '';
    const secTitleNorm = norm(secTitle);

    // Match section to any flight key present in our data, case-insensitive (e.g., "Alpha Flight")
    const matchingFlights = Object.keys(byFlightAll).filter(f => secTitleNorm.includes(norm(f)));
    if (matchingFlights.length > 0) {
      Logger.log(`[MATCH] Section "${secTitle}" matched flights: ${matchingFlights.join(', ')}`);
      debugSectionContents(sec.index);
    }

    matchingFlights.forEach(flightKey => {
      flightsInForm.add(flightKey);
      const groupData = byFlight[flightKey] || {};
      AS_GROUP_TITLES.forEach(groupTitle => {
        const cb = findCheckboxInSection(sec.index, groupTitle);
        if (!cb) {
          Logger.log(`(Skip) Checkbox "${groupTitle}" not found in section "${secTitle}" for flight "${flightKey}"`);
          return;
        }
        const names = groupData[groupTitle] || [];
        setCheckboxChoicesSafe(
          cb,
          names,
          `Per-flight "${groupTitle}" in section "${secTitle}" for flight "${flightKey}"`
        );
      });
    });
  });

  // 2) Update "Secondary Attendance" section — one checkbox per "<Flight> Flight" with ALL cadets (no AS split)
  const SECONDARY_SECTION_TITLE = 'Secondary/Crosstown Attendance';
  const secondary = sections.find(s => norm(s.title) === norm(SECONDARY_SECTION_TITLE));
  if (!secondary) {
    Logger.log(`Secondary Attendance section not found (title must be exactly "${SECONDARY_SECTION_TITLE}").`);
  } else {
    Logger.log(`[MATCH] Found "Secondary Attendance" section at index ${secondary.index}.`);
    debugSectionContents(secondary.index);
    Object.keys(byFlightAll).forEach(flightKey => {
      const questionTitle = `${flightKey} Flight`; // e.g., "Alpha Flight"
      const cb = findCheckboxInSection(secondary.index, questionTitle);
      if (!cb) {
        Logger.log(`(Skip) Secondary Attendance: checkbox "${questionTitle}" not found.`);
        return;
      }
      const names = byFlightAll[flightKey] || [];
      setCheckboxChoicesSafe(
        cb,
        names,
        `Secondary Attendance "${questionTitle}"`
      );
    });
  }

  // Optional diagnostics for flights present in data but no per-flight section
  Object.keys(byFlightAll).forEach(f => {
    if (!flightsInForm.has(f)) {
      Logger.log(`No per-flight section matched for "${f}".`);
    }
  });

  Logger.log('Sync complete.');
}

/***** OPTIONAL: dump a quick overview of all items (helps layout debugging) *****/
function debugDumpAllItems() {
  const form = getFormByIdOrUrl(FORM_ID_OR_URL);
  const items = form.getItems();
  Logger.log(`Form has ${items.length} items:`);
  items.forEach((it, i) => {
    let type = it.getType(), title = '';
    try {
      if (type === FormApp.ItemType.CHECKBOX) title = it.asCheckboxItem().getTitle();
      else if (type === FormApp.ItemType.LIST) title = it.asListItem().getTitle();
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE) title = it.asMultipleChoiceItem().getTitle();
      else if (type === FormApp.ItemType.PAGE_BREAK) title = it.asPageBreakItem().getTitle();
      else if (type === FormApp.ItemType.SECTION_HEADER) title = it.asSectionHeaderItem().getTitle();
      else if (type === FormApp.ItemType.TEXT) title = it.asTextItem().getTitle();
      else if (type === FormApp.ItemType.PARAGRAPH_TEXT) title = it.asParagraphTextItem().getTitle();
      else if (type === FormApp.ItemType.SCALE) title = it.asScaleItem().getTitle();
      else if (type === FormApp.ItemType.TIME) title = it.asTimeItem().getTitle();
      else if (type === FormApp.ItemType.DATE) title = it.asDateItem().getTitle();
      else if (type === FormApp.ItemType.DATETIME) title = it.asDateTimeItem().getTitle();
    } catch (e) {}
    Logger.log(`[${i}] type=${type} title="${title}"`);
  });
}

/***** TRIGGERS *****/
// Run once to install a daily trigger (runs at midnight by default)
function installTimeTrigger_OncePerDay() {
  // Remove existing triggers for cleanliness
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFormCheckboxesFromSheet')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Create a new time-based trigger that runs once per day
  ScriptApp.newTrigger('syncFormCheckboxesFromSheet')
    .timeBased()
    .everyDays(1)
    .atHour(0) // midnight; change to 1–23 for other times
    .create();
}

/***** HELPERS *****/
function getSpreadsheetByIdOrUrl(idOrUrl) {
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('Invalid Spreadsheet ID/URL: ' + idOrUrl);
  const id = m[0];
  DriveApp.getFileById(id); // permission check
  return SpreadsheetApp.openById(id);
}

function getFormByIdOrUrl(idOrUrl) {
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('Invalid Form ID/URL: ' + idOrUrl);
  const id = m[0];
  DriveApp.getFileById(id); // permission check
  return FormApp.openById(id);
}

function setCheckboxChoicesSafe(checkboxItem, values, contextLabel) {
  // values: array of strings; may be empty
  if (Array.isArray(values) && values.length > 0) {
    checkboxItem.setChoiceValues(values);
    Logger.log(`[OK] ${contextLabel}: wrote ${values.length} choices.`);
  } else {
    if (WRITE_PLACEHOLDER_WHEN_EMPTY) {
      checkboxItem.setChoiceValues([EMPTY_PLACEHOLDER_LABEL]);
      Logger.log(`[OK] ${contextLabel}: no cadets; wrote placeholder "${EMPTY_PLACEHOLDER_LABEL}".`);
    } else {
      Logger.log(`[SKIP] ${contextLabel}: no cadets; left existing choices unchanged.`);
    }
  }
}

// Name formatter
const formatName = (last, first) => `${last}, ${first}`;
```