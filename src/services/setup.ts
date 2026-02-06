// Setup service: idempotent provisioning of spreadsheets, sheets, and forms.

namespace SetupService {
  function extractFormIdFromUrl(url: string): string | null {
    if (!url) return null;
    // Common Forms URL formats:
    // - https://docs.google.com/forms/d/e/<ID>/viewform
    // - https://docs.google.com/forms/d/<ID>/edit
    const match = url.match(/\/forms\/d\/(?:e\/)?([a-zA-Z0-9_-]+)/);
    return match?.[1] || null;
  }

  function getFormDestinationSpreadsheetId(form: GoogleAppsScript.Forms.Form): string | null {
    try {
      const anyForm = form as any;
      const destinationType = anyForm.getDestinationType?.();
      const destinationId = anyForm.getDestinationId?.();
      if (destinationType === FormApp.DestinationType.SPREADSHEET && typeof destinationId === 'string') {
        return destinationId;
      }
      return null;
    } catch {
      return null;
    }
  }

  function ensureSpreadsheet(role: Types.WorkbookRole, name: string, propertyKey: string): Types.EnsureSpreadsheetResult {
    Log.info(`Ensuring spreadsheet for role=${role}`);
    const props = Config.scriptProperties();
    const existingId = props.getProperty(propertyKey);
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
    let created = false;

    if (existingId) {
      try {
        spreadsheet = SpreadsheetApp.openById(existingId);
        Log.info(`Found existing spreadsheet id=${existingId}`);
      } catch (err) {
        Log.warn(`Stored spreadsheet id invalid for key=${propertyKey}; creating new. Error: ${err}`);
      }
    }

    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.create(name);
      props.setProperty(propertyKey, spreadsheet.getId());
      created = true;
      Log.info(`Created spreadsheet name=${name} id=${spreadsheet.getId()}`);
    }

    return {
      role,
      id: spreadsheet.getId(),
      name: spreadsheet.getName(),
      created,
      url: spreadsheet.getUrl(),
    };
  }

  function ensureSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, schema: Types.TabSchema): Types.EnsureSheetResult {
    const { name, machineHeaders, displayHeaders } = schema;
    Log.info(`Ensuring sheet name=${name} in spreadsheet=${spreadsheet.getId()}`);
    const existingSheet = spreadsheet.getSheetByName(name);
    let sheet = existingSheet;
    let created = false;
    let headersApplied = false;

    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
      created = true;
      Log.info(`Created sheet name=${name}`);
    }

    if (sheet && machineHeaders && machineHeaders.length > 0) {
      const headerWidth = machineHeaders.length;
      const firstRow = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
      const secondRow = sheet.getRange(2, 1, 1, headerWidth).getValues()[0];
      const firstRowEmpty = firstRow.every((cell) => cell === '' || cell === null);
      const secondRowEmpty = secondRow.every((cell) => cell === '' || cell === null);

      if (firstRowEmpty) {
        sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
        headersApplied = true;
        Log.info(`Applied machine headers for ${name}`);
      }

      if (secondRowEmpty) {
        const display = displayHeaders && displayHeaders.length === machineHeaders.length ? displayHeaders : machineHeaders;
        sheet.getRange(2, 1, 1, headerWidth).setValues([display]);
        headersApplied = true;
        Log.info(`Applied display headers for ${name}`);
      }

      if (!headersApplied && firstRowEmpty === false && secondRowEmpty === false) {
        const nonEmptyFirst = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
        const nonEmptySecond = sheet.getRange(2, 1, 1, headerWidth).getValues()[0];
        const firstBlankCount = nonEmptyFirst.filter((v) => v === '' || v === null).length;
        const secondBlankCount = nonEmptySecond.filter((v) => v === '' || v === null).length;
        Log.warn(`Headers present for ${name}; blanks in row1=${firstBlankCount}/${headerWidth}, row2=${secondBlankCount}/${headerWidth}`);
      }

      // Ensure column count matches schema (add missing columns if schema grew)
      const maxCols = sheet.getMaxColumns();
      if (maxCols < headerWidth) {
        const addCount = headerWidth - maxCols;
        Log.info(`Adding ${addCount} missing columns to ${name} to match schema (${maxCols} -> ${headerWidth})`);
        sheet.insertColumnsAfter(maxCols, addCount);
        // Update headers in the new columns
        const display = displayHeaders && displayHeaders.length === machineHeaders.length ? displayHeaders : machineHeaders;
        sheet.getRange(1, maxCols + 1, 1, addCount).setValues([machineHeaders.slice(maxCols)]);
        sheet.getRange(2, maxCols + 1, 1, addCount).setValues([display.slice(maxCols)]);
        headersApplied = true;
      } else if (maxCols > headerWidth) {
        const deleteCount = maxCols - headerWidth;
        Log.info(`Deleting ${deleteCount} extra columns in ${name} (keeps ${headerWidth})`);
        sheet.deleteColumns(headerWidth + 1, deleteCount);
      }
    }

    return {
      spreadsheetId: spreadsheet.getId(),
      sheetName: name,
      created,
      headersApplied,
    };
  }

  function restoreMissingHeaders(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, schemas: Types.TabSchema[]) {
    schemas.forEach((schema) => {
      const { name, machineHeaders, displayHeaders } = schema;
      if (!machineHeaders || machineHeaders.length === 0) {
        Log.warn(`Schema for ${name} is missing machine headers; skipping header restoration.`);
        return;
      }
      const sheet = spreadsheet.getSheetByName(name);
      if (!sheet) {
        Log.warn(`Sheet ${name} not found while restoring headers; skipping.`);
        return;
      }
      const headerWidth = machineHeaders.length;
      const firstRow = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
      const firstRowEmpty = firstRow.every((cell) => cell === '' || cell === null);
      if (!firstRowEmpty) return;
      const display = displayHeaders && displayHeaders.length === headerWidth ? displayHeaders : machineHeaders;
      sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
      sheet.getRange(2, 1, 1, headerWidth).setValues([display]);
      Log.warn(`Restored missing machine/display headers on ${name}`);
    });
  }

  function resetSheetToSchema(sheet: GoogleAppsScript.Spreadsheet.Sheet, schema: Types.TabSchema) {
    const { machineHeaders, displayHeaders } = schema;
    if (!machineHeaders || machineHeaders.length === 0) {
      Log.warn(`Cannot reset sheet ${sheet.getName()} to schema: machine headers missing.`);
      return;
    }
    const headerWidth = machineHeaders.length;

    // Ensure column count matches schema width.
    const maxCols = sheet.getMaxColumns();
    if (maxCols < headerWidth) {
      sheet.insertColumnsAfter(maxCols, headerWidth - maxCols);
    } else if (maxCols > headerWidth) {
      sheet.deleteColumns(headerWidth + 1, maxCols - headerWidth);
    }

    // Clear all content and reapply headers.
    sheet.clear();
    sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
    const display = displayHeaders && displayHeaders.length === headerWidth ? displayHeaders : machineHeaders;
    sheet.getRange(2, 1, 1, headerWidth).setValues([display]);
  }

  function ensureTableForSheet(spreadsheetId: string, sheetName: string, tableId: string) {
    // Sheets advanced service may be disabled in some environments; skip gracefully if absent.
    if (typeof (globalThis as any).Sheets === 'undefined') {
      Log.warn(`Sheets advanced service unavailable; cannot create tables for ${sheetName}`);
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheet = Config.getBackendId() === spreadsheetId
        ? Config.getBackendSheet(sheetName)
        : Config.getFrontendId() === spreadsheetId
          ? Config.getFrontendSheet(sheetName)
          : (() => {
              const found = ss.getSheetByName(sheetName);
              if (!found) {
                const msg = `Sheet ${sheetName} missing; cannot ensure table ${tableId}.`;
                Log.error(msg);
                throw new Error(msg);
              }
              return found;
            })();
      const sheetId = sheet.getSheetId();

      const svc = (Sheets as any)?.Spreadsheets;
      if (!svc || !svc.batchUpdate) {
        Log.warn('Sheets advanced service unavailable; cannot create tables');
        return;
      }

      const headerRow = 2; // display headers live on row 2
      const headerValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colCount = headerValues.length;
      if (colCount === 0) return;
      const endColIndex = colCount; // zero-based exclusive
      const endRowIndex = Math.max(headerRow + 1, sheet.getLastRow());

      const columnProperties = headerValues.map((name, idx) => ({
        columnIndex: idx,
        columnName: String(name || `Column ${idx + 1}`),
      }));

      // Attempt to replace any existing table with the same id.
      try {
        svc.batchUpdate({ requests: [{ deleteTable: { tableId } } as any] }, spreadsheetId);
      } catch (err) {
        Log.info(`No existing table to delete for ${tableId} on ${sheetName}: ${err}`);
      }

      svc.batchUpdate(
        {
          requests: [
            {
              addTable: {
                table: {
                  name: tableId,
                  tableId,
                  range: {
                    sheetId,
                    startColumnIndex: 0,
                    endColumnIndex: endColIndex,
                    startRowIndex: headerRow - 1, // zero-based (row 2)
                    endRowIndex,
                  },
                  columnProperties,
                },
              },
            } as any,
          ],
        },
        spreadsheetId,
      );
      Log.info(`Ensured table ${tableId} on sheet ${sheetName}`);
    } catch (err) {
      Log.warn(`Unable to ensure table ${tableId} on sheet ${sheetName}: ${err}`);
    }
  }

  function isFrontendFormattingDisabled(): boolean {
    try {
      const prop = Config.scriptProperties().getProperty('DISABLE_FRONTEND_FORMATTING');
      return String(prop || '').toLowerCase() === 'true';
    } catch (err) {
      Log.warn(`Unable to read DISABLE_FRONTEND_FORMATTING property: ${err}`);
      return false;
    }
  }

  const RESPONSE_SHEET_REGEX = /^form responses?/i;

  function ensureResponseSheetName(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, desiredName: string): boolean {
    const current = spreadsheet.getSheetByName(desiredName);
    const candidates = spreadsheet.getSheets().filter((s) => RESPONSE_SHEET_REGEX.test(s.getName()));

    if (current) {
      // Desired sheet already present; do not delete other response sheets to avoid breaking links.
      return true;
    }

    if (candidates.length === 0) {
      const names = spreadsheet
        .getSheets()
        .map((s) => s.getName())
        .join(', ');
      Log.info(`No response sheet found to rename to ${desiredName} in spreadsheet=${spreadsheet.getId()} (sheet names: ${names})`);
      return false;
    }

    const primary = candidates[0];
    if (primary.getName() !== desiredName) {
      Log.info(`Renaming response sheet ${primary.getName()} -> ${desiredName}`);
      primary.setName(desiredName);
    }

    // Leave any additional Form Responses sheets untouched to avoid deleting linked sheets; log for awareness.
    candidates.slice(1).forEach((s) => {
      if (s.getName() !== desiredName) {
        Log.warn(`Additional response sheet present (${s.getName()}); leaving as-is to avoid unlinking forms.`);
      }
    });
    return true;
  }

  function copySheetToArchive(
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
    source: GoogleAppsScript.Spreadsheet.Sheet,
    archivePrefix: string,
  ): GoogleAppsScript.Spreadsheet.Sheet | null {
    const archiveName = `${archivePrefix}${source.getName()}`.trim();

    // Replace only the canonical archive sheet for this source; leave any user-renamed archives intact.
    const existingArchive = ss.getSheetByName(archiveName);
    if (existingArchive) {
      try {
        ss.deleteSheet(existingArchive);
      } catch (err) {
        Log.warn(`Unable to delete existing archive sheet ${archiveName}: ${err}`);
        return null;
      }
    }

    let archived: GoogleAppsScript.Spreadsheet.Sheet;
    try {
      archived = source.copyTo(ss);
    } catch (err) {
      Log.warn(`Unable to copy sheet ${source.getName()} to archive ${archiveName}: ${err}`);
      return null;
    }

    try {
      archived.setName(archiveName);
    } catch (err) {
      Log.warn(`Unable to rename archive copy to ${archiveName}: ${err}`);
    }

    // Sever links: strip formulas, named ranges, and protections.
    const range = archived.getDataRange();
    range.copyTo(range, { contentsOnly: true });
    archived.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach((p) => p.remove());
    archived.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
    try {
      const protection = archived.protect().setDescription(`${archiveName} (locked)`);
      protection.setWarningOnly(false);
      try {
        protection.removeEditors(protection.getEditors());
      } catch (err) {
        Log.warn(`Unable to remove editors from ${archiveName}: ${err}`);
      }
      if (protection.canDomainEdit && protection.canDomainEdit()) {
        try {
          protection.setDomainEdit(false);
        } catch (err) {
          Log.warn(`Unable to disable domain edit on ${archiveName}: ${err}`);
        }
      }
    } catch (err) {
      Log.warn(`Unable to protect archive sheet ${archiveName}: ${err}`);
    }

    ss.setActiveSheet(archived);
    ss.moveActiveSheet(ss.getSheets().length);

    return archived;
  }

  function archiveAndResetSheets(
    spreadsheetId: string,
    schemas: Types.TabSchema[],
    names: string[],
    archivePrefix = 'Archive ',
  ) {
    if (!spreadsheetId) {
      Log.warn('No spreadsheetId provided to archiveAndResetSheets; skipping.');
      return;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);

    names.forEach((name) => {
      const schema = schemas.find((s) => s.name === name);
      if (!schema || !schema.machineHeaders) {
        Log.warn(`Schema missing for ${name}; skipping archive/reset.`);
        return;
      }
      const sheet = ss.getSheetByName(name);
      if (!sheet) {
        Log.warn(`Sheet ${name} not found in spreadsheet=${spreadsheetId}; skipping archive/reset.`);
        return;
      }

      copySheetToArchive(ss, sheet, archivePrefix);

      resetSheetToSchema(sheet, schema);
    });
  }

  function restoreFromArchiveSheets(
    spreadsheetId: string,
    schemas: Types.TabSchema[],
    names: string[],
    archivePrefix = 'Archive ',
  ) {
    if (!spreadsheetId) {
      Log.warn('No spreadsheetId provided to restoreFromArchiveSheets; skipping.');
      return;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);

    names.forEach((name) => {
      const schema = schemas.find((s) => s.name === name);
      if (!schema || !schema.machineHeaders) {
        Log.warn(`Schema missing for ${name}; skipping restore from archive.`);
        return;
      }
      let target = ss.getSheetByName(name);
        const archive = ss.getSheetByName(`${archivePrefix}${name}`);
        if (!archive) {
          Log.warn(`No archive sheet found for ${name}; skipping restore.`);
          return;
        }
      if (!target) {
        target = ss.insertSheet(name);
      }

      const values = archive.getDataRange().getValues();
      const width = Math.max(schema.machineHeaders.length, values[0]?.length || 0);

      const maxCols = target.getMaxColumns();
      if (maxCols < width) target.insertColumnsAfter(maxCols, width - maxCols);
      if (maxCols > width) target.deleteColumns(width + 1, maxCols - width);

      target.clear();
      if (values.length && values[0].length) {
        target.getRange(1, 1, values.length, values[0].length).setValues(values);
      }
    });
  }

  function ensureResponseSheetNameWithRetry(spreadsheetId: string, desiredName: string, retries = 3, delayMs = 500) {
    for (let attempt = 0; attempt < retries; attempt++) {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const ok = ensureResponseSheetName(ss, desiredName);
      if (ok) return;
      Utilities.sleep(delayMs);
    }
    Log.warn(`Unable to find response sheet for ${desiredName} after ${retries} attempts in spreadsheet=${spreadsheetId}; skipping placeholder.`);
  }

  // Minimal response sheet handling: set destination once, wait, rename the linked sheet (or first response-ish) to desiredName.
  function ensureResponseSheetForForm(form: GoogleAppsScript.Forms.Form, desiredName: string, spreadsheetId: string) {
    const currentDest = getFormDestinationSpreadsheetId(form);
    if (!currentDest || currentDest !== spreadsheetId) {
      try {
        form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
      } catch (err) {
        Log.warn(`Unable to set destination for formId=${form.getId()} to spreadsheet=${spreadsheetId}: ${err}`);
        return;
      }
    }

    // Poll for the linked response sheet to appear; Forms can lag before creating it.
    for (let attempt = 0; attempt < 5; attempt++) {
      if (attempt > 0) Utilities.sleep(500);
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheets = ss.getSheets();

      // Prefer the sheet actually linked to this form.
      const linked = sheets.filter((s) => {
        try {
          const url = s.getFormUrl?.();
          return url && url.includes(form.getId());
        } catch {
          return false;
        }
      });

      const candidates = linked.length
        ? linked
        : sheets.filter((s) => RESPONSE_SHEET_REGEX.test(s.getName()) || s.getName() === desiredName);

      const target = ss.getSheetByName(desiredName) || candidates[0];
      if (!target) continue;

      if (target.getName() !== desiredName) {
        try {
          target.setName(desiredName);
        } catch (err) {
          Log.warn(`Unable to rename response sheet ${target.getName()} -> ${desiredName}: ${err}`);
        }
      }
      return;
    }

    // As a last resort, create a placeholder sheet so downstream setup steps do not crash.
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const existing = ss.getSheetByName(desiredName);
      if (!existing) {
        ss.insertSheet(desiredName);
        Log.warn(`Created placeholder response sheet ${desiredName} because none were found for formId=${form.getId()}.`);
      }
    } catch (err) {
      Log.warn(`Unable to create placeholder response sheet ${desiredName}: ${err}`);
    }
  }

  function slimAttendanceResponseSheet() {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
    try {
      sheet = Config.getBackendSheet(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
    } catch (err) {
      Log.warn(`Attendance response sheet missing; skipping slim. Error: ${err}`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));
    const dataRows = Math.max(0, lastRow - 1);
    const startRow = 2;

    // Headers that should not exist (legacy items we removed from the form).
    const bannedHeaders = new Set(['Submitted By Email']);

    // Group columns by header; keep the first occurrence, merge data from later duplicates.
    const indicesByHeader = new Map<string, number[]>();
    headers.forEach((h, idx) => {
      const key = h.trim();
      if (!key) return;
      const arr = indicesByHeader.get(key) || [];
      arr.push(idx + 1); // 1-based
      indicesByHeader.set(key, arr);
    });

    indicesByHeader.forEach((cols, header) => {
      if (cols.length <= 1) return;
      if (bannedHeaders.has(header)) {
        // Delete all occurrences of banned headers.
        cols
          .slice()
          .sort((a, b) => b - a)
          .forEach((col) => {
            const currentMax = sheet.getMaxColumns();
            if (col > currentMax) return;
            try {
              sheet.deleteColumn(col);
            } catch (err) {
              try {
                sheet.hideColumn(sheet.getRange(1, col));
              } catch (err2) {
                Log.warn(
                  `Unable to delete or hide banned header '${header}' column ${col} in ${Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET}: ${err}; hide failed: ${err2}`,
                );
              }
            }
          });
        return;
      }
      // Merge all duplicate columns' data together (deduping values) and write the merged value into every duplicate column.
      if (dataRows > 0) {
        const colValues = cols.map((col) => sheet.getRange(startRow, col, dataRows, 1).getValues());
        const merged: string[][] = Array.from({ length: dataRows }, () => ['']);

        for (let r = 0; r < dataRows; r++) {
          const seen = new Set<string>();
          const parts: string[] = [];
          colValues.forEach((vals) => {
            const raw = String(vals[r][0] || '').trim();
            if (!raw) return;
            raw.split('|').forEach((p) => {
              const part = p.trim();
              if (!part) return;
              if (seen.has(part)) return;
              seen.add(part);
              parts.push(part);
            });
          });
          merged[r][0] = parts.join(' | ');
        }

        cols.forEach((col) => {
          sheet.getRange(startRow, col, dataRows, 1).setValues(merged);
        });
      }

      // Attempt to delete all duplicates; the column that cannot be deleted (form-linked) will remain.
      let survivor: number | null = null;
      const sorted = cols.slice().sort((a, b) => b - a); // delete right-to-left to reduce shifting issues
      sorted.forEach((col, idx) => {
        // If we have no survivor yet and this is the last column, keep it to guarantee one remains.
        if (survivor === null && idx === sorted.length - 1) {
          survivor = col;
          return;
        }

        const currentMax = sheet.getMaxColumns();
        if (col > currentMax) return;
        try {
          sheet.deleteColumn(col);
        } catch (err) {
          // Likely the form-linked column; keep it but continue pruning other duplicates.
          survivor = survivor ?? col;
        }
      });
    });
  }

  function pruneAttendanceResponseColumnsExplicit() {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet;
    try {
      sheet = Config.getBackendSheet(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
    } catch (err) {
      Log.warn(`Attendance response sheet missing; skipping prune. Error: ${err}`);
      return;
    }

    // First merge any duplicate data so deletes do not drop content.
    slimAttendanceResponseSheet();

    // Re-run pruning a few times to tolerate column shifting or prior delete failures.
    const bannedHeaders = new Set(['Submitted By Email']);
    for (let attempt = 0; attempt < 5; attempt++) {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));

      const indicesByHeader = new Map<string, number[]>();
      headers.forEach((h, idx) => {
        const key = h.trim();
        if (!key) return;
        const arr = indicesByHeader.get(key) || [];
        arr.push(idx + 1);
        indicesByHeader.set(key, arr);
      });

      let changed = false;
      let sawDuplicate = false;

      indicesByHeader.forEach((cols, header) => {
        if (cols.length <= 1) return;
        if (bannedHeaders.has(header)) {
          cols
            .slice()
            .sort((a, b) => b - a)
            .forEach((col) => {
              const currentMax = sheet.getMaxColumns();
              if (col > currentMax) return;
              try {
                sheet.deleteColumn(col);
                changed = true;
              } catch (err) {
                try {
                  sheet.hideColumn(sheet.getRange(1, col));
                  changed = true;
                } catch (err2) {
                  Log.warn(
                    `Unable to delete or hide banned header '${header}' column ${col} in ${Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET}: ${err}; hide failed: ${err2}`,
                  );
                }
              }
            });
          return;
        }
        sawDuplicate = true;

        let kept = false;
        const sorted = cols.slice().sort((a, b) => b - a);
        sorted.forEach((col, idx) => {
          const remaining = sorted.length - idx;

          // Always leave at least one column untouched (last remaining if none kept yet).
          if (!kept && remaining === 1) {
            kept = true;
            return;
          }

          const currentMax = sheet.getMaxColumns();
          if (col > currentMax) return;
          try {
            sheet.deleteColumn(col);
            changed = true;
          } catch (err) {
            // Likely form-linked; keep it and continue.
            kept = true;
          }
        });
      });

      if (!sawDuplicate || !changed) break;
    }

    normalizeAttendanceBackendHeaders();
  }

  function pruneExcusalsResponseColumnsExplicit(verbose = true) {
    const vlog = (msg: string) => {
      if (verbose) Log.info(`[excusal-prune] ${msg}`);
    };
    let sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
    try {
      sheet = Config.getBackendSheet(Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET);
    } catch (err) {
      Log.warn(`Excusals response sheet missing; skipping prune. Error: ${err}`);
      return;
    }

    for (let attempt = 0; attempt < 3; attempt++) {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      let changed = false;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));
      if (verbose) {
        vlog(`Attempt ${attempt + 1}: lastCol=${lastCol}`);
        headers.forEach((h, idx) => vlog(`  Col ${idx + 1}: '${h}'`));
      }

      const normalizeHeader = (raw: string) => {
        const orig = raw || '';
        const h = orig.trim().toLowerCase();
        if (
          h === 'event' ||
          /^event\s*\d+$/i.test(orig) ||
          /^event\s*\(\d+\)$/i.test(orig) ||
          /^event\b.*other/i.test(orig) ||
          h.startsWith('other event')
        ) {
          return 'event';
        }
        return orig.trim();
      };

      const indicesByHeader = new Map<string, number[]>();
      headers.forEach((h, idx) => {
        const key = normalizeHeader(h);
        if (!key) return;
        const arr = indicesByHeader.get(key) || [];
        arr.push(idx + 1);
        indicesByHeader.set(key, arr);
      });

      indicesByHeader.forEach((cols, header) => {
        if (cols.length <= 1) return;
        const sorted = cols.slice().sort((a, b) => a - b); // Sort ascending by column position
        vlog(`Header '${header}' has ${sorted.length} columns: ${sorted.join(', ')}`);

        // Keep the LAST (rightmost) column as the survivor - it's typically the newest form-linked column
        const survivorCol = sorted[sorted.length - 1];
        const columnsToDelete = sorted.slice(0, -1); // All except the last one
        
        vlog(`  Survivor: column ${survivorCol}, will delete: ${columnsToDelete.join(', ')}`);

        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          // Merge data from all columns into the survivor
          const survivorData = sheet.getRange(2, survivorCol, lastRow - 1, 1).getValues();
          
          // For each column that will be deleted, merge its data into survivor
          columnsToDelete.forEach((col) => {
            const currentMax = sheet.getMaxColumns();
            if (col > currentMax) return;
            
            try {
              const sourceData = sheet.getRange(2, col, lastRow - 1, 1).getValues();
              
              // Merge: copy non-empty cells from source to survivor where survivor is empty
              for (let i = 0; i < sourceData.length; i++) {
                const sourceValue = sourceData[i][0];
                const survivorValue = survivorData[i][0];
                
                if ((!survivorValue || String(survivorValue).trim() === '') && sourceValue && String(sourceValue).trim() !== '') {
                  survivorData[i][0] = sourceValue;
                  vlog(`  Merged data from col ${col} row ${i + 2} to survivor col ${survivorCol}`);
                }
              }
            } catch (err) {
              Log.warn(`Unable to read data from column ${col} for merging: ${err}`);
            }
          });
          
          // Write merged data back to survivor
          try {
            sheet.getRange(2, survivorCol, lastRow - 1, 1).setValues(survivorData);
            vlog(`  Wrote merged data to survivor column ${survivorCol}`);
          } catch (err) {
            Log.warn(`Unable to write merged data to survivor column ${survivorCol}: ${err}`);
          }
        }

        // Delete columns in reverse order to preserve indices during deletion
        columnsToDelete.reverse().forEach((col) => {
          const currentMax = sheet.getMaxColumns();
          if (col > currentMax) return;
          
          try {
            sheet.deleteColumn(col);
            changed = true;
            vlog(`  Deleted column ${col} for header '${header}'`);
          } catch (err) {
            Log.warn(`Unable to delete column ${col} for '${header}' in ${Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET}: ${err}`);
          }
        });
      });

      if (!changed) break;
    }
  }

  // Verbose debug entrypoint to inspect and prune excusals Event columns; callable from Apps Script.
  export function debugExcusalsResponseColumnsVerbose() {
    pruneExcusalsResponseColumnsExplicit(true);
    try {
      const sheet = Config.getBackendSheet(Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET);
      const lastCol = sheet.getLastColumn();
      const headers = lastCol
        ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''))
        : [];
      Log.info(`[excusal-prune] Final headers (${headers.length} cols): ${headers.map((h, i) => `${i + 1}:'${h}'`).join(', ')}`);
      return headers;
    } catch (err) {
      Log.warn(`[excusal-prune] Unable to log final headers: ${err}`);
      return [];
    }
  }

  /**
   * Process any existing rows in the Excusals Form Responses sheet that haven't been
   * inserted into Excusals Backend yet (pre-online submissions/backfill).
   */
  export function processExcusalsFormBacklog() {
    try {
      // Normalize response headers first to ensure single 'Event' column.
      pruneExcusalsResponseColumnsExplicit(false);

      const respSheet = Config.getBackendSheet(Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET);
      const lastCol = respSheet.getLastColumn();
      const lastRow = respSheet.getLastRow();
      if (lastCol === 0 || lastRow < 2) {
        Log.info('No excusals form responses found; nothing to backfill.');
        return;
      }

      const headers = respSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
      const headerIndex = (name: string, normalize?: (s: string) => string) => {
        const n = normalize ? headers.map((h) => normalize(h)) : headers;
        return n.indexOf(name);
      };

      // Normalize 'Event' header variants like 'Event (Other)', 'Event 2', etc.
      const normalizeEventHeader = (raw: string) => {
        const h = (raw || '').trim().toLowerCase();
        if (h === 'event' || /^event\s*\d+$/i.test(raw) || /^event\s*\(\d+\)$/i.test(raw) || /^event\b.*other/i.test(raw)) {
          return 'event';
        }
        return (raw || '').trim();
      };

      // Normalize 'Email' header variants like 'Email Address', 'Email', etc.
      const normalizeEmailHeader = (raw: string) => {
        const h = (raw || '').trim().toLowerCase();
        if (h === 'email' || h.startsWith('email')) {
          return 'email';
        }
        return (raw || '').trim();
      };

      const tsIdx = headerIndex('Timestamp');
      const emailIdx = headers.map((h) => normalizeEmailHeader(h)).indexOf('email');
      const lastIdx = headerIndex('Last Name');
      const firstIdx = headerIndex('First Name');
      const reasonIdx = headerIndex('Reason');
      const eventIdx = headers.map((h) => normalizeEventHeader(h)).indexOf('event');

      if (eventIdx < 0 || emailIdx < 0) {
        Log.warn('Excusals responses missing required headers (Event/Email); cannot backfill.');
        return;
      }

      // Build existing key set from Excusals Backend to avoid duplicates.
      const backendSheet = Config.getBackendSheet('Excusals Backend');
      const backendLastCol = backendSheet.getLastColumn();
      const backendHeaders = backendLastCol
        ? backendSheet.getRange(1, 1, 1, backendLastCol).getValues()[0].map((h) => String(h || '').trim().toLowerCase())
        : [];
      const backend = SheetUtils.readTable(backendSheet);
      const emailColB = backendHeaders.indexOf('email');
      const eventColB = backendHeaders.indexOf('event');
      const submittedColB = backendHeaders.indexOf('submitted_at');
      const existingKeys = new Set<string>();
      backend.rows.forEach((row) => {
        const e = String(row['email'] || '').toLowerCase().trim();
        const ev = String(row['event'] || '').trim();
        const ts = String(row['submitted_at'] || '').trim();
        if (e && ev && ts) existingKeys.add(`${e}|${ev}|${ts}`);
      });

      const toAppend: Record<string, any>[] = [];
      const toSync: Record<string, any>[] = [];

      // Helper: lookup cadet by email from Directory Backend
      const lookupCadetByEmail = (addr: string) => {
        const backendId = Config.getBackendId();
        if (!backendId || !addr) return null;
        const directorySheet = SheetUtils.getSheet(backendId, 'Directory Backend');
        if (!directorySheet) return null;
        const data = SheetUtils.readTable(directorySheet);
        const lower = addr.toLowerCase();
        return data.rows.find((r) => String(r['email'] || '').toLowerCase() === lower) || null;
      };

      const values = respSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      values.forEach((row) => {
        const email = emailIdx >= 0 ? String(row[emailIdx] || '').trim() : '';
        const lastName = lastIdx >= 0 ? String(row[lastIdx] || '').trim() : '';
        const firstName = firstIdx >= 0 ? String(row[firstIdx] || '').trim() : '';
        const reason = reasonIdx >= 0 ? String(row[reasonIdx] || '').trim() : '';
        const tsVal = tsIdx >= 0 ? row[tsIdx] : '';
        const submittedAt = (() => {
          try { return new Date(tsVal).toISOString(); } catch { return new Date().toISOString(); }
        })();

        const eventsRaw = String(row[eventIdx] || '').trim();
        const events = eventsRaw
          .split(',')
          .map((ev) => ev.trim())
          .filter(Boolean);

        if (!email || !events.length) return;

        const cadet = lookupCadetByEmail(email);

        events.forEach((eventName) => {
          const key = `${email.toLowerCase()}|${eventName}|${submittedAt}`;
          if (existingKeys.has(key)) return;

          const requestId = `exc-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
          const backendRow = {
            request_id: requestId,
            event: eventName,
            email,
            last_name: cadet?.last_name || lastName,
            first_name: cadet?.first_name || firstName,
            flight: cadet?.flight || '',
            squadron: cadet?.squadron || '',
            status: 'Pending',
            decision: '',
            decided_by: '',
            decided_at: '',
            attendance_effect: '',
            submitted_at: submittedAt,
            last_updated_at: submittedAt,
            notes: reason,
          };

          toAppend.push(backendRow);
          toSync.push(backendRow);
        });
      });

      if (!toAppend.length) {
        Log.info('No unprocessed excusals responses found.');
        return;
      }

      SheetUtils.appendRows(backendSheet, toAppend);
      toSync.forEach((row) => {
        ExcusalsService.notifySquadronCommanderOfNewExcusal(row);
        ExcusalsService.syncExcusalToManagementPanel(row);
      });

      Log.info(`Processed excusals form backlog: appended ${toAppend.length} event rows.`);
    } catch (err) {
      Log.warn(`Failed to process excusals form backlog: ${err}`);
    }
  }

  /**
   * Debug helper: shows what columns exist in the Attendance Form Response sheet.
   */
  export function debugAttendanceResponseSheet() {
    try {
      const respSheet = Config.getBackendSheet(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
      const lastCol = respSheet.getLastColumn();
      const headers = respSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h, idx) => `${idx}: "${String(h || '').trim()}"`);
      Log.info(`Attendance Response Sheet has ${lastCol} columns:`);
      headers.forEach((h) => Log.info(`  ${h}`));
      
      // Show first data row
      const lastRow = respSheet.getLastRow();
      if (lastRow >= 2) {
        const firstDataRow = respSheet.getRange(2, 1, 1, lastCol).getValues()[0];
        Log.info('First data row:');
        firstDataRow.forEach((val, idx) => {
          const v = String(val || '').trim();
          if (v) Log.info(`  Col ${idx}: "${v.substring(0, 100)}"`);
        });
      }
      return headers;
    } catch (err) {
      Log.warn(`debugAttendanceResponseSheet failed: ${err}`);
      return [];
    }
  }

  /**
   * Process any existing rows in the Attendance Form Responses sheet that haven't been
   * inserted into Attendance Backend yet (backfill for new multi-category form structure).
   */
  export function processAttendanceFormBacklog() {
    try {
      const respSheet = Config.getBackendSheet(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
      const lastCol = respSheet.getLastColumn();
      const lastRow = respSheet.getLastRow();
      if (lastCol === 0 || lastRow < 2) {
        Log.info('No attendance form responses found; nothing to backfill.');
        return;
      }

      const headers = respSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());

      // Find column indices
      const tsIdx = headers.indexOf('Timestamp');
      const emailIdx = headers.findIndex((h) => h.toLowerCase().includes('email'));
      const nameIdx = headers.findIndex((h) => h.toLowerCase() === 'name');
      const eventTypeIdx = headers.findIndex((h) => h.toLowerCase() === 'event type');
      const flightCrosstownIdx = headers.findIndex((h) => h.toLowerCase().includes('flight') && h.toLowerCase().includes('crosstown'));

      // Find all "Select Event" columns (but exclude any that contain "Cadets" to avoid confusion)
      const selectEventIndices: number[] = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase();
        if (lower.includes('select event') && !lower.includes('cadet')) {
          selectEventIndices.push(idx);
        }
      });

      // Find all cadet checkbox columns (pattern: "Cadets (...) AS AS... (...)")
      const cadetColumnIndices: Array<{ idx: number; title: string }> = [];
      headers.forEach((h, idx) => {
        if (h.toLowerCase().includes('cadets') && h.toLowerCase().includes('as ')) {
          cadetColumnIndices.push({ idx, title: h });
        }
      });

      if (selectEventIndices.length === 0) {
        Log.warn('Attendance responses missing "Select Event" columns; cannot backfill.');
        return;
      }

      Log.info(`Found ${selectEventIndices.length} "Select Event" columns and ${cadetColumnIndices.length} cadet columns`);
      Log.info(`Email col: ${emailIdx}, Name col: ${nameIdx}, Flight col: ${flightCrosstownIdx}`);

      // Build existing key set from Attendance Backend to avoid duplicates
      const backendSheet = Config.getBackendSheet('Attendance Backend');
      const backend = SheetUtils.readTable(backendSheet);
      const existingKeys = new Set<string>();
      backend.rows.forEach((row) => {
        const e = String(row['email'] || '').toLowerCase().trim();
        const ev = String(row['event'] || '').trim();
        const ts = String(row['submitted_at'] || '').trim();
        if (e && ev && ts) existingKeys.add(`${e}|${ev}|${ts}`);
      });

      // Helper: normalize cadet list from response value
      const normalizeCadetList = (val: any): string[] => {
        if (!val) return [];
        const s = String(val).trim();
        if (!s) return [];
        // Google Forms checkbox responses are semicolon-separated: "Last, First; Last, First; ..."
        return s.split(';').map((n) => n.trim()).filter(Boolean);
      };

      // Helper: lookup cadet by email for flight/squadron info
      const lookupCadetByEmail = (addr: string) => {
        const backendId = Config.getBackendId();
        if (!backendId || !addr) return null;
        const directorySheet = SheetUtils.getSheet(backendId, 'Directory Backend');
        if (!directorySheet) return null;
        const data = SheetUtils.readTable(directorySheet);
        const lower = addr.toLowerCase();
        return data.rows.find((r) => String(r['email'] || '').toLowerCase() === lower) || null;
      };

      const toAppend: Record<string, any>[] = [];
      const logEntries: Array<{ event: string; attendance_type: string; cadets: string }> = [];

      const values = respSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      let rowNum = 2; // Starting row number for logging
      values.forEach((row) => {
        const email = emailIdx >= 0 ? String(row[emailIdx] || '').trim() : '';
        const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
        const flightCrosstown = flightCrosstownIdx >= 0 ? String(row[flightCrosstownIdx] || '').trim() : '';
        const tsVal = tsIdx >= 0 ? row[tsIdx] : '';
        const submittedAt = (() => {
          try { return new Date(tsVal).toISOString(); } catch { return new Date().toISOString(); }
        })();

        // Collect all selected events from "Select Event" columns
        const selectedEvents: string[] = [];
        selectEventIndices.forEach((idx) => {
          const val = String(row[idx] || '').trim();
          if (val) {
            // Events should be single values, not semicolon-separated lists
            // If we see semicolons, it's likely cadet names incorrectly in an event column
            if (val.includes(';')) {
              Log.warn(`Row ${rowNum}: "Select Event" column ${idx} contains suspicious value (looks like cadet list): "${val.substring(0, 50)}"`);
            } else {
              // Single event name
              if (!selectedEvents.includes(val)) {
                selectedEvents.push(val);
              }
            }
          }
        });

        if (selectedEvents.length === 0) {
          rowNum++;
          return;
        }

        // For each selected event, determine event type and collect relevant cadets
        selectedEvents.forEach((eventName) => {
          const key = `${email.toLowerCase()}|${eventName}|${submittedAt}`;
          if (existingKeys.has(key)) return;

          // Determine event type from event name pattern
          let eventType = '';
          if (eventName.includes('LLAB') || eventName.includes('TW-')) {
            if (eventName.includes('POC Third Hour')) {
              eventType = 'POC';
            } else if (eventName.includes('Secondary')) {
              eventType = 'Secondary';
            } else if (eventName.includes('LLAB')) {
              eventType = 'LLAB';
            } else {
              eventType = 'Mando';
            }
          } else {
            eventType = 'Other';
          }

          // Collect relevant cadet selections for this event type
          const relevantCadets: string[] = [];
          cadetColumnIndices.forEach(({ idx, title }) => {
            const titleLower = title.toLowerCase();
            const matches =
              (eventType === 'Mando' && titleLower.includes('(mando)')) ||
              (eventType === 'LLAB' && titleLower.includes('(llab)')) ||
              (eventType === 'POC' && titleLower.includes('(poc)')) ||
              (eventType === 'Secondary' && titleLower.includes('(secondary)')) ||
              (eventType === 'Other' && titleLower.includes('(all)'));

            if (matches) {
              const cadets = normalizeCadetList(row[idx]);
              cadets.forEach((c) => {
                if (!relevantCadets.includes(c)) {
                  relevantCadets.push(c);
                }
              });
            }
          });

          const cadetField = relevantCadets.join('; ');
          const cadet = lookupCadetByEmail(email);
          const submissionId = `att-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;

          // Determine flight field based on event type
          let flightValue = '';
          if (eventType === 'Mando' || eventType === 'LLAB') {
            // For Mando/LLAB, use the selected flight (Alpha-Foxtrot, Trine, Valparaiso)
            flightValue = flightCrosstown || cadet?.flight || '';
          } else if (eventType === 'Secondary') {
            flightValue = 'Secondary';
          } else if (eventType === 'POC') {
            flightValue = 'POC Third Hour';
          } else if (eventType === 'Other') {
            flightValue = 'Other';
          }

          const backendRow = {
            submission_id: submissionId,
            submitted_at: submittedAt,
            event: eventName,
            attendance_type: 'P',
            email,
            name,
            flight: flightValue,
            cadets: cadetField,
          };

          toAppend.push(backendRow);
          logEntries.push({
            event: eventName,
            attendance_type: 'P',
            cadets: cadetField,
          });

          // Log first few entries for debugging
          if (toAppend.length <= 3) {
            Log.info(`Row ${rowNum} Event: "${eventName}", Cadets: "${cadetField.substring(0, 50)}", Email: ${email}, Name: ${name}`);
          }
        });
        
        rowNum++;
      });

      if (!toAppend.length) {
        Log.info('No unprocessed attendance responses found.');
        return;
      }

      SheetUtils.appendRows(backendSheet, toAppend);
      
      // Apply each log entry to update the attendance matrix
      logEntries.forEach((entry) => {
        AttendanceService.applyAttendanceLogEntry(entry);
      });

      applyAttendanceBackendFormatting();

      Log.info(`Processed attendance form backlog: appended ${toAppend.length} event rows, updated matrix.`);
    } catch (err) {
      Log.warn(`Failed to process attendance form backlog: ${err}`);
    }
  }

  function normalizeAttendanceBackendHeaders() {
    const sheet = Config.getBackendSheet('Attendance Backend');

    const attendanceSchema = Schemas.BACKEND_TABS.find((t) => t.name === 'Attendance Backend');
    if (!attendanceSchema?.machineHeaders || !attendanceSchema?.displayHeaders) {
      Log.warn('Attendance Backend schema is missing machine or display headers; skipping normalizing Attendance Backend headers.');
      return;
    }
    const targetHeaders = attendanceSchema.machineHeaders;
    const displayHeaders = attendanceSchema.displayHeaders;

    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastCol = Math.max(sheet.getLastColumn(), targetHeaders.length);
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    const sourceHeaders = (values[0] || []).map((h) => String(h || '').trim());
    const sourceLookup = new Map<string, number>();
    sourceHeaders.forEach((h, idx) => {
      const key = h.toLowerCase();
      if (key) sourceLookup.set(key, idx);
    });

    const altKeys: Record<string, string[]> = {
      submission_id: ['submission id'],
      submitted_at: ['submitted at', 'timestamp', 'submission time'],
      event: ['event'],
      email: ['email', 'email address', 'submitted by email'],
      name: ['name', 'submitted by name'],
      flight: ['flight', 'flight / crosstown (mando)', 'flight (mando pt)', 'flight / crosstown', 'flight / crosstown (llab)', 'flight (llab)'],
      cadets: ['cadets', 'cadet selections', 'cadet list'],
    };

    const headerMatches = targetHeaders.map((h) => {
      const key = h.toLowerCase();
      if (sourceLookup.has(key)) return sourceLookup.get(key)!;
      const alts = altKeys[h] || [];
      for (const alt of alts) {
        const altIdx = sourceLookup.get(alt.toLowerCase());
        if (altIdx !== undefined) return altIdx;
      }
      return -1;
    });

    // Detect if row 2 is a display/header row to skip when rebuilding data.
    const displayRowMatches = (values[1] || []).every((cell: any, idx: number) => {
      const expected = displayHeaders[idx] || '';
      return String(cell || '').trim().toLowerCase() === expected.toLowerCase();
    });
    const dataStart = displayRowMatches ? 3 : 2;
    const dataRows: any[][] = [];
    for (let r = dataStart - 1; r < lastRow; r++) {
      const row = values[r] || [];
      const out = targetHeaders.map((_, idx) => {
        const srcIdx = headerMatches[idx];
        return srcIdx >= 0 ? row[srcIdx] || '' : '';
      });
      if (out.some((v) => v !== '')) dataRows.push(out);
    }

    sheet.clear();
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    sheet.getRange(2, 1, 1, targetHeaders.length).setValues([displayHeaders]);
    if (dataRows.length) {
      sheet.getRange(3, 1, dataRows.length, targetHeaders.length).setValues(dataRows);
    }
  }

    function applyAttendanceBackendFormatting() {
        const sheet = Config.getBackendSheet('Attendance Backend');

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
      const flightCol = headers.indexOf('flight') + 1;
      if (flightCol <= 0) return;

      const startRow = 3; // data starts after header rows
      const lastRow = Math.max(startRow, sheet.getLastRow());
      const numRows = Math.max(1, lastRow - startRow + 1);
      const dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());

      const columnToLetter = (col: number) => {
        let temp = '';
        let n = col;
        while (n > 0) {
          const rem = (n - 1) % 26;
          temp = String.fromCharCode(65 + rem) + temp;
          n = Math.floor((n - 1) / 26);
        }
        return temp;
      };

      // Clear existing rules to avoid duplicates.
      sheet.clearConditionalFormatRules();

      const palette: Record<string, string> = {
        Alpha: '#E3F2FD',
        Bravo: '#FCE4EC',
        Charlie: '#F3E5F5',
        Delta: '#E8F5E9',
        Echo: '#FFF3E0',
        Foxtrot: '#E0F7FA',
        Abroad: '#ECEFF1',
        Trine: '#FFFDE7',
        Valparaiso: '#EDE7F6',
      };

      const flightColLetter = columnToLetter(flightCol);
      const rules = Object.entries(palette).map(([flight, color]) =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=$${flightColLetter}${startRow}="${flight}"`)
          .setBackground(color)
          .setRanges([dataRange])
          .build(),
      );

      try {
        sheet.setConditionalFormatRules(rules);
      } catch (err) {
        Log.warn(`Unable to set conditional formatting on Attendance Backend: ${err}`);
      }
    }


  function ensureFormTrigger(handlerName: string, formId: string) {
    if (!formId) {
      Log.warn(`Cannot create form trigger for handler=${handlerName}: formId missing`);
      return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    const matching = triggers.filter((t) => t.getHandlerFunction() === handlerName);
    const alreadyCorrect = matching.some((t) => {
      try {
        return t.getTriggerSourceId && t.getTriggerSourceId() === formId;
      } catch {
        return false;
      }
    });
    if (alreadyCorrect) return;

    // Clean up stale triggers for the same handler so we don't keep firing against old/deleted forms.
    matching.forEach((t) => {
      try {
        const sourceId = t.getTriggerSourceId?.();
        if (sourceId && sourceId !== formId) {
          Log.warn(`Deleting stale trigger for handler=${handlerName} sourceId=${sourceId}`);
          ScriptApp.deleteTrigger(t);
        }
      } catch {
        // Ignore; we'll create a new correct trigger below.
      }
    });

    Log.info(`Creating form submit trigger for handler=${handlerName} formId=${formId}`);
    ScriptApp.newTrigger(handlerName).forForm(formId).onFormSubmit().create();
  }

  function normalizeResponseSheetsForForms(
    spreadsheetId: string,
    forms: Array<{ formId: string; desiredSheetName: string }>,
  ) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const desiredByFormId = new Map(forms.map((f) => [f.formId, f.desiredSheetName] as const));

    const responseishSheets = ss
      .getSheets()
      .filter((s) => RESPONSE_SHEET_REGEX.test(s.getName()) || Array.from(desiredByFormId.values()).includes(s.getName()));

    // Group response(-ish) sheets by linked form ID, when present.
    const sheetsByFormId = new Map<string, GoogleAppsScript.Spreadsheet.Sheet[]>();
    const unlinked: GoogleAppsScript.Spreadsheet.Sheet[] = [];
    responseishSheets.forEach((sheet) => {
      let formId: string | null = null;
      try {
        formId = extractFormIdFromUrl(sheet.getFormUrl() || '');
      } catch {
        formId = null;
      }
      if (!formId) {
        unlinked.push(sheet);
        return;
      }
      const arr = sheetsByFormId.get(formId) || [];
      arr.push(sheet);
      sheetsByFormId.set(formId, arr);
    });

    // For each known SHAMROCK form, ensure its linked response sheet has the desired name.
    forms.forEach(({ formId, desiredSheetName }) => {
      const linked = sheetsByFormId.get(formId) || [];
      if (linked.length === 0) {
        Log.warn(`No response sheet currently linked to formId=${formId} to rename to '${desiredSheetName}'`);
        return;
      }

      // Prefer a sheet already named correctly.
      const primary = linked.find((s) => s.getName() === desiredSheetName) || linked[0];
      if (primary.getName() !== desiredSheetName) {
        Log.info(`Renaming linked response sheet ${primary.getName()} -> ${desiredSheetName} (formId=${formId})`);
        try {
          primary.setName(desiredSheetName);
        } catch (err) {
          Log.warn(`Unable to rename response sheet to '${desiredSheetName}'. Error: ${err}`);
        }
      }

      // Any other linked sheets for the same form are likely historical destination churn; archive their names.
      linked
        .filter((s) => s.getSheetId() !== primary.getSheetId())
        .forEach((s) => {
          if (/^Archived - /i.test(s.getName())) return;
          const archivedName = `Archived - ${desiredSheetName} (${s.getName()})`;
          try {
            Log.warn(`Archiving extra linked response sheet ${s.getName()} -> ${archivedName} (formId=${formId})`);
            s.setName(archivedName);
          } catch (err) {
            Log.warn(`Unable to archive response sheet ${s.getName()}. Error: ${err}`);
          }
        });
    });

    // For unlinked "Form Responses" sheets, just archive them so they stop looking active.
    unlinked.forEach((s) => {
      if (!RESPONSE_SHEET_REGEX.test(s.getName())) return;
      const archivedName = `Archived - ${s.getName()}`;
      try {
        Log.warn(`Archiving unlinked response sheet ${s.getName()} -> ${archivedName}`);
        s.setName(archivedName);
      } catch {
        // Ignore name collisions or protected states.
      }
    });
  }

  function ensureSpreadsheetTrigger(handlerName: string, spreadsheetId: string, event: 'open' | 'edit') {
    if (!spreadsheetId) {
      Log.warn(`Cannot create ${event} trigger for ${handlerName}: spreadsheetId missing.`);
      return;
    }
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some((t) => t.getHandlerFunction() === handlerName && t.getTriggerSourceId?.() === spreadsheetId);
    if (exists) return;
    Log.info(`Creating ${event} trigger for handler=${handlerName} spreadsheet=${spreadsheetId}`);
    const builder = ScriptApp.newTrigger(handlerName).forSpreadsheet(spreadsheetId);
    if (event === 'open') {
      builder.onOpen().create();
    } else {
      builder.onEdit().create();
    }
  }

  function ensureTimeTrigger(handlerName: string, weekDay: GoogleAppsScript.Base.Weekday, hour: number) {
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some((t) => t.getHandlerFunction() === handlerName && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK);
    if (exists) return;
    Log.info(`Creating time trigger handler=${handlerName} day=${weekDay} hour=${hour}`);
    ScriptApp.newTrigger(handlerName).timeBased().onWeekDay(weekDay).atHour(hour).create();
  }

  function ensurePeriodicTrigger(handlerName: string, intervalMinutes: number) {
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some((t) => t.getHandlerFunction() === handlerName && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK);
    if (exists) return;
    Log.info(`Creating periodic trigger handler=${handlerName} interval=${intervalMinutes}min`);
    ScriptApp.newTrigger(handlerName).timeBased().everyMinutes(intervalMinutes).create();
  }

  // Deletes all installable triggers, then reinstalls the canonical SHAMROCK triggers for forms and spreadsheets.
  export function reinstallAllTriggers() {
    Log.info('Reinstalling all installable triggers');

    // Clear existing triggers first.
    ScriptApp.getProjectTriggers().forEach((t) => {
      try {
        ScriptApp.deleteTrigger(t);
      } catch (err) {
        Log.warn(`Unable to delete trigger ${t.getUniqueId?.() || ''}: ${err}`);
      }
    });

    // Recreate spreadsheet triggers (onOpen/onEdit) for frontend, backend, and excusals management.
    const frontendId = Config.getFrontendId();
    const backendId = Config.getBackendId();
    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID') || '';
    ensureSpreadsheetTrigger('onFrontendOpen', frontendId, 'open');
    ensureSpreadsheetTrigger('onFrontendEdit', frontendId, 'edit');
    ensureSpreadsheetTrigger('onBackendOpen', backendId, 'open');
    ensureSpreadsheetTrigger('onBackendEdit', backendId, 'edit');
    ensureSpreadsheetTrigger('onExcusalsManagementEdit', managementId, 'edit');

    // Time-based trigger: reconcile frontend Directory edits every 10 minutes (handles edits by unauthorized users).
    ensurePeriodicTrigger('reconcilePendingDirectoryEdits', 10);

    // Recreate form submit triggers for attendance/excusal/directory.
    const props = Config.scriptProperties();
    const attendanceFormId = props.getProperty(Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID) || '';
    const excusalFormId = props.getProperty(Config.PROPERTY_KEYS.EXCUSALS_FORM_ID) || '';
    const directoryFormId = props.getProperty(Config.PROPERTY_KEYS.DIRECTORY_FORM_ID) || '';

    if (attendanceFormId) ensureFormTrigger('onAttendanceFormSubmit', attendanceFormId);
    else Log.warn('Cannot reinstall attendance form trigger: ATTENDANCE_FORM_ID missing. Run setup first.');

    if (excusalFormId) ensureFormTrigger('onExcusalsFormSubmit', excusalFormId);
    else Log.warn('Cannot reinstall excusals form trigger: EXCUSAL_FORM_ID missing. Run setup first.');

    if (directoryFormId) ensureFormTrigger('onDirectoryFormSubmit', directoryFormId);
    else Log.warn('Cannot reinstall directory form trigger: DIRECTORY_FORM_ID missing. Run setup first.');

    // Time-driven reminders
    ensureTimeTrigger('sendWeeklyMandoExcusedSummary', ScriptApp.WeekDay.THURSDAY, 5);
    ensureTimeTrigger('sendWeeklyLlabExcusedSummary', ScriptApp.WeekDay.TUESDAY, 12);
    ensureTimeTrigger('sendWeeklyUnexcusedSummary', ScriptApp.WeekDay.SUNDAY, 19);
  }

  function removeDefaultSheetIfPresent(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, allowedNames: Set<string>) {
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet && !allowedNames.has('Sheet1')) {
      // Only remove if there is more than one sheet to avoid deleting the last sheet in a spreadsheet.
      if (spreadsheet.getSheets().length > 1) {
        Log.info(`Removing default sheet 'Sheet1' from spreadsheet=${spreadsheet.getId()}`);
        spreadsheet.deleteSheet(defaultSheet);
      } else {
        Log.warn(`Default sheet 'Sheet1' present but is the only sheet; skipping delete in spreadsheet=${spreadsheet.getId()}`);
      }
    }
  }

  function ensureForm(
    kind: 'attendance' | 'excusals' | 'directory',
    name: string,
    propertyKey: string,
    destinationSpreadsheetId?: string,
  ): Types.EnsureFormResult {
    Log.info(`Ensuring form kind=${kind}`);
    const props = Config.scriptProperties();
    const existingId = props.getProperty(propertyKey);
    let form: GoogleAppsScript.Forms.Form | null = null;
    let created = false;

    if (existingId) {
      try {
        form = FormApp.openById(existingId);
        Log.info(`Found existing form id=${existingId}`);
      } catch (err) {
        Log.warn(`Stored form id invalid for key=${propertyKey}; creating new. Error: ${err}`);
      }
    }

    if (!form) {
      form = FormApp.create(name);
      created = true;
      props.setProperty(propertyKey, form.getId());
      Log.info(`Created form name=${name} id=${form.getId()}`);
    }

    // Keep form title stable (helps ops/debugging).
    try {
      if (form.getTitle() !== name) form.setTitle(name);
    } catch (err) {
      Log.warn(`Unable to set form title. Error: ${err}`);
    }

    // Enforce responder email collection and login requirement (verified identity).
    form.setCollectEmail(true);
    try {
      form.setRequireLogin(true);
    } catch (err) {
      // setRequireLogin is not supported for consumer accounts; log and continue.
      Log.warn(`setRequireLogin not supported in this environment; continuing without it. Error: ${err}`);
    }

    // Response edit policy per form type.
    try {
      if (kind === 'directory') {
        form.setAllowResponseEdits(true);
      } else {
        form.setAllowResponseEdits(false);
      }
    } catch (err) {
      Log.warn(`setAllowResponseEdits not supported in this environment; continuing without it. Error: ${err}`);
    }

    // Route responses into the backend spreadsheet when provided, with retries.
    let destinationConfigured = false;
    if (destinationSpreadsheetId) {
      const currentDestinationId = getFormDestinationSpreadsheetId(form);
      if (currentDestinationId && currentDestinationId === destinationSpreadsheetId) {
        destinationConfigured = true;
      } else {
        if (currentDestinationId && currentDestinationId !== destinationSpreadsheetId) {
          Log.warn(
            `Form destination differs from desired; updating. current=${currentDestinationId} desired=${destinationSpreadsheetId} formId=${form.getId()}`,
          );
        }

        const attemptSetDestination = () => {
        try {
          form.setDestination(FormApp.DestinationType.SPREADSHEET, destinationSpreadsheetId);
          return true;
        } catch (err) {
          Log.warn(`Unable to set form destination to spreadsheet=${destinationSpreadsheetId}. Error: ${err}`);
          return false;
        }
      };

        // Only set destination when needed; retry to handle transient failures.
        for (let i = 0; i < 3 && !destinationConfigured; i++) {
          if (attemptSetDestination()) {
            destinationConfigured = true;
            break;
          }
          Utilities.sleep(500);
        }
      }
    }

    // Note: Built-in "email a copy of responses" setting is not reliably controllable via Apps Script.
    // Future: implement onFormSubmit email receipt as part of the form handler.
    if (kind === 'directory') {
      try {
        form.setConfirmationMessage(
          'Thanks! Please save your response edit link from the confirmation screen so you can update your information later.',
        );
      } catch (err) {
        Log.warn(`Unable to set confirmation message. Error: ${err}`);
      }
    }

    // Seed questions if empty.
    if (kind === 'attendance') FormService.ensureAttendanceForm(form);
    if (kind === 'excusals') FormService.ensureExcusalsForm(form);
    if (kind === 'directory') FormService.ensureDirectoryForm(form);

    // Ensure the real response sheet exists and is named correctly (avoid dummy placeholders).
    if (destinationSpreadsheetId) {
      const desired =
        kind === 'attendance'
          ? Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET
          : kind === 'excusals'
          ? Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET
          : Config.RESOURCE_NAMES.DIRECTORY_FORM_SHEET;
      ensureResponseSheetForForm(form, desired, destinationSpreadsheetId);
    }

    return {
      kind,
      id: form.getId(),
      created,
      url: form.getEditUrl(),
    };
  }

  export function applyFrontendFormatting() {
    const frontendId = Config.getFrontendId();
    if (isFrontendFormattingDisabled()) {
      Log.info('Frontend formatting skipped because DISABLE_FRONTEND_FORMATTING is true.');
      return;
    }
    FrontendFormattingService.applyAll(frontendId);
  }

  export function rebuildDashboard() {
    const frontendId = Config.getFrontendId();
    if (isFrontendFormattingDisabled()) {
      Log.info('Dashboard rebuild skipped because DISABLE_FRONTEND_FORMATTING is true.');
      return;
    }
    FrontendFormattingService.applyDashboardOnly(frontendId);
  }

  export function reapplyFrontendProtections() {
    const frontendId = Config.getFrontendId();
    ProtectionService.applyFrontendProtections(frontendId);
  }

  export function toggleFrontendFormatting() {
    const props = Config.scriptProperties();
    const current = String(props.getProperty('DISABLE_FRONTEND_FORMATTING') || '').toLowerCase();
    const next = current === 'true' ? '' : 'true';
    props.setProperty('DISABLE_FRONTEND_FORMATTING', next);
    const status = next === 'true' ? 'OFF (disabled)' : 'ON (enabled)';
    const msg = `Frontend formatting is now ${status}.`;
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  function reorderSheets(spreadsheetId: string, desiredOrder: string[]) {
    if (!spreadsheetId) {
      Log.warn('Cannot reorder sheets: spreadsheetId missing.');
      return;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let position = 1;

    const moveSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
      try {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(position++);
      } catch (err) {
        Log.warn(`Unable to move sheet ${sheet.getName()} in spreadsheet=${spreadsheetId}: ${err}`);
      }
    };

    desiredOrder.forEach((name) => {
      const sheet = ss.getSheetByName(name);
      if (sheet) moveSheet(sheet);
    });

    ss.getSheets()
      .filter((s) => !desiredOrder.includes(s.getName()))
      .forEach((sheet) => moveSheet(sheet));
  }

  export function toggleFrontendColumnWidths() {
    const props = Config.scriptProperties();
    const current = String(props.getProperty('DISABLE_FRONTEND_COLUMN_WIDTHS') || '').toLowerCase();
    const next = current === 'true' ? '' : 'true';
    props.setProperty('DISABLE_FRONTEND_COLUMN_WIDTHS', next);
    const status = next === 'true' ? 'OFF (disabled)' : 'ON (enabled)';
    const msg = `Frontend column width formatting is now ${status}.`;
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  export function pauseAutomations(reason = 'manual pause') {
    PauseService.pause(reason);
    const msg = `Automation is now PAUSED (${PauseService.pauseInfo()}). Frontend edits will be deferred.`;
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  export function resumeAutomations() {
    const wasPaused = PauseService.isPaused();
    PauseService.resume();

    // Batch mirror any frontend Directory edits made while paused back into the backend, then resync artifacts.
    const reconciliation = FrontendEditService.reconcilePendingDirectoryEdits();
    SyncService.syncAllMapped();
    rebuildAttendanceMatrix();
    refreshAttendanceFormEventChoices();
    refreshExcusalsFormEventChoices();
    applyFrontendFormatting();

    const msg = wasPaused
      ? `Automation resumed. Applied ${reconciliation.updated} Directory updates (missing matches: ${reconciliation.missing}).`
      : 'Automation was not paused; performed refresh anyway.';
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  export function refreshDataLegendAndFrontend() {
    DataLegendService.refreshLegendFromArrays();
    SyncService.syncByBackendSheetName('Data Legend');
    applyFrontendFormatting();
  }

  export function syncDirectoryBackendToFrontend() {
    syncDirectoryFrontend();
    applyFrontendFormatting();
  }

  export function syncLeadershipBackendToFrontend() {
    SyncService.syncByBackendSheetName('Leadership Backend');
    applyFrontendFormatting();
  }

  export function syncDataLegendBackendToFrontend() {
    SyncService.syncByBackendSheetName('Data Legend');
    applyFrontendFormatting();
  }

  export function syncAllBackendToFrontend() {
    SyncService.syncAllMapped();
    applyFrontendFormatting();
  }

  export function applyAttendanceBackendFormattingPublic() {
    applyAttendanceBackendFormatting();
  }

  export function syncDirectoryFrontend() {
    const frontendId = Config.getFrontendId();
    if (frontendId) DirectoryService.protectFrontendDirectory(frontendId);
    DirectoryService.syncDirectoryFrontend();
  }

  export function rebuildAttendanceMatrix() {
    AttendanceService.rebuildMatrix();
    try {
      // Re-apply Attendance header formatting and validations after matrix rebuild.
      fixAttendanceHeaders();
      reapplyFrontendProtections();
    } catch (err) {
      Log.warn(`fixAttendanceHeaders post-rebuild failed: ${err}`);
    }
  }

  export function refreshAttendanceForm() {
    const backendId = Config.getBackendId();
    ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();
    pruneExcusalsResponseColumnsExplicit();
  }

  export function refreshExcusalsForm() {
    const backendId = Config.getBackendId();
    const ensured = ensureForm('excusals', Config.RESOURCE_NAMES.EXCUSALS_FORM, Config.PROPERTY_KEYS.EXCUSALS_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.refreshExcusalsFormEventChoices(form);
    pruneExcusalsResponseColumnsExplicit();
  }

  export function rebuildAttendanceForm() {
    const backendId = Config.getBackendId();
    const ensured = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.rebuildAttendanceForm(form);
    // After rebuilding questions, refresh event list and clean up response artifacts.
    FormService.refreshAttendanceFormEventChoices(form);
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();
  }

  export function refreshAttendanceFormEventChoices() {
    const backendId = Config.getBackendId();
    const ensured = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.refreshAttendanceFormEventChoices(form);
  }

  export function reorderFrontendSheets() {
    const frontendId = Config.getFrontendId();
    const desired = ['FAQs', 'Dashboard', 'Leadership', 'Directory', 'Attendance', 'Data Legend'];
    reorderSheets(frontendId, desired);
  }

  export function reorderBackendSheets() {
    const backendId = Config.getBackendId();
    const desired = [
      Config.RESOURCE_NAMES.DIRECTORY_FORM_SHEET,
      Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET,
      Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET,
      'Leadership Backend',
      'Directory Backend',
      'Attendance Matrix Backend',
      'Attendance Backend',
      'Events Backend',
      'Excusals Backend',
      'Audit Backend',
      'Data Legend',
    ];
    reorderSheets(backendId, desired);
  }

  export function refreshExcusalsFormEventChoices() {
    const backendId = Config.getBackendId();
    const ensured = ensureForm('excusals', Config.RESOURCE_NAMES.EXCUSALS_FORM, Config.PROPERTY_KEYS.EXCUSALS_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.refreshExcusalsFormEventChoices(form);
  }

  export function pruneAttendanceResponseColumns() {
    pruneAttendanceResponseColumnsExplicit();
  }

  export function pruneExcusalsResponseColumns() {
    pruneExcusalsResponseColumnsExplicit();
  }

  export function refreshEventsArtifacts() {
    SyncService.syncByBackendSheetName('Events Backend');
    rebuildAttendanceMatrix();
    refreshAttendanceFormEventChoices();
    applyFrontendFormatting();
  }

  export function showMenuHelp() {
    const lines = [
      'Sync Directory/Leadership/Data Legend: copies backend sheet -> frontend twin (frontend is overwritten for those sheets).',
      'Sync ALL mapped: runs all backend -> frontend copies (Directory, Leadership, Data Legend).',
      'Refresh Events + Attendance: sync events backend -> frontend artifacts and rebuild attendance matrix/form choices.',
      'Rebuild Attendance Matrix: replay attendance backend log -> frontend matrix.',
      'Rebuild Attendance Form: rebuild questions + event choices from backend events.',
      'CSV exports: create Drive CSV from the specified backend table. Imports overwrite the target backend sheet (headers must match).',
      'Protections/formatting items only affect frontend presentation; data comes from backend tables.',
    ];

    const msg = lines.join('\n');
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  export function archiveCoreSheets() {
    const frontendId = Config.getFrontendId();
    const backendId = Config.getBackendId();
    const frontendNames = ['Leadership', 'Directory', 'Attendance'];
    const backendNames = ['Leadership Backend', 'Directory Backend', 'Attendance Backend'];

    archiveAndResetSheets(frontendId, Schemas.FRONTEND_TABS, frontendNames);
    archiveAndResetSheets(backendId, Schemas.BACKEND_TABS, backendNames);

    if (frontendId) {
      ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
        ensureTableForSheet(frontendId, name, name.replace(/\s+/g, '_').toLowerCase());
      });
      if (!isFrontendFormattingDisabled()) {
        FrontendFormattingService.applyAll(frontendId);
      } else {
        Log.info('Skipping frontend formatting during archive because DISABLE_FRONTEND_FORMATTING is true.');
      }
      ProtectionService.applyFrontendProtections(frontendId);
    }

    if (backendId) {
      applyAttendanceBackendFormatting();
    }
  }

  export function restoreCoreSheetsFromArchive() {
    const frontendId = Config.getFrontendId();
    const backendId = Config.getBackendId();
    const frontendNames = ['Leadership', 'Directory', 'Attendance'];
    const backendNames = ['Leadership Backend', 'Directory Backend', 'Attendance Backend'];

    restoreFromArchiveSheets(frontendId, Schemas.FRONTEND_TABS, frontendNames);
    restoreFromArchiveSheets(backendId, Schemas.BACKEND_TABS, backendNames);

    if (frontendId) {
      ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
        ensureTableForSheet(frontendId, name, name.replace(/\s+/g, '_').toLowerCase());
      });
      if (!isFrontendFormattingDisabled()) {
        FrontendFormattingService.applyAll(frontendId);
      } else {
        Log.info('Skipping frontend formatting during restore because DISABLE_FRONTEND_FORMATTING is true.');
      }
      ProtectionService.applyFrontendProtections(frontendId);
    }

    if (backendId) {
      applyAttendanceBackendFormatting();
    }
  }

  export function runSetup(): Types.SetupSummary {
    Log.info('Starting setup (ensure-exists)');
    const spreadsheetResults: Types.EnsureSpreadsheetResult[] = [];
    const sheetResults: Types.EnsureSheetResult[] = [];
    const formResults: Types.EnsureFormResult[] = [];

    // Ensure spreadsheets.
    const frontend = ensureSpreadsheet('frontend', Config.RESOURCE_NAMES.FRONTEND_SPREADSHEET, Config.PROPERTY_KEYS.FRONTEND_SHEET_ID);
    const backend = ensureSpreadsheet('backend', Config.RESOURCE_NAMES.BACKEND_SPREADSHEET, Config.PROPERTY_KEYS.BACKEND_SHEET_ID);
    spreadsheetResults.push(frontend, backend);

    // Ensure excusals management spreadsheet.
    try {
      ExcusalsService.ensureManagementSpreadsheet();
      ExcusalsService.shareAndProtectManagementSpreadsheet();
    } catch (err) {
      Log.warn(`Failed to ensure excusals management spreadsheet: ${err}`);
    }

    // Ensure frontend sheets.
    const frontendSheet = SpreadsheetApp.openById(frontend.id);
    Schemas.FRONTEND_TABS.forEach((tab) => {
      sheetResults.push(ensureSheet(frontendSheet, tab));
    });
    removeDefaultSheetIfPresent(frontendSheet, new Set(Schemas.FRONTEND_TABS.map((t) => t.name)));
    restoreMissingHeaders(frontendSheet, Schemas.FRONTEND_TABS);

    // Ensure backend sheets.
    const backendSheet = SpreadsheetApp.openById(backend.id);
    Schemas.BACKEND_TABS.forEach((tab) => {
      sheetResults.push(ensureSheet(backendSheet, tab));
    });
    removeDefaultSheetIfPresent(backendSheet, new Set(Schemas.BACKEND_TABS.map((t) => t.name)));
    restoreMissingHeaders(backendSheet, Schemas.BACKEND_TABS);

    // Ensure forms.
    const attendanceForm = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backend.id);
    const excusalForm = ensureForm('excusals', Config.RESOURCE_NAMES.EXCUSALS_FORM, Config.PROPERTY_KEYS.EXCUSALS_FORM_ID, backend.id);
    const directoryForm = ensureForm('directory', Config.RESOURCE_NAMES.DIRECTORY_FORM, Config.PROPERTY_KEYS.DIRECTORY_FORM_ID, backend.id);
    formResults.push(attendanceForm, excusalForm, directoryForm);

    // Normalize response sheet names based on the form actually linked to each sheet.
    normalizeResponseSheetsForForms(backend.id, [
      { formId: attendanceForm.id, desiredSheetName: Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET },
      { formId: excusalForm.id, desiredSheetName: Config.RESOURCE_NAMES.EXCUSALS_FORM_SHEET },
      { formId: directoryForm.id, desiredSheetName: Config.RESOURCE_NAMES.DIRECTORY_FORM_SHEET },
    ]);
    // Slim attendance response sheet to drop stale/duplicate columns left over from form rebuilds (keeps any columns that still have data).
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();

    // Refresh event choices for forms (attendance + excusals) after ensuring sheets/forms.
    refreshAttendanceFormEventChoices();
    refreshExcusalsFormEventChoices();

    // Ensure form submit triggers for receipts/processing.
    ensureFormTrigger('onAttendanceFormSubmit', attendanceForm.id);
    ensureFormTrigger('onExcusalsFormSubmit', excusalForm.id);
    ensureFormTrigger('onDirectoryFormSubmit', directoryForm.id);

    // Refresh Data Legend from canonical arrays and sync to frontend.
    refreshDataLegendAndFrontend();

    // Protect user-facing directory and sync it from backend.
    ProtectionService.applyFrontendProtections(frontend.id);
    DirectoryService.syncDirectoryFrontend();

    // Apply frontend validations/banding after syncs.
    if (!isFrontendFormattingDisabled()) {
      FrontendFormattingService.applyAll(frontend.id);
    } else {
      Log.info('Skipping frontend formatting during setup because DISABLE_FRONTEND_FORMATTING is true.');
    }

    // Create structured tables on key frontend sheets via Sheets API (skip if formatting disabled).
    if (!isFrontendFormattingDisabled()) {
      ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
        ensureTableForSheet(frontend.id, name, name.replace(/\s+/g, '_').toLowerCase());
      });
    } else {
      Log.info('Skipping table creation because DISABLE_FRONTEND_FORMATTING is true.');
    }

    // Build attendance matrix initially.
    rebuildAttendanceMatrix();

    // Order sheets for predictable UX
    reorderFrontendSheets();
    reorderBackendSheets();

    // Install onOpen triggers for menus and onEdit trigger for backend directory sync.
    ensureSpreadsheetTrigger('onFrontendOpen', frontend.id, 'open');
    ensureSpreadsheetTrigger('onFrontendEdit', frontend.id, 'edit');
    ensureSpreadsheetTrigger('onBackendOpen', backend.id, 'open');
    ensureSpreadsheetTrigger('onBackendEdit', backend.id, 'edit');

    // Time-based trigger: reconcile frontend Directory edits every 10 minutes (handles edits by unauthorized users).
    ensurePeriodicTrigger('reconcilePendingDirectoryEdits', 10);

    Log.info(`Setup finished: spreadsheets=${spreadsheetResults.length}, sheets=${sheetResults.length}, forms=${formResults.length}`);

    return {
      spreadsheets: spreadsheetResults,
      sheets: sheetResults,
      forms: formResults,
    };
  }
}
