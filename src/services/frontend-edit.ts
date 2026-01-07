// Frontend onEdit handling for Directory edits with audit and backend sync.

namespace FrontendEditService {
  type AllowedField = 'as_year' | 'flight' | 'squadron' | 'university';

  const ALLOWED_FIELDS: AllowedField[] = ['as_year', 'flight', 'squadron', 'university'];

  function getAllowedEditors(): string[] {
    try {
      const prop = Config.scriptProperties().getProperty('SHAMROCK_MENU_ALLOWED_EMAILS') || '';
      return prop
        .split(',')
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);
    } catch (err) {
      Log.warn(`Unable to read SHAMROCK_MENU_ALLOWED_EMAILS for edit gate: ${err}`);
      return [];
    }
  }

  function actorEmail(): string {
    try {
      return (Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (err) {
      Log.warn(`Unable to read active user email during onEdit: ${err}`);
      return '';
    }
  }

  function allowedFieldFromHeader(header: string): AllowedField | null {
    const key = String(header || '').trim().toLowerCase();
    return (ALLOWED_FIELDS as string[]).includes(key) ? (key as AllowedField) : null;
  }

  function findBackendRowIndex(table: SheetUtils.TableData, predicate: (row: any) => boolean): number {
    for (let i = 0; i < table.rows.length; i++) {
      if (predicate(table.rows[i])) return i;
    }
    return -1;
  }

  function appendAuditEntry(params: {
    backendId: string;
    targetRange: string;
    targetKey: string;
    header: string;
    oldValue: string;
    newValue: string;
  }) {
    const { backendId, targetRange, targetKey, header, oldValue, newValue } = params;
    const audit = SheetUtils.getSheet(backendId, 'Audit Backend');
    if (!audit) return;

    const headers = Schemas.BACKEND_TABS.find((t) => t.name === 'Audit Backend')?.machineHeaders || [];
    const row: any = {};
    headers.forEach((h) => (row[h] = ''));

    row['audit_id'] = Utilities.getUuid();
    row['timestamp'] = new Date();
    row['actor_email'] = actorEmail() || 'unknown';
    row['role'] = 'frontend_editor';
    row['action'] = 'edit';
    row['target_sheet'] = 'Directory';
    row['target_table'] = 'directory';
    row['target_key'] = targetKey;
    row['target_range'] = targetRange;
    row['old_value'] = oldValue;
    row['new_value'] = newValue;
    row['result'] = 'ok';
    row['source'] = 'onFrontendEdit';
    row['version'] = 'v1';

    // Append respecting header order
    const values = headers.map((h) => row[h] ?? '');
    const nextRow = Math.max(3, audit.getLastRow() + 1);
    audit.getRange(nextRow, 1, 1, headers.length).setValues([values]);
  }

  function applyDirectoryEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    const range = e?.range;
    const sheet = range?.getSheet();
    if (!sheet || sheet.getName() !== 'Directory') return;

    const row = range.getRow();
    const col = range.getColumn();
    if (row < 3) return; // headers

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    const header = headers[col - 1];
    const allowedField = allowedFieldFromHeader(header);
    if (!allowedField) return;

    const newValue = String(e?.value ?? range.getValue() ?? '');

    const backendId = Config.getBackendId();
    const backendSheet = SheetUtils.getSheet(backendId, 'Directory Backend');
    if (!backendSheet) {
      Log.warn('Backend Directory sheet missing; cannot mirror frontend edit');
      return;
    }

    const table = SheetUtils.readTable(backendSheet);
    const emailIdx = table.headers.indexOf('email');
    const lastIdx = table.headers.indexOf('last_name');
    const firstIdx = table.headers.indexOf('first_name');

    const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    const email = emailIdx >= 0 ? String(rowValues[headers.indexOf('email')] || '').toLowerCase() : '';
    const last = lastIdx >= 0 ? String(rowValues[headers.indexOf('last_name')] || '').toLowerCase() : '';
    const first = firstIdx >= 0 ? String(rowValues[headers.indexOf('first_name')] || '').toLowerCase() : '';

    const matchIdx = findBackendRowIndex(table, (r) => {
      const rEmail = String(r['email'] || '').toLowerCase();
      if (email && rEmail === email) return true;
      return String(r['last_name'] || '').toLowerCase() === last && String(r['first_name'] || '').toLowerCase() === first;
    });

    if (matchIdx < 0) {
      Log.warn(`No backend Directory match for edit row=${row} email=${email}`);
      return;
    }

    const backendHeaders = table.headers;
    const backendColIdx = backendHeaders.indexOf(allowedField);
    if (backendColIdx < 0) return;

    const oldValue = String(table.rows[matchIdx][allowedField] || '');
    if (oldValue === newValue) return;

    backendSheet.getRange(matchIdx + 3, backendColIdx + 1).setValue(newValue);

    const targetKey = email || `${last},${first}`;
    appendAuditEntry({
      backendId,
      targetRange: `${sheet.getName()}!${range.getA1Notation()}`,
      targetKey,
      header: allowedField,
      oldValue,
      newValue,
    });

    // Propagate: sync frontend, rebuild attendance matrix (frontend + backend), refresh attendance form roster.
    SyncService.syncByBackendSheetName('Directory');
    AttendanceService.rebuildMatrix();
  }

  export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    if (PauseService.isPaused()) {
      Log.info('Frontend onEdit is paused; skipping propagation.');
      return;
    }

    const email = actorEmail();
    const allowed = getAllowedEditors();
    if (allowed.length && !allowed.includes(email)) {
      Log.warn(`Frontend onEdit blocked for user=${email || 'unknown'}`);
      return;
    }

    try {
      applyDirectoryEdit(e);
    } catch (err) {
      Log.warn(`Frontend onEdit failed: ${err}`);
    }
  }

  export function reconcilePendingDirectoryEdits(): { updated: number; missing: number } {
    const backendId = Config.getBackendId();
    const frontendId = Config.getFrontendId();
    const backendSheet = SheetUtils.getSheet(backendId, 'Directory Backend');
    const frontendSheet = SheetUtils.getSheet(frontendId, 'Directory');
    if (!backendSheet || !frontendSheet) return { updated: 0, missing: 0 };

    const backendTable = SheetUtils.readTable(backendSheet);
    const frontendTable = SheetUtils.readTable(frontendSheet);
    const backendHeaders = backendTable.headers;

    const backendLookup = new Map<string, { idx: number; row: any }>();
    backendTable.rows.forEach((row, idx) => {
      const email = String(row['email'] || '').toLowerCase();
      const last = String(row['last_name'] || '').toLowerCase();
      const first = String(row['first_name'] || '').toLowerCase();
      const key = email || (last && first ? `${last},${first}` : '');
      if (key) backendLookup.set(key, { idx, row });
    });

    let updated = 0;
    let missing = 0;

    frontendTable.rows.forEach((row) => {
      const email = String(row['email'] || '').toLowerCase();
      const last = String(row['last_name'] || '').toLowerCase();
      const first = String(row['first_name'] || '').toLowerCase();
      const key = email || (last && first ? `${last},${first}` : '');
      if (!key) return;

      const backendMatch = backendLookup.get(key);
      if (!backendMatch) {
        missing++;
        return;
      }

      ALLOWED_FIELDS.forEach((field) => {
        const backendColIdx = backendHeaders.indexOf(field);
        if (backendColIdx < 0) return;
        const oldValue = String(backendMatch.row[field] || '');
        const newValue = String(row[field] || '');
        if (oldValue === newValue) return;

        backendSheet.getRange(backendMatch.idx + 3, backendColIdx + 1).setValue(newValue);
        appendAuditEntry({
          backendId,
          targetRange: 'Directory (batch reconcile)',
          targetKey: key,
          header: field,
          oldValue,
          newValue,
        });
        updated++;
      });
    });

    return { updated, missing };
  }
}
