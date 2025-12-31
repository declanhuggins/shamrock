// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

type AuditEntry = {
  audit_id?: string;
  timestamp?: string;
  actor_email?: string;
  actor_role?: string;
  action: string;
  target_sheet?: string;
  target_table?: string;
  target_key?: string;
  target_range?: string;
  event_id?: string;
  request_id?: string;
  old_value?: string;
  new_value?: string;
  result?: string;
  reason?: string;
  notes?: string;
  source: string;
  script_version?: string;
  run_id?: string;
};

Shamrock.logAudit = function (entry: AuditEntry): void {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.audit) || ss.insertSheet(Shamrock.BACKEND_SHEET_NAMES.audit);
  const runId = entry.run_id || Utilities.getUuid();
  const auditId = entry.audit_id || `AUD-${Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd-HHmmss")}-${Utilities.getUuid().slice(0, 4)}`;
  const row: Record<typeof Shamrock.AUDIT_FIELDS[number], any> = {
    audit_id: auditId,
    timestamp: entry.timestamp || Shamrock.nowIso(),
    actor_email: entry.actor_email || getActorEmail(),
    actor_role: entry.actor_role || "System",
    action: entry.action,
    target_sheet: entry.target_sheet || "",
    target_table: entry.target_table || "",
    target_key: entry.target_key || "",
    target_range: entry.target_range || "",
    event_id: entry.event_id || "",
    request_id: entry.request_id || "",
    old_value: entry.old_value || "",
    new_value: entry.new_value || "",
    result: entry.result || "success",
    reason: entry.reason || "",
    notes: entry.notes || "",
    source: entry.source,
    script_version: entry.script_version || "v0",
    run_id: runId,
  };
  Shamrock.withLock(() => Shamrock.appendAuditRow(sheet, row));
};

function getActorEmail(): string {
  const email = Session.getActiveUser().getEmail();
  if (!email) return "system";
  return email.toLowerCase();
}
