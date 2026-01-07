// Sync helpers to mirror backend tables to frontend counterparts where schemas align.

namespace SyncService {
  const MAPPINGS: { backend: string; frontend: string }[] = [
    { backend: 'Directory Backend', frontend: 'Directory' },
    { backend: 'Leadership Backend', frontend: 'Leadership' },
    { backend: 'Data Legend', frontend: 'Data Legend' },
  ];

  function getIds() {
    return {
      backendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.BACKEND_SHEET_ID) || '',
      frontendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.FRONTEND_SHEET_ID) || '',
    };
  }

  function copyTable(backendSheetName: string, frontendSheetName: string) {
    const { backendId, frontendId } = getIds();
    if (!backendId || !frontendId) return;
    const backendSheet = SheetUtils.getSheet(backendId, backendSheetName);
    const frontendSheet = SheetUtils.getSheet(frontendId, frontendSheetName);
    if (!backendSheet || !frontendSheet) return;
    const data = SheetUtils.readTable(backendSheet);
    SheetUtils.writeTable(frontendSheet, data.rows);
  }

  export function syncByBackendSheetName(name: string) {
    const mapping = MAPPINGS.find((m) => m.backend === name);
    if (!mapping) return;
    copyTable(mapping.backend, mapping.frontend);
  }

  export function syncAllMapped() {
    MAPPINGS.forEach((m) => copyTable(m.backend, m.frontend));
  }
}