// eslint-disable-next-line @typescript-eslint/no-explicit-any
var Shamrock: any = (this as any).Shamrock || ((this as any).Shamrock = {});

Shamrock.getFrontendSpreadsheetId = function (): string {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(Shamrock.PROPERTY_KEYS.frontendSpreadsheetId);
  if (!id) {
    throw new Error("Frontend spreadsheet ID is not configured. Use the SHAMROCK menu to set it.");
  }
  return id;
};

Shamrock.setFrontendSpreadsheetId = function (id: string): void {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(Shamrock.PROPERTY_KEYS.frontendSpreadsheetId, id.trim());
};

Shamrock.getBackendSpreadsheetIdSafe = function (): string | null {
  try {
    return Shamrock.getBackendSpreadsheetId();
  } catch (err) {
    return null;
  }
};

Shamrock.getBackendSpreadsheetId = function (): string {
  const props = PropertiesService.getScriptProperties();
  const propId = props.getProperty(Shamrock.PROPERTY_KEYS.backendSpreadsheetId);
  const globalObj: any = typeof globalThis !== "undefined" ? (globalThis as any) : (typeof this !== "undefined" ? (this as any) : {});
  // Prefer script properties, then explicit globals, then placeholder fallback detection
  const fromConst = typeof SHAMROCK_BACKEND_ID !== "undefined" ? (SHAMROCK_BACKEND_ID as string) : null;
  const fromGlobals = (globalObj && (globalObj.SHAMROCK_BACKEND_ID || globalObj.SHAMROCK_BACKEND_SPREADSHEET_ID)) || null;
  const candidate = propId || fromConst || fromGlobals || null;
  if (!candidate || candidate === "SHAMROCK_BACKEND_SPREADSHEET_ID") {
    throw new Error("Backend spreadsheet ID is not configured. Set it via script properties or SHAMROCK_BACKEND_ID.");
  }
  return candidate;
};

Shamrock.setBackendSpreadsheetId = function (id: string): void {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(Shamrock.PROPERTY_KEYS.backendSpreadsheetId, id.trim());
};
