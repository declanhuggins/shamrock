// Configuration helpers for Script Properties and resource naming.

namespace Config {
  export const PROPERTY_KEYS = {
    FRONTEND_SHEET_ID: 'FRONTEND_SHEET_ID',
    BACKEND_SHEET_ID: 'BACKEND_SHEET_ID',
    ATTENDANCE_FORM_ID: 'ATTENDANCE_FORM_ID',
    EXCUSAL_FORM_ID: 'EXCUSAL_FORM_ID',
    DIRECTORY_FORM_ID: 'DIRECTORY_FORM_ID',
    CADET_CSV_FILE_ID: 'CADET_CSV_FILE_ID',
  } as const;

  export const RESOURCE_NAMES = {
    FRONTEND_SPREADSHEET: 'SHAMROCK Frontend',
    BACKEND_SPREADSHEET: 'SHAMROCK Backend',
    ATTENDANCE_FORM: 'SHAMROCK Attendance Form',
    EXCUSAL_FORM: 'SHAMROCK Excusal Form',
    DIRECTORY_FORM: 'SHAMROCK Directory Form',
    ATTENDANCE_FORM_SHEET: 'Attendance Form Responses',
    EXCUSAL_FORM_SHEET: 'Excusal Form Responses',
    DIRECTORY_FORM_SHEET: 'Directory Form Responses',
  } as const;

  export function scriptProperties(): GoogleAppsScript.Properties.Properties {
    return PropertiesService.getScriptProperties();
  }
}
