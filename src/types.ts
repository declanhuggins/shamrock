// Shared types and enums for SHAMROCK

namespace Types {
  export type WorkbookRole = 'frontend' | 'backend';

  export interface TabSchema {
    name: string;
    machineHeaders?: string[];
    displayHeaders?: string[];
  }

  export interface EnsureSpreadsheetResult {
    role: WorkbookRole;
    id: string;
    name: string;
    created: boolean;
    url: string;
  }

  export interface EnsureSheetResult {
    spreadsheetId: string;
    sheetName: string;
    created: boolean;
    headersApplied: boolean;
  }

  export interface EnsureFormResult {
    kind: 'attendance' | 'excusal' | 'directory';
    id: string;
    created: boolean;
    url: string;
  }

  export interface SetupSummary {
    spreadsheets: EnsureSpreadsheetResult[];
    sheets: EnsureSheetResult[];
    forms: EnsureFormResult[];
  }
}
