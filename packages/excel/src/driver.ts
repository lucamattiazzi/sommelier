import type { JsonValue } from "@lucamattiazzi/sommelier-core";

/** Cell value accepted by bounded workbook operations. */
export type CellValue = string | number | boolean | null;

/** Supported format subset; arbitrary Office.js style objects are intentionally excluded. */
export interface RangeFormat {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly fillColor?: string;
  readonly fontColor?: string;
  readonly numberFormat?: string;
  readonly horizontalAlignment?: "left" | "center" | "right";
  readonly autofit?: boolean;
}

/** Bounded range data returned by a workbook driver. */
export interface RangeData {
  readonly worksheet: string;
  readonly range: string;
  readonly values: readonly (readonly CellValue[])[];
  readonly formulas: readonly (readonly CellValue[])[];
  readonly valueTypes: readonly (readonly string[])[];
  readonly rowCount: number;
  readonly columnCount: number;
}

/** Workbook structure without cell contents. */
export interface WorkbookStructure {
  readonly worksheets: readonly {
    readonly name: string;
    readonly usedRange?: string;
    readonly tables: readonly { readonly name: string; readonly range: string }[];
  }[];
  readonly namedItems: readonly { readonly name: string; readonly type: string }[];
}

/** Named table data. */
export interface TableData {
  readonly worksheet: string;
  readonly name: string;
  readonly range: string;
  readonly headers: readonly string[];
  readonly rows: readonly (readonly CellValue[])[];
}

/** Runtime capabilities that can differ by Excel host and requirement set. */
export interface WorkbookCapabilities {
  readonly comments: boolean;
  readonly requirementSets: Readonly<Record<string, string>>;
}

/** Replaceable workbook execution port used by all Excel tools. */
export interface WorkbookDriver {
  getSelection(signal?: AbortSignal): Promise<RangeData>;
  getWorkbookStructure(signal?: AbortSignal): Promise<WorkbookStructure>;
  readRange(worksheet: string, range: string, signal?: AbortSignal): Promise<RangeData>;
  readTable(name: string, offset: number, limit: number, signal?: AbortSignal): Promise<TableData>;
  writeRange(
    worksheet: string,
    range: string,
    values: readonly (readonly CellValue[])[],
    signal?: AbortSignal,
  ): Promise<void>;
  setFormulas(
    worksheet: string,
    range: string,
    formulas: readonly (readonly string[])[],
    signal?: AbortSignal,
  ): Promise<void>;
  createTable(
    worksheet: string,
    range: string,
    name: string,
    hasHeaders: boolean,
    signal?: AbortSignal,
  ): Promise<void>;
  formatRange(
    worksheet: string,
    range: string,
    format: RangeFormat,
    signal?: AbortSignal,
  ): Promise<void>;
  addComment(worksheet: string, cell: string, comment: string, signal?: AbortSignal): Promise<void>;
  getCapabilities(signal?: AbortSignal): Promise<WorkbookCapabilities>;
}

/** Common bounded result metadata returned by Excel tools. */
export interface ExcelToolResultBase {
  readonly summary: string;
  readonly worksheet?: string;
  readonly range?: string;
  readonly cellCount: number;
  readonly truncated: boolean;
  readonly warnings: readonly string[];
  readonly capabilityLimitations: readonly string[];
  readonly message: string;
  readonly details: JsonValue;
}
