import { SommelierError } from "@lucamattiazzi/sommelier-core";
import {
  type CellValue,
  parseA1Range,
  type RangeData,
  type RangeFormat,
  type TableData,
  toA1Cell,
  type WorkbookCapabilities,
  type WorkbookDriver,
  type WorkbookStructure,
} from "@lucamattiazzi/sommelier-excel";
import { z } from "zod";

const cellSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);
/** Runtime schema for synthetic virtual-workbook fixtures. */
export const virtualWorkbookFixtureSchema = z.object({
  selection: z.object({ worksheet: z.string(), range: z.string() }).optional(),
  worksheets: z
    .array(
      z.object({
        name: z.string().min(1),
        values: z.array(z.array(cellSchema)).default([]),
        formulas: z.record(z.string(), z.string().startsWith("=")).default({}),
        formats: z.record(z.string(), z.record(z.string(), z.unknown())).default({}),
        tables: z
          .array(
            z.object({
              name: z.string(),
              range: z.string(),
              hasHeaders: z.boolean().default(true),
            }),
          )
          .default([]),
      }),
    )
    .min(1),
  namedItems: z.array(z.object({ name: z.string(), type: z.string() })).default([]),
});

/** Serializable synthetic workbook fixture. */
export type VirtualWorkbookFixture = z.infer<typeof virtualWorkbookFixtureSchema>;

interface SheetState {
  readonly name: string;
  readonly values: CellValue[][];
  readonly formulas: Map<string, string>;
  readonly formats: Map<string, RangeFormat>;
  readonly tables: Map<string, { range: string; hasHeaders: boolean }>;
  readonly comments: Map<string, string>;
}

/** Explicit error for Excel behaviors intentionally absent from the virtual driver. */
export class UnsupportedVirtualWorkbookOperationError extends SommelierError {
  constructor(operation: string) {
    super({
      code: "SOMMELIER_VIRTUAL_OPERATION_UNSUPPORTED",
      message: `The virtual workbook does not support ${operation}.`,
      context: { operation },
      suggestedAction: "Use only documented MVP tools or run an Office.js integration test.",
    });
    this.name = "UnsupportedVirtualWorkbookOperationError";
  }
}

/** In-memory driver faithfully implementing all MVP Excel tool operations. */
export class VirtualWorkbookDriver implements WorkbookDriver {
  readonly #sheets = new Map<string, SheetState>();
  readonly #namedItems: readonly { name: string; type: string }[];
  #selection: { worksheet: string; range: string };

  constructor(input: VirtualWorkbookFixture) {
    const fixture = virtualWorkbookFixtureSchema.parse(input);
    for (const sheet of fixture.worksheets) {
      this.#sheets.set(sheet.name, {
        name: sheet.name,
        values: sheet.values.map((row) => [...row]),
        formulas: new Map(Object.entries(sheet.formulas)),
        formats: new Map(
          Object.entries(sheet.formats).map(([cell, value]) => [cell, value as RangeFormat]),
        ),
        tables: new Map(
          sheet.tables.map((table) => [
            table.name,
            { range: parseA1Range(table.range).address, hasHeaders: table.hasHeaders },
          ]),
        ),
        comments: new Map(),
      });
    }
    this.#namedItems = fixture.namedItems;
    this.#selection = fixture.selection ?? {
      worksheet: fixture.worksheets[0]?.name ?? "Sheet1",
      range: "A1:A1",
    };
  }

  /** Change the selection returned by `excel.get_selection`. */
  setSelection(worksheet: string, range: string): void {
    this.#sheet(worksheet);
    this.#selection = { worksheet, range: parseA1Range(range).address };
  }

  async getSelection(): Promise<RangeData> {
    return this.readRange(this.#selection.worksheet, this.#selection.range);
  }

  async getWorkbookStructure(): Promise<WorkbookStructure> {
    return {
      worksheets: [...this.#sheets.values()].map((sheet) => {
        const usedRange = this.#usedRange(sheet);
        return {
          name: sheet.name,
          ...(usedRange === undefined ? {} : { usedRange }),
          tables: [...sheet.tables].map(([name, table]) => ({ name, range: table.range })),
        };
      }),
      namedItems: this.#namedItems,
    };
  }

  async readRange(worksheet: string, address: string): Promise<RangeData> {
    const sheet = this.#sheet(worksheet);
    const parsed = parseA1Range(address);
    const values: CellValue[][] = [];
    const formulas: CellValue[][] = [];
    const valueTypes: string[][] = [];
    for (let row = parsed.startRow; row <= parsed.endRow; row += 1) {
      const valueRow: CellValue[] = [];
      const formulaRow: CellValue[] = [];
      const typeRow: string[] = [];
      for (let column = parsed.startColumn; column <= parsed.endColumn; column += 1) {
        const cell = toA1Cell(row, column);
        const value = sheet.values[row]?.[column] ?? null;
        const formula = sheet.formulas.get(cell);
        valueRow.push(value);
        formulaRow.push(formula ?? value);
        typeRow.push(formula ? "Formula" : value === null ? "Empty" : typeof value);
      }
      values.push(valueRow);
      formulas.push(formulaRow);
      valueTypes.push(typeRow);
    }
    return {
      worksheet,
      range: parsed.address,
      values,
      formulas,
      valueTypes,
      rowCount: parsed.rowCount,
      columnCount: parsed.columnCount,
    };
  }

  async readTable(name: string, offset: number, limit: number): Promise<TableData> {
    for (const sheet of this.#sheets.values()) {
      const table = sheet.tables.get(name);
      if (!table) continue;
      const data = await this.readRange(sheet.name, table.range);
      const headers = table.hasHeaders
        ? (data.values[0] ?? []).map(String)
        : (data.values[0]?.map((_, index) => `Column${index + 1}`) ?? []);
      const rows = data.values.slice(
        table.hasHeaders ? 1 + offset : offset,
        (table.hasHeaders ? 1 : 0) + offset + limit,
      );
      return { worksheet: sheet.name, name, range: table.range, headers, rows };
    }
    throw new SommelierError({
      code: "SOMMELIER_EXCEL_TABLE_NOT_FOUND",
      message: `Table ${name} does not exist.`,
    });
  }

  async writeRange(
    worksheet: string,
    address: string,
    values: readonly (readonly CellValue[])[],
  ): Promise<void> {
    const sheet = this.#sheet(worksheet);
    const parsed = parseA1Range(address);
    this.#ensureSize(sheet, parsed.endRow, parsed.endColumn);
    for (let row = 0; row < parsed.rowCount; row += 1) {
      for (let column = 0; column < parsed.columnCount; column += 1) {
        const targetRow = parsed.startRow + row;
        const targetColumn = parsed.startColumn + column;
        const target = sheet.values[targetRow];
        if (target) target[targetColumn] = values[row]?.[column] ?? null;
        sheet.formulas.delete(toA1Cell(targetRow, targetColumn));
      }
    }
  }

  async setFormulas(
    worksheet: string,
    address: string,
    formulas: readonly (readonly string[])[],
  ): Promise<void> {
    const sheet = this.#sheet(worksheet);
    const parsed = parseA1Range(address);
    this.#ensureSize(sheet, parsed.endRow, parsed.endColumn);
    for (let row = 0; row < parsed.rowCount; row += 1) {
      for (let column = 0; column < parsed.columnCount; column += 1) {
        const targetRow = parsed.startRow + row;
        const targetColumn = parsed.startColumn + column;
        sheet.formulas.set(toA1Cell(targetRow, targetColumn), formulas[row]?.[column] ?? "=");
      }
    }
  }

  async createTable(
    worksheet: string,
    address: string,
    name: string,
    hasHeaders: boolean,
  ): Promise<void> {
    if ([...this.#sheets.values()].some((sheet) => sheet.tables.has(name))) {
      throw new SommelierError({
        code: "SOMMELIER_EXCEL_TABLE_CONFLICT",
        message: `Table ${name} already exists.`,
      });
    }
    this.#sheet(worksheet).tables.set(name, { range: parseA1Range(address).address, hasHeaders });
  }

  async formatRange(worksheet: string, address: string, format: RangeFormat): Promise<void> {
    const sheet = this.#sheet(worksheet);
    const parsed = parseA1Range(address);
    for (let row = parsed.startRow; row <= parsed.endRow; row += 1) {
      for (let column = parsed.startColumn; column <= parsed.endColumn; column += 1) {
        const cell = toA1Cell(row, column);
        sheet.formats.set(cell, { ...sheet.formats.get(cell), ...format });
      }
    }
  }

  async addComment(worksheet: string, cell: string, comment: string): Promise<void> {
    this.#sheet(worksheet).comments.set(parseA1Range(cell).address.split(":")[0] ?? cell, comment);
  }

  async getCapabilities(): Promise<WorkbookCapabilities> {
    return { comments: true, requirementSets: { ExcelApi: "virtual-mvp" } };
  }

  /** Read a formula for deterministic assertions. */
  getFormula(worksheet: string, cell: string): string | undefined {
    return this.#sheet(worksheet).formulas.get(cell.toUpperCase());
  }

  /** Read the merged supported format for one cell. */
  getFormat(worksheet: string, cell: string): RangeFormat | undefined {
    return this.#sheet(worksheet).formats.get(cell.toUpperCase());
  }

  /** Determine whether a named table exists. */
  hasTable(name: string): boolean {
    return [...this.#sheets.values()].some((sheet) => sheet.tables.has(name));
  }

  #sheet(name: string): SheetState {
    const sheet = this.#sheets.get(name);
    if (!sheet)
      throw new SommelierError({
        code: "SOMMELIER_EXCEL_WORKSHEET_NOT_FOUND",
        message: `Worksheet ${name} does not exist.`,
      });
    return sheet;
  }

  #ensureSize(sheet: SheetState, endRow: number, endColumn: number): void {
    while (sheet.values.length <= endRow) sheet.values.push([]);
    for (const row of sheet.values) while (row.length <= endColumn) row.push(null);
  }

  #usedRange(sheet: SheetState): string | undefined {
    const lastRow = sheet.values.reduce(
      (maximum, row, index) => (row.some((cell) => cell !== null) ? index : maximum),
      -1,
    );
    const lastColumn = sheet.values.reduce<number>(
      (maximum, row) =>
        Math.max(
          maximum,
          row.reduce<number>((inner, cell, index) => (cell !== null ? index : inner), -1),
        ),
      -1,
    );
    return lastRow < 0 || lastColumn < 0 ? undefined : `A1:${toA1Cell(lastRow, lastColumn)}`;
  }
}
