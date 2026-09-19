import {
  defineTool,
  type JsonValue,
  SommelierError,
  type ToolDefinition,
} from "@lucamattiazzi/sommelier-core";
import { z } from "zod";
import type { CellValue, ExcelToolResultBase, RangeFormat, WorkbookDriver } from "./driver.js";
import { assertMatrixShape, parseA1Range } from "./range.js";

/** Per-tool bounds applied before the workbook driver is called. */
export interface ExcelToolLimits {
  readonly maxCellsPerRead: number;
  readonly maxCellsPerWrite: number;
  readonly maxBytesPerResult: number;
  readonly maxPreviewCells: number;
}

/** Options for creating the standard Excel tool set. */
export interface CreateExcelToolsOptions {
  readonly driver: WorkbookDriver;
  readonly limits?: Partial<ExcelToolLimits>;
}

const defaults: ExcelToolLimits = {
  maxCellsPerRead: 10_000,
  maxCellsPerWrite: 1_000,
  maxBytesPerResult: 1_000_000,
  maxPreviewCells: 20,
};
const worksheet = z.string().trim().min(1).max(255);
const range = z.string().trim().min(2).max(50);
const cellValue = z.union([z.string(), z.number().finite(), z.boolean(), z.null()]);
const valuesMatrix = z.array(z.array(cellValue).min(1)).min(1);
const formulasMatrix = z.array(z.array(z.string().startsWith("=")).min(1)).min(1);
const tableName = z
  .string()
  .regex(/^[A-Za-z_\\][A-Za-z0-9_.]*$/, "Use a valid Excel table name.")
  .max(255);
const color = z.string().regex(/^#[0-9A-F]{6}$/i, "Use a six-digit hex color such as #336699.");

function base(
  summary: string,
  details: JsonValue,
  metadata: {
    worksheet?: string;
    range?: string;
    cellCount: number;
    message?: string;
    warnings?: readonly string[];
  },
): ExcelToolResultBase {
  return {
    summary,
    ...(metadata.worksheet === undefined ? {} : { worksheet: metadata.worksheet }),
    ...(metadata.range === undefined ? {} : { range: metadata.range }),
    cellCount: metadata.cellCount,
    truncated: false,
    warnings: metadata.warnings ?? [],
    capabilityLimitations: [],
    message: metadata.message ?? summary,
    details,
  };
}

function enforce(count: number, maximum: number, operation: string): void {
  if (count > maximum) {
    throw new SommelierError({
      code: "SOMMELIER_EXCEL_CELL_LIMIT_EXCEEDED",
      message: `${operation} targets ${count} cells; the configured limit is ${maximum}.`,
      context: { cellCount: count, limit: maximum, operation },
      suggestedAction: "Request a smaller explicit range or deliberately raise the host limit.",
    });
  }
}

function boundedResult(result: ExcelToolResultBase, maximum: number): ExcelToolResultBase {
  const bytes = new TextEncoder().encode(JSON.stringify(result)).byteLength;
  if (bytes <= maximum) return result;
  throw new SommelierError({
    code: "SOMMELIER_EXCEL_BYTE_LIMIT_EXCEEDED",
    message: `The workbook result is ${bytes} bytes; the configured limit is ${maximum}.`,
    context: { bytes, limit: maximum },
    suggestedAction: "Read a smaller range or table subset.",
  });
}

/** Create the complete bounded MVP Excel tool set for a workbook driver. */
export function createExcelTools(options: CreateExcelToolsOptions): readonly ToolDefinition[] {
  const driver = options.driver;
  const limits = { ...defaults, ...options.limits };
  const rangedInput = z.object({ worksheet, range });

  return [
    defineTool({
      name: "excel.get_selection",
      description: "Read the currently selected bounded Excel range.",
      classification: "read",
      input: z.object({}),
      execute: async (_, context) => {
        const data = await driver.getSelection(context.signal);
        enforce(data.rowCount * data.columnCount, limits.maxCellsPerRead, "get selection");
        return boundedResult(
          base("Read current selection", data as unknown as JsonValue, {
            worksheet: data.worksheet,
            range: data.range,
            cellCount: data.rowCount * data.columnCount,
          }),
          limits.maxBytesPerResult,
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.get_workbook_structure",
      description:
        "Inspect worksheet, used-range, table, and named-item metadata without cell contents.",
      classification: "read",
      input: z.object({}),
      execute: async (_, context) => {
        const structure = await driver.getWorkbookStructure(context.signal);
        return boundedResult(
          base("Read workbook structure", structure as unknown as JsonValue, { cellCount: 0 }),
          limits.maxBytesPerResult,
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.read_range",
      description: "Read values, formulas, and value types from an explicit bounded A1 range.",
      classification: "read",
      input: rangedInput,
      estimateCells: (input) => ({ read: parseA1Range(input.range).cellCount, written: 0 }),
      execute: async (input, context) => {
        const parsed = parseA1Range(input.range);
        enforce(parsed.cellCount, limits.maxCellsPerRead, "read range");
        const data = await driver.readRange(input.worksheet, parsed.address, context.signal);
        return boundedResult(
          base("Read range", data as unknown as JsonValue, {
            worksheet: data.worksheet,
            range: data.range,
            cellCount: parsed.cellCount,
          }),
          limits.maxBytesPerResult,
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.read_table",
      description: "Read a named table or a bounded row subset, including headers.",
      classification: "read",
      input: z.object({
        name: tableName,
        offset: z.number().int().min(0).default(0),
        limit: z.number().int().min(1).max(1000).default(100),
      }),
      execute: async (input, context) => {
        const data = await driver.readTable(input.name, input.offset, input.limit, context.signal);
        const count = data.rows.length * data.headers.length;
        enforce(count, limits.maxCellsPerRead, "read table");
        return boundedResult(
          base("Read table", data as unknown as JsonValue, {
            worksheet: data.worksheet,
            range: data.range,
            cellCount: count,
          }),
          limits.maxBytesPerResult,
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.write_range",
      description: "Write a rectangular value matrix to an explicit existing range.",
      classification: "write",
      input: rangedInput.extend({ values: valuesMatrix }),
      estimateCells: (input) => ({ read: 0, written: parseA1Range(input.range).cellCount }),
      preview: async (input) => {
        const parsed = parseA1Range(input.range);
        assertMatrixShape(input.values, parsed, "values");
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "write range");
        return {
          summary: `Write ${parsed.cellCount} cells`,
          target: `${input.worksheet}!${parsed.address}`,
          affectedCells: parsed.cellCount,
          sample: input.values.slice(0, limits.maxPreviewCells) as JsonValue,
          warnings: ["Existing non-empty cells may be overwritten."],
        };
      },
      execute: async (input, context) => {
        const parsed = parseA1Range(input.range);
        assertMatrixShape(input.values, parsed, "values");
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "write range");
        await driver.writeRange(
          input.worksheet,
          parsed.address,
          input.values as CellValue[][],
          context.signal,
        );
        return base(
          "Wrote values",
          { written: true },
          { worksheet: input.worksheet, range: parsed.address, cellCount: parsed.cellCount },
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.set_formulas",
      description: "Set a rectangular matrix of formulas beginning with = in an explicit range.",
      classification: "write",
      input: rangedInput.extend({ formulas: formulasMatrix }),
      estimateCells: (input) => ({ read: 0, written: parseA1Range(input.range).cellCount }),
      preview: async (input) => {
        const parsed = parseA1Range(input.range);
        assertMatrixShape(input.formulas, parsed, "formulas");
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "set formulas");
        return {
          summary: `Set ${parsed.cellCount} formulas`,
          target: `${input.worksheet}!${parsed.address}`,
          affectedCells: parsed.cellCount,
          sample: input.formulas.slice(0, limits.maxPreviewCells) as JsonValue,
          warnings: ["Existing cell values may be overwritten."],
        };
      },
      execute: async (input, context) => {
        const parsed = parseA1Range(input.range);
        assertMatrixShape(input.formulas, parsed, "formulas");
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "set formulas");
        await driver.setFormulas(input.worksheet, parsed.address, input.formulas, context.signal);
        return base(
          "Set formulas",
          { written: true },
          { worksheet: input.worksheet, range: parsed.address, cellCount: parsed.cellCount },
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.create_table",
      description: "Create a named Excel table from an explicit existing range.",
      classification: "structural",
      input: rangedInput.extend({ name: tableName, hasHeaders: z.boolean().default(true) }),
      estimateCells: (input) => ({ read: 0, written: parseA1Range(input.range).cellCount }),
      preview: async (input) => ({
        summary: `Create table ${input.name}`,
        target: `${input.worksheet}!${parseA1Range(input.range).address}`,
        affectedCells: parseA1Range(input.range).cellCount,
        warnings: [],
      }),
      execute: async (input, context) => {
        const parsed = parseA1Range(input.range);
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "create table");
        await driver.createTable(
          input.worksheet,
          parsed.address,
          input.name,
          input.hasHeaders,
          context.signal,
        );
        return base(
          "Created table",
          { name: input.name },
          { worksheet: input.worksheet, range: parsed.address, cellCount: parsed.cellCount },
        ) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.format_range",
      description: "Apply an allowlisted formatting subset to an explicit range.",
      classification: "write",
      input: rangedInput.extend({
        format: z
          .object({
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            fillColor: color.optional(),
            fontColor: color.optional(),
            numberFormat: z.string().max(100).optional(),
            horizontalAlignment: z.enum(["left", "center", "right"]).optional(),
            autofit: z.boolean().optional(),
          })
          .refine(
            (value) => Object.keys(value).length > 0,
            "Specify at least one format property.",
          ),
      }),
      estimateCells: (input) => ({ read: 0, written: parseA1Range(input.range).cellCount }),
      preview: async (input) => ({
        summary: "Format range",
        target: `${input.worksheet}!${parseA1Range(input.range).address}`,
        affectedCells: parseA1Range(input.range).cellCount,
        sample: input.format as JsonValue,
        warnings: [],
      }),
      execute: async (input, context) => {
        const parsed = parseA1Range(input.range);
        enforce(parsed.cellCount, limits.maxCellsPerWrite, "format range");
        await driver.formatRange(
          input.worksheet,
          parsed.address,
          input.format as RangeFormat,
          context.signal,
        );
        return base("Formatted range", input.format as JsonValue, {
          worksheet: input.worksheet,
          range: parsed.address,
          cellCount: parsed.cellCount,
        }) as unknown as JsonValue;
      },
    }),
    defineTool({
      name: "excel.add_comment",
      description: "Add a comment to one cell when the current Excel host supports comments.",
      classification: "write",
      input: z.object({ worksheet, cell: range, comment: z.string().min(1).max(10_000) }),
      estimateCells: () => ({ read: 0, written: 1 }),
      preview: async (input) => ({
        summary: "Add comment",
        target: `${input.worksheet}!${parseA1Range(input.cell).address}`,
        affectedCells: 1,
        warnings: [],
      }),
      execute: async (input, context) => {
        const parsed = parseA1Range(input.cell);
        if (parsed.cellCount !== 1)
          throw new SommelierError({
            code: "SOMMELIER_EXCEL_SINGLE_CELL_REQUIRED",
            message: "Comments target exactly one cell.",
          });
        const capabilities = await driver.getCapabilities(context.signal);
        if (!capabilities.comments)
          throw new SommelierError({
            code: "SOMMELIER_EXCEL_CAPABILITY_UNSUPPORTED",
            message: "Comments are not supported by this Excel host.",
            suggestedAction: "Check the host requirement set or disable excel.add_comment.",
          });
        await driver.addComment(input.worksheet, parsed.address, input.comment, context.signal);
        return base(
          "Added comment",
          { added: true },
          { worksheet: input.worksheet, range: parsed.address, cellCount: 1 },
        ) as unknown as JsonValue;
      },
    }),
  ];
}
