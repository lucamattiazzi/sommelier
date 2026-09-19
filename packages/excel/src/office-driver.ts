/// <reference types="office-js" />

import { SommelierError } from "@lucamattiazzi/sommelier-core";
import type {
  CellValue,
  RangeData,
  RangeFormat,
  TableData,
  WorkbookCapabilities,
  WorkbookDriver,
  WorkbookStructure,
} from "./driver.js";

function abortIfNeeded(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason;
}

function valuesOf(values: unknown[][]): CellValue[][] {
  return values.map((row) =>
    row.map((value) =>
      typeof value === "string" || typeof value === "number" || typeof value === "boolean"
        ? value
        : null,
    ),
  );
}

function rangeData(range: Excel.Range, worksheet: string): RangeData {
  return {
    worksheet,
    range: range.address.split("!").pop()?.replaceAll("$", "") ?? range.address,
    values: valuesOf(range.values),
    formulas: valuesOf(range.formulas),
    valueTypes: range.valueTypes.map((row) => row.map(String)),
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}

/** Office.js implementation that operates on the workbook open in the current task pane. */
export class OfficeJsWorkbookDriver implements WorkbookDriver {
  async getSelection(signal?: AbortSignal): Promise<RangeData> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load("address,values,formulas,valueTypes,rowCount,columnCount");
      selected.worksheet.load("name");
      await context.sync();
      abortIfNeeded(signal);
      return rangeData(selected, selected.worksheet.name);
    });
  }

  async getWorkbookStructure(signal?: AbortSignal): Promise<WorkbookStructure> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      const names = context.workbook.names;
      worksheets.load("items/name");
      names.load("items/name,items/type");
      await context.sync();
      const entries = worksheets.items.map((sheet) => {
        const used = sheet.getUsedRangeOrNullObject(true);
        const tables = sheet.tables;
        used.load("address,isNullObject");
        tables.load("items/name");
        return { sheet, used, tables };
      });
      await context.sync();
      const tableRanges = entries.flatMap(({ tables }) =>
        tables.items.map((table) => {
          const tableRange = table.getRange();
          tableRange.load("address");
          return { table, tableRange };
        }),
      );
      await context.sync();
      abortIfNeeded(signal);
      return {
        worksheets: entries.map(({ sheet, used, tables }) => ({
          name: sheet.name,
          ...(used.isNullObject
            ? {}
            : { usedRange: used.address.split("!").pop()?.replaceAll("$", "") ?? used.address }),
          tables: tables.items.map((table) => {
            const found = tableRanges.find((entry) => entry.table === table);
            return {
              name: table.name,
              range: found?.tableRange.address.split("!").pop()?.replaceAll("$", "") ?? "",
            };
          }),
        })),
        namedItems: names.items.map((item) => ({ name: item.name, type: String(item.type) })),
      };
    });
  }

  async readRange(worksheet: string, address: string, signal?: AbortSignal): Promise<RangeData> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const selected = context.workbook.worksheets.getItem(worksheet).getRange(address);
      selected.load("address,values,formulas,valueTypes,rowCount,columnCount");
      await context.sync();
      abortIfNeeded(signal);
      return rangeData(selected, worksheet);
    });
  }

  async readTable(
    name: string,
    offset: number,
    limit: number,
    signal?: AbortSignal,
  ): Promise<TableData> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(name);
      const header = table.getHeaderRowRange();
      const body = table.getDataBodyRange();
      const full = table.getRange();
      table.load("name");
      table.worksheet.load("name");
      header.load("values");
      body.load("values,rowCount,columnCount");
      full.load("address");
      await context.sync();
      abortIfNeeded(signal);
      return {
        worksheet: table.worksheet.name,
        name: table.name,
        range: full.address.split("!").pop()?.replaceAll("$", "") ?? full.address,
        headers: valuesOf(header.values)[0]?.map(String) ?? [],
        rows: valuesOf(body.values).slice(offset, offset + limit),
      };
    });
  }

  async writeRange(
    worksheet: string,
    address: string,
    values: readonly (readonly CellValue[])[],
    signal?: AbortSignal,
  ): Promise<void> {
    abortIfNeeded(signal);
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem(worksheet).getRange(address).values = values.map(
        (row) => [...row],
      );
      await context.sync();
      abortIfNeeded(signal);
    });
  }

  async setFormulas(
    worksheet: string,
    address: string,
    formulas: readonly (readonly string[])[],
    signal?: AbortSignal,
  ): Promise<void> {
    abortIfNeeded(signal);
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem(worksheet).getRange(address).formulas = formulas.map(
        (row) => [...row],
      );
      await context.sync();
      abortIfNeeded(signal);
    });
  }

  async createTable(
    worksheet: string,
    address: string,
    name: string,
    hasHeaders: boolean,
    signal?: AbortSignal,
  ): Promise<void> {
    abortIfNeeded(signal);
    await Excel.run(async (context) => {
      const table = context.workbook.worksheets.getItem(worksheet).tables.add(address, hasHeaders);
      table.name = name;
      await context.sync();
      abortIfNeeded(signal);
    });
  }

  async formatRange(
    worksheet: string,
    address: string,
    format: RangeFormat,
    signal?: AbortSignal,
  ): Promise<void> {
    abortIfNeeded(signal);
    await Excel.run(async (context) => {
      const target = context.workbook.worksheets.getItem(worksheet).getRange(address);
      if (format.bold !== undefined) target.format.font.bold = format.bold;
      if (format.italic !== undefined) target.format.font.italic = format.italic;
      if (format.fillColor !== undefined) target.format.fill.color = format.fillColor;
      if (format.fontColor !== undefined) target.format.font.color = format.fontColor;
      if (format.horizontalAlignment !== undefined) {
        target.format.horizontalAlignment = {
          left: Excel.HorizontalAlignment.left,
          center: Excel.HorizontalAlignment.center,
          right: Excel.HorizontalAlignment.right,
        }[format.horizontalAlignment];
      }
      if (format.numberFormat !== undefined) {
        target.load("rowCount,columnCount");
        await context.sync();
        target.numberFormat = Array.from({ length: target.rowCount }, () =>
          Array.from({ length: target.columnCount }, () => format.numberFormat ?? "General"),
        );
      }
      if (format.autofit) {
        target.format.autofitColumns();
        target.format.autofitRows();
      }
      await context.sync();
      abortIfNeeded(signal);
    });
  }

  async addComment(
    worksheet: string,
    cell: string,
    comment: string,
    signal?: AbortSignal,
  ): Promise<void> {
    abortIfNeeded(signal);
    const capabilities = await this.getCapabilities(signal);
    if (!capabilities.comments) {
      throw new SommelierError({
        code: "SOMMELIER_EXCEL_CAPABILITY_UNSUPPORTED",
        message: "This Excel host does not support the comments requirement set.",
      });
    }
    await Excel.run(async (context) => {
      context.workbook.comments.add(`${worksheet}!${cell}`, comment);
      await context.sync();
      abortIfNeeded(signal);
    });
  }

  async getCapabilities(_signal?: AbortSignal): Promise<WorkbookCapabilities> {
    const requirements = globalThis.Office?.context?.requirements;
    return {
      comments: requirements?.isSetSupported("ExcelApi", "1.10") ?? false,
      requirementSets: {
        ExcelApi: requirements?.isSetSupported("ExcelApi", "1.13") ? "1.13" : "unknown",
      },
    };
  }
}
