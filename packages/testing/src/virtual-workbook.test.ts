import { createExcelTools } from "@lucamattiazzi/sommelier-excel";
import { describe, expect, it } from "vitest";
import { VirtualWorkbookDriver } from "./virtual-workbook.js";

const fixture = {
  selection: { worksheet: "Forecast", range: "A1:B2" },
  worksheets: [
    {
      name: "Forecast",
      values: [
        ["Product", "Q1"],
        ["Widget", 10],
      ],
      tables: [{ name: "ForecastTable", range: "A1:B2", hasHeaders: true }],
      formulas: {},
      formats: {},
    },
  ],
  namedItems: [],
};

describe("VirtualWorkbookDriver", () => {
  it("supports every mutation used by MVP tools", async () => {
    const driver = new VirtualWorkbookDriver(fixture);
    await driver.writeRange("Forecast", "B2:B2", [[20]]);
    await driver.setFormulas("Forecast", "C2:C2", [["=SUM(B2:B2)"]]);
    await driver.formatRange("Forecast", "A1:C1", { bold: true, fillColor: "#336699" });
    await driver.createTable("Forecast", "A1:C2", "NewForecastTable", true);
    await driver.addComment("Forecast", "B2", "Synthetic note");
    expect((await driver.readRange("Forecast", "B2:C2")).values[0]?.[0]).toBe(20);
    expect(driver.getFormula("Forecast", "C2")).toBe("=SUM(B2:B2)");
    expect(driver.getFormat("Forecast", "A1")).toMatchObject({ bold: true });
    expect(driver.hasTable("NewForecastTable")).toBe(true);
  });

  it("creates all nine validated Excel tools", () => {
    expect(createExcelTools({ driver: new VirtualWorkbookDriver(fixture) })).toHaveLength(9);
  });

  it("executes every MVP tool through the shared driver contract", async () => {
    const driver = new VirtualWorkbookDriver(fixture);
    const tools = new Map(createExcelTools({ driver }).map((tool) => [tool.name, tool]));
    const context = {
      signal: new AbortController().signal,
      sessionId: "session",
      turnId: "turn",
      toolCallId: "call",
    };
    const inputs: Record<string, unknown> = {
      "excel.get_selection": {},
      "excel.get_workbook_structure": {},
      "excel.read_range": { worksheet: "Forecast", range: "A1:B2" },
      "excel.read_table": { name: "ForecastTable", offset: 0, limit: 10 },
      "excel.write_range": { worksheet: "Forecast", range: "B2:B2", values: [[20]] },
      "excel.set_formulas": {
        worksheet: "Forecast",
        range: "C2:C2",
        formulas: [["=SUM(B2:B2)"]],
      },
      "excel.create_table": {
        worksheet: "Forecast",
        range: "A1:C2",
        name: "ExpandedForecast",
        hasHeaders: true,
      },
      "excel.format_range": {
        worksheet: "Forecast",
        range: "A1:C1",
        format: { bold: true },
      },
      "excel.add_comment": {
        worksheet: "Forecast",
        cell: "B2",
        comment: "Synthetic note",
      },
    };
    for (const [name, input] of Object.entries(inputs)) {
      const tool = tools.get(name);
      if (!tool) throw new Error(`Missing ${name}`);
      await expect(tool.execute(tool.input.parse(input), context)).resolves.toBeDefined();
    }
    expect(driver.hasTable("ExpandedForecast")).toBe(true);
    expect(driver.getFormula("Forecast", "C2")).toBe("=SUM(B2:B2)");
  });
});
