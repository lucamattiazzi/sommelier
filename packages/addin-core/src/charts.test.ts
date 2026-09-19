import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import { PROTOCOL_VERSION, type ProtocolRequest } from "@lucamattiazzi/sommelier-protocol";
import { describe, expect, it, vi } from "vitest";
import { createAddinController } from "./index.js";

describe("chart operations", () => {
  it("previews without changes, requires approval, creates a named chart and lists it", async () => {
    const adapter = new InMemoryExcelAdapter({
      sheets: [
        {
          name: "Sales",
          values: [
            ["Month", "Sales"],
            ["Jan", 12],
          ],
        },
      ],
    });
    const approval = vi.fn(async () => true);
    const controller = createAddinController({
      adapter,
      workbook: { id: "test", name: "Test" },
      requestApproval: approval,
    });
    await controller.start();
    const call = (method: string, params: unknown) =>
      controller.handle({
        protocolVersion: PROTOCOL_VERSION,
        type: "request",
        id: crypto.randomUUID(),
        method,
        params,
      } as ProtocolRequest);
    const params = {
      range: { sheetId: "Sales", address: "A1:B2" },
      name: "Monthly",
      title: "Monthly sales",
      chartType: "column",
    };
    const preview = await call("excel.operation.preview", { method: "excel.chart.create", params });
    expect(preview.type).toBe("response");
    if (preview.type !== "response") throw new Error("Preview failed");
    expect(await call("excel.chart.list", { sheetId: "Sales" })).toMatchObject({
      result: { charts: [] },
    });
    expect(
      await call("excel.operation.commit", {
        operationId: (preview.result as { operationId: string }).operationId,
      }),
    ).toMatchObject({ result: { status: "committed" } });
    expect(approval).toHaveBeenCalledWith(
      expect.objectContaining({
        method: "excel.chart.create",
        summary: expect.stringContaining("Monthly"),
      }),
    );
    expect(await call("excel.chart.list", { sheetId: "Sales" })).toMatchObject({
      result: { charts: [{ name: "Monthly", title: "Monthly sales" }] },
    });
    expect((await call("excel.chart.create", params)).type).toBe("error");
    await controller.stop();
  });
  it("does not create a chart without approval or with an oversized source", async () => {
    const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sales" }] });
    const create = vi.spyOn(adapter, "createChart");
    const controller = createAddinController({ adapter, workbook: { id: "test", name: "Test" } });
    await controller.start();
    for (const address of ["A1:B2", "A1:Z10000"]) {
      expect(
        (
          await controller.handle({
            protocolVersion: PROTOCOL_VERSION,
            type: "request",
            id: "1",
            method: "excel.chart.create",
            params: {
              range: { sheetId: "Sales", address },
              name: "Chart",
              title: "Chart",
              chartType: "line",
            },
          } as ProtocolRequest)
        ).type,
      ).toBe("error");
    }
    expect(create).not.toHaveBeenCalled();
    await controller.stop();
  });
});
