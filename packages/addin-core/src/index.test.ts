import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import { PROTOCOL_VERSION, type ProtocolRequest } from "@lucamattiazzi/sommelier-protocol";
import { describe, expect, it } from "vitest";
import { createAddinController } from "./index.js";

function request<Method extends ProtocolRequest["method"]>(
  id: string,
  method: Method,
  params: Extract<ProtocolRequest, { method: Method }>["params"],
): Extract<ProtocolRequest, { method: Method }> {
  return { protocolVersion: PROTOCOL_VERSION, type: "request", id, method, params } as Extract<
    ProtocolRequest,
    { method: Method }
  >;
}

function setup(requestApproval?: () => boolean | Promise<boolean>) {
  const adapter = new InMemoryExcelAdapter({
    sheets: [{ name: "Sheet1", values: [["Amount"], [10]] }],
    selection: "A1",
  });
  const events: unknown[] = [];
  let id = 0;
  const controller = createAddinController({
    adapter,
    workbook: { id: "book-1", name: "Synthetic.xlsx" },
    ...(requestApproval ? { requestApproval } : {}),
    idFactory: () => `generated-${++id}`,
    timeFactory: () => new Date("2026-08-31T12:00:00.000Z"),
  });
  controller.subscribe((event) => events.push(event));
  return { adapter, controller, events };
}

describe("add-in controller", () => {
  it("reads ranges and emits lightweight bidirectional selection events", async () => {
    const { adapter, controller, events } = setup();
    await controller.start();
    const response = await controller.handle(
      request("read-1", "excel.range.read", {
        range: { sheetId: "Sheet1", address: "A1:A2" },
      }),
    );
    expect(response).toMatchObject({ type: "response", result: { values: [["Amount"], [10]] } });

    await adapter.selectRange({ worksheet: "Sheet1", range: "A2" });
    expect(events).toContainEqual(
      expect.objectContaining({
        event: "excel.selection.changed",
        data: { range: { sheetId: "Sheet1", address: "A2:A2" } },
      }),
    );
    await controller.stop();
  });

  it("blocks writes by default and executes them after explicit approval", async () => {
    const blocked = setup();
    const write = request("write-1", "excel.range.write", {
      range: { sheetId: "Sheet1", address: "A2" },
      values: [[20]],
    });
    expect(await blocked.controller.handle(write)).toMatchObject({ type: "error" });
    expect((await blocked.adapter.readRange({ worksheet: "Sheet1", range: "A2" })).values).toEqual([
      [10],
    ]);

    const allowed = setup(() => true);
    expect(await allowed.controller.handle(write)).toMatchObject({ type: "response" });
    expect((await allowed.adapter.readRange({ worksheet: "Sheet1", range: "A2" })).values).toEqual([
      [20],
    ]);
  });

  it("previews, commits, audits, and undoes a mutation", async () => {
    const { adapter, controller, events } = setup(() => true);
    const preview = await controller.handle(
      request("preview-1", "excel.operation.preview", {
        method: "excel.range.write",
        params: { range: { sheetId: "Sheet1", address: "A2" }, values: [[30]] },
      }),
    );
    expect(preview).toMatchObject({
      type: "response",
      result: { operationId: "generated-1", status: "previewed" },
    });
    expect(
      await controller.handle(
        request("commit-1", "excel.operation.commit", { operationId: "generated-1" }),
      ),
    ).toMatchObject({
      type: "response",
      result: { status: "committed" },
    });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A2" })).values).toEqual([[30]]);
    expect(
      await controller.handle(
        request("undo-1", "excel.operation.undo", { operationId: "generated-1" }),
      ),
    ).toMatchObject({
      type: "response",
      result: { status: "undone" },
    });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A2" })).values).toEqual([[10]]);
    expect(events).toEqual(
      expect.arrayContaining([expect.objectContaining({ event: "excel.audit.recorded" })]),
    );
  });
});
