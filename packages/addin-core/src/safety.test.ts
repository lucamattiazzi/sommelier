import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import {
  PROTOCOL_VERSION,
  type ProtocolRequest,
  type ProtocolResponse,
} from "@lucamattiazzi/sommelier-protocol";
import { describe, expect, it, vi } from "vitest";
import { type ApprovalRequest, createAddinController } from "./index.js";

function request(method: ProtocolRequest["method"], params: unknown): ProtocolRequest {
  return {
    protocolVersion: PROTOCOL_VERSION,
    type: "request",
    id: crypto.randomUUID(),
    method,
    params,
  } as ProtocolRequest;
}
const range = { sheetId: "Sheet1", address: "A1" };
function setup(approve: (request: ApprovalRequest) => boolean | Promise<boolean> = () => true) {
  const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1", values: [[10]] }] });
  const controller = createAddinController({
    adapter,
    workbook: { id: "test", name: "Test" },
    requestApproval: approve,
  });
  return { adapter, controller };
}
async function preview(controller: ReturnType<typeof createAddinController>) {
  const response = await controller.handle(
    request("excel.operation.preview", {
      method: "excel.range.write",
      params: { range, values: [[20]] },
    }),
  );
  return (response as ProtocolResponse<"excel.operation.preview">).result.operationId;
}

describe("Pair workbook safety", () => {
  it("does not write if the session stops during the final conflict check", async () => {
    const { adapter, controller } = setup();
    const readRange = adapter.readRange.bind(adapter);
    let reads = 0;
    vi.spyOn(adapter, "readRange").mockImplementation(async (input) => {
      const result = await readRange(input);
      if (++reads === 2) await controller.stop();
      return result;
    });
    const write = vi.spyOn(adapter, "writeRange");
    expect(
      await controller.handle(request("excel.range.write", { range, values: [[20]] })),
    ).toMatchObject({ type: "error", error: { code: "SESSION_STOPPED" } });
    expect(write).not.toHaveBeenCalled();
  });
  it("rejects excessive reads before accessing the adapter", async () => {
    const { adapter, controller } = setup();
    const read = vi.spyOn(adapter, "readRange");
    expect(
      await controller.handle(
        request("excel.range.read", { range: { ...range, address: "A1:A10001" } }),
      ),
    ).toMatchObject({ type: "error", error: { code: "RANGE_LIMIT_EXCEEDED" } });
    expect(read).not.toHaveBeenCalled();
  });
  it("rejects oversized clear requests before reading any cells", async () => {
    const { adapter, controller } = setup();
    const read = vi.spyOn(adapter, "readRange");
    expect(
      await controller.handle(
        request("excel.range.clear", {
          range: { ...range, address: "A1:A1001" },
          applyTo: "contents",
        }),
      ),
    ).toMatchObject({ type: "error", error: { code: "RANGE_LIMIT_EXCEEDED" } });
    expect(read).not.toHaveBeenCalled();
  });
  it("never interprets a format-only clear as deleting contents", async () => {
    const { adapter, controller } = setup();
    expect(
      await controller.handle(request("excel.range.clear", { range, applyTo: "formats" })),
    ).toMatchObject({ type: "error", error: { code: "UNSUPPORTED_OPERATION" } });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A1" })).values).toEqual([[10]]);
  });
  it("rejects a commit when the target changed after preview", async () => {
    const { adapter, controller } = setup();
    const operationId = await preview(controller);
    await adapter.writeRange({ worksheet: "Sheet1", range: "A1", values: [[99]] });
    expect(
      await controller.handle(request("excel.operation.commit", { operationId })),
    ).toMatchObject({ type: "error", error: { code: "WORKBOOK_CONFLICT" } });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A1" })).values).toEqual([[99]]);
  });
  it("rechecks after approval so edits made while the dialog is open survive", async () => {
    const { adapter, controller } = setup(async () => {
      await adapter.writeRange({ worksheet: "Sheet1", range: "A1", values: [[99]] });
      return true;
    });
    expect(
      await controller.handle(request("excel.range.write", { range, values: [[20]] })),
    ).toMatchObject({ type: "error", error: { code: "WORKBOOK_CONFLICT" } });
  });
  it("refuses undo over a later user edit", async () => {
    const { adapter, controller } = setup();
    const operationId = await preview(controller);
    await controller.handle(request("excel.operation.commit", { operationId }));
    await adapter.writeRange({ worksheet: "Sheet1", range: "A1", values: [[99]] });
    expect(await controller.handle(request("excel.operation.undo", { operationId }))).toMatchObject(
      { type: "error", error: { code: "WORKBOOK_CONFLICT" } },
    );
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A1" })).values).toEqual([[99]]);
  });
  it("rejects concurrent actions instead of overwriting a pending approval", async () => {
    let resolveApproval: (approved: boolean) => void = () => {};
    const approve = vi.fn(
      () =>
        new Promise<boolean>((resolve) => {
          resolveApproval = resolve;
        }),
    );
    const { controller } = setup(approve);
    const first = controller.handle(request("excel.range.write", { range, values: [[20]] }));
    await vi.waitFor(() => expect(approve).toHaveBeenCalledOnce());
    const second = await controller.handle(
      request("excel.range.clear", { range, applyTo: "contents" }),
    );
    resolveApproval(false);
    await first;
    expect(second).toMatchObject({ type: "error", error: { code: "HOST_BUSY" } });
    expect(approve).toHaveBeenCalledOnce();
  });
  it("delivers exact before/after cells to the approval UI", async () => {
    const approve = vi.fn(() => true);
    const { controller } = setup(approve);
    await controller.handle(request("excel.range.write", { range, values: [[20]] }));
    expect(approve).toHaveBeenCalledWith(
      expect.objectContaining({ affectedCells: 1, before: [[10]], after: [[20]] }),
    );
  });
  it("bounds table reads before loading their contents", async () => {
    const { adapter, controller } = setup();
    vi.spyOn(adapter, "listTables").mockResolvedValue([
      { name: "BigTable", worksheet: "Sheet1", range: "A1:A20000" },
    ]);
    const read = vi.spyOn(adapter, "readTable");
    expect(
      await controller.handle(
        request("excel.table.read", { tableId: "BigTable", offset: 0, limit: 10 }),
      ),
    ).toMatchObject({ type: "error", error: { code: "RANGE_LIMIT_EXCEEDED" } });
    expect(read).not.toHaveBeenCalled();
  });
});
