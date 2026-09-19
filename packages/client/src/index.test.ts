import { createAddinController } from "@lucamattiazzi/sommelier-addin-core";
import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import { createInMemoryTransportPair } from "@lucamattiazzi/sommelier-transport";
import { describe, expect, it } from "vitest";
import { createPairAddinSession, createPairClient, createPairRelay } from "./index.js";

describe("Pair vertical slice", () => {
  it("routes selection, read, and approved write through a semantic-free relay", async () => {
    const [addinTransport, relayAddin] = createInMemoryTransportPair();
    const [relayAgent, agentTransport] = createInMemoryTransportPair();
    const adapter = new InMemoryExcelAdapter({
      sheets: [{ name: "Sheet1", values: [["Amount"], [10]] }],
    });
    const controller = createAddinController({
      adapter,
      workbook: { id: "book-1", name: "Synthetic.xlsx" },
      requestApproval: () => true,
    });
    const relay = createPairRelay(relayAddin, relayAgent);
    const addin = createPairAddinSession({ transport: addinTransport, controller });
    const client = createPairClient({ transport: agentTransport, createId: () => "request-1" });
    const events: string[] = [];
    client.subscribe((event) => events.push(event.event));

    await relay.start();
    await addin.start();
    await client.connect();
    expect(
      await client.request("excel.range.read", {
        range: { sheetId: "Sheet1", address: "A2" },
      }),
    ).toMatchObject({ values: [[10]] });
    await client.request("excel.range.select", {
      range: { sheetId: "Sheet1", address: "A2" },
    });
    expect(events).toContain("excel.selection.changed");
    await client.request("excel.range.write", {
      range: { sheetId: "Sheet1", address: "A2" },
      values: [[25]],
    });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A2" })).values).toEqual([[25]]);

    await client.close();
    await addin.stop();
    await relay.stop();
  });

  it("carries task-pane chat and repeated Excel work over one agent connection", async () => {
    const [addinTransport, agentTransport] = createInMemoryTransportPair();
    const controller = createAddinController({
      adapter: new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1", values: [[10]] }] }),
      workbook: { id: "book-1", name: "Synthetic.xlsx" },
    });
    let messageNumber = 0;
    const addin = createPairAddinSession({
      transport: addinTransport,
      controller,
      createId: () => `message-${++messageNumber}`,
      now: () => new Date("2026-08-31T10:00:00Z"),
    });
    const client = createPairClient({
      transport: agentTransport,
      createId: () => `agent-${++messageNumber}`,
      now: () => new Date("2026-08-31T10:00:01Z"),
    });
    const fromUser: string[] = [];
    const fromAgent: string[] = [];
    client.subscribeChat((message) => fromUser.push(message.data.content));
    addin.subscribeChat((message) => fromAgent.push(message.data.content));

    await addin.start();
    await client.connect();

    addin.sendUserMessage("Which sheets exist?");
    expect(fromUser).toEqual(["Which sheets exist?"]);
    expect(fromAgent).toEqual([]);

    expect(await client.request("excel.sheet.list", {})).toMatchObject({
      sheets: [{ id: "Sheet1" }],
    });
    client.sendAgentMessage("Only Sheet1 is visible.");
    expect(fromAgent).toEqual(["Only Sheet1 is visible."]);

    addin.sendUserMessage("Read A1.");
    expect(
      await client.request("excel.range.read", {
        range: { sheetId: "Sheet1", address: "A1" },
      }),
    ).toMatchObject({ values: [[10]] });
    expect(fromUser).toEqual(["Which sheets exist?", "Read A1."]);

    await client.close();
    await addin.stop();
  });

  it("keeps ask-before-write as the Pair default", async () => {
    const [addinTransport, agentTransport] = createInMemoryTransportPair();
    const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1", values: [[10]] }] });
    const controller = createAddinController({
      adapter,
      workbook: { id: "book-1", name: "Synthetic.xlsx" },
    });
    const addin = createPairAddinSession({ transport: addinTransport, controller });
    const client = createPairClient({ transport: agentTransport });
    await addin.start();
    await client.connect();

    await expect(
      client.request("excel.range.write", {
        range: { sheetId: "Sheet1", address: "A1" },
        values: [[99]],
      }),
    ).rejects.toMatchObject({ code: "APPROVAL_REQUIRED" });
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A1" })).values).toEqual([[10]]);
  });
});
