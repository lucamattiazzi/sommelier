import { createAddinController } from "@lucamattiazzi/sommelier-addin-core";
import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import { createInMemoryTransportPair } from "@lucamattiazzi/sommelier-transport";
import { expect, it, vi } from "vitest";
import { createPairAddinSession, createPairClient } from "./index.js";

it("rejects a pending request immediately when its transport disconnects", async () => {
  const [agent, peer] = createInMemoryTransportPair();
  await peer.connect();
  const client = createPairClient({ transport: agent, timeoutMs: 100 });
  await client.connect();
  const pending = client.request("excel.context.get", {});
  const result = expect(pending).rejects.toThrow(/disconnect/i);
  await agent.close();
  await result;
  await client.close();
});

it("cleans up a failed start so the same binding can retry", async () => {
  const [addin, agent] = createInMemoryTransportPair();
  const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1" }] });
  vi.spyOn(adapter, "getContext").mockRejectedValueOnce(new Error("Excel is not ready"));
  const controller = createAddinController({ adapter, workbook: { id: "test", name: "Test" } });
  const session = createPairAddinSession({ transport: addin, controller });
  await expect(session.start()).rejects.toThrow("Excel is not ready");
  expect(addin.state).toBe("disconnected");
  await session.start();
  const client = createPairClient({ transport: agent });
  await client.connect();
  expect(await client.request("excel.context.get", {})).toMatchObject({ mode: "follow-selection" });
  await client.close();
  await session.stop();
});

it("rejects a response with the wrong method even if its request ID matches", async () => {
  const [agent, peer] = createInMemoryTransportPair();
  await peer.connect();
  const client = createPairClient({ transport: agent, createId: () => "one" });
  await client.connect();
  const pending = client.request("excel.context.get", {});
  const result = expect(pending).rejects.toThrow(/method/i);
  peer.send({
    protocolVersion: "0.2",
    type: "response",
    id: "one",
    method: "excel.sheet.list",
    result: { sheets: [] },
  });
  await result;
  await client.close();
});
