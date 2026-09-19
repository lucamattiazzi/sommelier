import { once } from "node:events";
import { PROTOCOL_VERSION } from "@lucamattiazzi/sommelier-protocol";
import { describe, expect, it } from "vitest";
import WebSocket from "ws";
import { createPairRelayServer } from "./index.js";

describe("Pair WebSocket relay", () => {
  it("routes validated envelopes between paired roles", async () => {
    const relay = await createPairRelayServer({ port: 0 });
    const base = `ws://${relay.host}:${relay.port}/pair?session=test-session`;
    const addin = new WebSocket(`${base}&role=addin`);
    const agent = new WebSocket(`${base}&role=agent`);
    await Promise.all([once(addin, "open"), once(agent, "open")]);
    const received = once(agent, "message");
    const request = {
      protocolVersion: PROTOCOL_VERSION,
      type: "request",
      id: "request-1",
      method: "excel.sheet.list",
      params: {},
    };
    addin.send(JSON.stringify(request));
    const [data] = await received;
    expect(JSON.parse(String(data))).toEqual(request);
    addin.close();
    agent.close();
    await relay.close();
  });
});
