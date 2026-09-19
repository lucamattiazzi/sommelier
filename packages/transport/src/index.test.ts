import { PROTOCOL_VERSION, type ProtocolMessage } from "@lucamattiazzi/sommelier-protocol";
import { describe, expect, it } from "vitest";
import { createInMemoryTransportPair } from "./index.js";

const message: ProtocolMessage = {
  protocolVersion: PROTOCOL_VERSION,
  type: "request",
  id: "request-1",
  method: "excel.sheet.list",
  params: {},
};

describe("in-memory transport", () => {
  it("delivers validated messages in order only while connected", async () => {
    const [left, right] = createInMemoryTransportPair();
    const received: ProtocolMessage[] = [];
    right.subscribe((incoming) => received.push(incoming));

    await left.connect();
    await right.connect();
    await left.send(message);
    await left.send({ ...message, id: "request-2" });

    expect(received).toEqual([message, { ...message, id: "request-2" }]);
    await right.close();
    expect(() => left.send(message)).toThrow("connected");
  });

  it("isolates subscribers and reports lifecycle state", async () => {
    const [left, right] = createInMemoryTransportPair();
    const states: string[] = [];
    left.subscribeState((state) => states.push(state));
    await left.connect();
    await right.connect();
    await left.close();
    expect(states).toEqual(["connecting", "connected", "closing", "disconnected"]);
  });
});
