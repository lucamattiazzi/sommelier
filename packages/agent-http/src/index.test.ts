import { once } from "node:events";
import { createServer } from "node:http";
import { PROTOCOL_VERSION } from "@lucamattiazzi/sommelier-core";
import { afterEach, describe, expect, it } from "vitest";
import { HttpAgentAdapter } from "./index.js";

const servers: ReturnType<typeof createServer>[] = [];
afterEach(() =>
  Promise.all(
    servers.map((server) => new Promise<void>((resolve) => server.close(() => resolve()))),
  ),
);

async function endpoint(body: unknown, status = 200): Promise<string> {
  const server = createServer((_, response) => {
    response.writeHead(status, { "content-type": "application/json" });
    response.end(JSON.stringify(body));
  });
  servers.push(server);
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Missing address");
  return `http://127.0.0.1:${address.port}`;
}

const request = {
  protocolVersion: PROTOCOL_VERSION,
  sessionId: "s",
  turnId: "t",
  iteration: 1,
  userMessage: {
    protocolVersion: PROTOCOL_VERSION,
    type: "user_message" as const,
    text: "hi",
    selectionIncluded: false,
  },
  tools: [],
  toolResults: [],
};

describe("HttpAgentAdapter", () => {
  it("validates and yields JSON events", async () => {
    const adapter = new HttpAgentAdapter({
      endpoint: await endpoint({
        protocolVersion: PROTOCOL_VERSION,
        events: [{ protocolVersion: PROTOCOL_VERSION, type: "completion" }],
      }),
    });
    const events = [];
    for await (const event of adapter.runTurn(request, { signal: new AbortController().signal }))
      events.push(event);
    expect(events).toHaveLength(1);
  });

  it("returns safe status errors", async () => {
    const adapter = new HttpAgentAdapter({
      endpoint: await endpoint({ token: "must-not-leak" }, 401),
      getAccessToken: async () => "private-token",
    });
    const run = async () => {
      for await (const _event of adapter.runTurn(request, {
        signal: new AbortController().signal,
      })) {
        /* consume */
      }
    };
    await expect(run()).rejects.toMatchObject({ code: "SOMMELIER_HTTP_STATUS_ERROR" });
    await expect(run()).rejects.not.toThrow(/private-token|must-not-leak/);
  });
});
