import { execFile } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { createAddinController } from "@lucamattiazzi/sommelier-addin-core";
import { createPairAddinSession } from "@lucamattiazzi/sommelier-client";
import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
import {
  createEncryptedSocket,
  createJsonSocketTransport,
  createPairIdentity,
  pairConnectionUrl,
  relaySocketUrl,
} from "@lucamattiazzi/sommelier-transport";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { expect, it } from "vitest";
import { z } from "zod";
import { createPairServer } from "../../../apps/server/src/index.js";
import { type BridgeTrace, bridgeCommand, subscribeBridge } from "./bridge.js";
import { createExcelMcpServer } from "./mcp.js";

it("runs MCP reads, approved writes/charts and independent private traces over the encrypted bridge", async () => {
  const dir = await mkdtemp(join(tmpdir(), "pair-harness-test-"));
  const env: NodeJS.ProcessEnv = { ...process.env, SOMMELIER_HOME: dir };
  delete env.SOMMELIER_URL;
  const invoke = (args: string[], url?: string) =>
    promisify(execFile)(
      process.execPath,
      ["skills/sommelier/scripts/session.mjs", ...args, "--name", "desk"],
      { env: { ...env, ...(url ? { SOMMELIER_URL: url } : {}) } },
    ).then((r) => JSON.parse(r.stdout));
  const relay = await createPairServer({ port: 0, publicOrigin: "https://example.test" });
  const origin = `http://${relay.host}:${relay.port}`;
  const identity = createPairIdentity();
  const agentUrl = await pairConnectionUrl(origin, identity, "agent");
  const addinUrl = await pairConnectionUrl(origin, identity, "addin");
  const native = new WebSocket(relaySocketUrl(addinUrl));
  const encrypted = createEncryptedSocket(native, addinUrl);
  const transport = createJsonSocketTransport(encrypted);
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
  const session = createPairAddinSession({
    transport,
    controller: createAddinController({
      adapter,
      workbook: { id: "synthetic", name: "Synthetic" },
      requestApproval: async () => true,
    }),
  });
  const traces: BridgeTrace[] = [],
    contents: BridgeTrace[] = [];
  let off = () => {},
    offContent = () => {};
  let mcp: ReturnType<typeof createExcelMcpServer> | undefined;
  const channelClient = new Client({ name: "channel-test", version: "1" });
  const client = new Client({ name: "pair-test", version: "1" });
  try {
    await session.start();
    const started = await invoke(["start"], agentUrl);
    const call = (payload: Parameters<typeof bridgeCommand>[1]) =>
      bridgeCommand(started.session, payload);
    await expect.poll(async () => (await call({ command: "status" })).taskPaneConnected).toBe(true);
    off = subscribeBridge(started.session, (e) => traces.push(e));
    offContent = subscribeBridge(started.session, (e) => contents.push(e), {
      includeContent: true,
    });
    mcp = createExcelMcpServer(call, true);
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([mcp.connect(serverTransport), client.connect(clientTransport)]);
    expect(
      (await client.listTools()).tools.some((tool) => tool.name === "excel_chart_create"),
    ).toBe(true);
    const request = async (name: string, args: Record<string, unknown>) => {
      const result = await client.callTool({ name, arguments: args });
      expect(result.isError).not.toBe(true);
      const content = result.content as { type: string; text: string }[];
      return JSON.parse(content[0]?.text ?? "null");
    };
    expect(
      await request("excel_range_read", { range: { sheetId: "Sales", address: "A1:B2" } }),
    ).toMatchObject({
      values: [
        ["Month", "Sales"],
        ["Jan", 12],
      ],
    });
    for (const [method, params] of [
      ["excel.range.write", { range: { sheetId: "Sales", address: "B2" }, values: [[42]] }],
      [
        "excel.chart.create",
        {
          range: { sheetId: "Sales", address: "A1:B2" },
          name: "Monthly",
          title: "Monthly sales",
          chartType: "column",
        },
      ],
    ] as const) {
      const preview = await request("excel_operation_preview", { method, params });
      await request("excel_operation_commit", { operationId: preview.operationId });
    }
    expect((await adapter.readRange({ worksheet: "Sales", range: "B2" })).values).toEqual([[42]]);
    expect(await adapter.listCharts("Sales")).toHaveLength(1);
    // A disconnected chat consumer must not steal the next user's message on restart.
    const abort = new AbortController();
    const waiting = bridgeCommand(
      started.session,
      { command: "next", timeoutMs: 30000 },
      31000,
      abort.signal,
    );
    const rejected = expect(waiting).rejects.toThrow();
    await expect.poll(async () => (await call({ command: "status" })).waitingConsumers).toBe(1);
    abort.abort();
    await rejected;
    await expect.poll(async () => (await call({ command: "status" })).waitingConsumers).toBe(0);
    const notification = new Promise<string>((resolve) => {
      channelClient.setNotificationHandler(
        z.object({
          method: z.literal("notifications/claude/channel"),
          params: z.object({ content: z.string(), meta: z.object({ profile: z.string() }) }),
        }),
        (event) => {
          expect(event.params.meta.profile).toBe("desk");
          resolve(event.params.content);
        },
      );
    });
    await channelClient.connect(
      new StdioClientTransport({
        command: process.execPath,
        args: [resolve("packages/bridge/dist/agent.js"), "mcp", "--name", "desk", "--channel"],
        env: {
          ...Object.fromEntries(
            Object.entries(env).filter(
              (entry): entry is [string, string] => entry[1] !== undefined,
            ),
          ),
        },
      }),
    );
    session.sendUserMessage("Synthetic request");
    expect(await notification).toBe("Synthetic request");
    await channelClient.callTool({ name: "pair_reply", arguments: { content: "Channel answer" } });
    await expect
      .poll(() => contents.some((e) => JSON.stringify(e).includes("Synthetic request")))
      .toBe(true);
    expect(
      traces.some(
        (e) =>
          e.kind === "rpc.response" && e.outcome === "success" && typeof e.durationMs === "number",
      ),
    ).toBe(true);
    expect(traces.every((e) => e.payload === undefined)).toBe(true);
    expect(JSON.stringify(contents)).not.toContain(identity.secret);
  } finally {
    off();
    offContent();
    await channelClient.close();
    await client.close();
    await mcp?.close();
    await invoke(["stop"]).catch(() => {});
    await session.stop();
    await relay.close();
    await rm(dir, { recursive: true, force: true });
  }
}, 15000);
