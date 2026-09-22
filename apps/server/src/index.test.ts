import { execFile } from "node:child_process";
import { once } from "node:events";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { PROTOCOL_VERSION } from "@lucamattiazzi/sommelier-protocol";
import { afterEach, describe, expect, it } from "vitest";
import WebSocket, { type RawData } from "ws";
import { createPairServer, type PairServer } from "./index.js";

let server: PairServer | undefined;
const execFileAsync = promisify(execFile);

afterEach(async () => {
  await server?.close();
  server = undefined;
});

describe("Sommelier server", () => {
  it("exposes health and creates authenticated, expiring pairing sessions", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const health = await fetch(`http://${server.host}:${server.port}/healthz`);
    expect(await health.json()).toEqual({ status: "ok" });

    const response = await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, {
      method: "POST",
    });
    expect(response.status).toBe(201);
    expect(await response.json()).toMatchObject({
      sessionId: expect.any(String),
      expiresAt: expect.any(String),
      addinUrl: expect.stringMatching(/^wss:\/\/pair\.example\.test\/pair\?/),
      agentUrl: expect.stringMatching(/^wss:\/\/pair\.example\.test\/pair\?/),
    });
  });

  it("serves the agent bridge script the pairing response points at", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
      bridgeScript: fileURLToPath(
        new URL("../../../skills/sommelier/scripts/session.mjs", import.meta.url),
      ),
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { bridgeUrl: string };
    expect(pairing.bridgeUrl).toBe("https://pair.example.test/agent/session.mjs");

    const script = await fetch(`http://${server.host}:${server.port}/agent/session.mjs`);
    expect(script.headers.get("content-type")).toContain("text/javascript");
    expect(await script.text()).toContain("session.mjs start");
  });

  it("can advertise a separate direct agent origin for local development", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://localhost:3000",
      agentOrigin: "ws://127.0.0.1:3001",
    });
    const response = await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, {
      method: "POST",
    });
    expect(await response.json()).toMatchObject({
      addinUrl: expect.stringMatching(/^wss:\/\/localhost:3000\/pair\?/),
      agentUrl: expect.stringMatching(/^ws:\/\/127\.0\.0\.1:3001\/pair\?/),
    });
  });

  it("relays only validated protocol messages between authenticated peers", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { addinUrl: string; agentUrl: string };
    const runningServer = server;
    const local = (url: string) =>
      url.replace("wss://pair.example.test", `ws://${runningServer.host}:${runningServer.port}`);
    const addin = new WebSocket(local(pairing.addinUrl));
    const agent = new WebSocket(local(pairing.agentUrl));
    await Promise.all([once(addin, "open"), once(agent, "open")]);
    const incoming = once(agent, "message");
    const message = {
      protocolVersion: PROTOCOL_VERSION,
      type: "request",
      id: "request-1",
      method: "excel.sheet.list",
      params: {},
    };
    addin.send(JSON.stringify(message));
    const [data] = await incoming;
    expect(JSON.parse(String(data))).toEqual(message);
    addin.close();
    agent.close();
  });

  it("notifies the add-in when its agent connects and disconnects", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { addinUrl: string; agentUrl: string };
    const origin = `ws://${server.host}:${server.port}`;
    const local = (url: string) => url.replace("wss://pair.example.test", origin);
    const addin = new WebSocket(local(pairing.addinUrl));
    await once(addin, "open");
    const connected = once(addin, "message");
    const agent = new WebSocket(local(pairing.agentUrl));
    await once(agent, "open");
    expect(JSON.parse(String((await connected)[0]))).toMatchObject({
      event: "pair.peer.changed",
      data: { role: "agent", connected: true },
    });
    const disconnected = once(addin, "message");
    agent.close();
    expect(JSON.parse(String((await disconnected)[0]))).toMatchObject({
      event: "pair.peer.changed",
      data: { role: "agent", connected: false },
    });
    addin.close();
  });

  it("rejects unknown pairing credentials", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const socket = new WebSocket(
      `ws://${server.host}:${server.port}/pair?session=missing&role=agent&token=invalid`,
    );
    const [code] = await once(socket, "close");
    expect(code).toBe(1008);
  });

  it("bounds active relay sessions", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
      maxActiveSessions: 1,
    });
    const endpoint = `http://${server.host}:${server.port}/api/pair/sessions`;
    expect((await fetch(endpoint, { method: "POST" })).status).toBe(201);
    expect((await fetch(endpoint, { method: "POST" })).status).toBe(503);
  });

  it("serves the add-in with restrictive browser headers", async () => {
    const directory = await mkdtemp(join(tmpdir(), "sommelier-"));
    await writeFile(join(directory, "index.html"), "<h1>Pair</h1>");
    try {
      server = await createPairServer({
        allowUnencryptedLegacy: true,
        port: 0,
        publicOrigin: "https://pair.example.test",
        staticDirectory: directory,
      });
      const response = await fetch(`http://${server.host}:${server.port}/`);
      expect(await response.text()).toBe("<h1>Pair</h1>");
      expect(response.headers.get("content-security-policy")).toContain("connect-src");
      expect(response.headers.get("cache-control")).toBe("no-cache");
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("keeps one agent connection for task-pane chat and Excel work", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { addinUrl: string; agentUrl: string };
    const origin = `ws://${server.host}:${server.port}`;
    const local = (url: string) => url.replace("wss://pair.example.test", origin);
    const addin = new WebSocket(local(pairing.addinUrl));
    await once(addin, "open");

    let agentConnections = 0;
    const agentMessages: string[] = [];
    addin.on("message", (data: RawData) => {
      const message = JSON.parse(String(data));
      if (message.type === "request") {
        addin.send(
          JSON.stringify({
            protocolVersion: PROTOCOL_VERSION,
            type: "response",
            id: message.id,
            method: message.method,
            result: { sheets: [{ id: "Sheet1", name: "Sheet1", visibility: "visible" }] },
          }),
        );
        return;
      }
      if (message.event === "pair.peer.changed" && message.data.connected) agentConnections += 1;
      if (message.event === "pair.chat.message" && message.data.sender === "agent") {
        agentMessages.push(message.data.content);
      }
    });

    const script = fileURLToPath(
      new URL("../../../skills/sommelier/scripts/session.mjs", import.meta.url),
    );
    const run = (args: readonly string[], session?: string) =>
      execFileAsync(process.execPath, [script, ...args], {
        env: {
          ...process.env,
          SOMMELIER_URL: local(pairing.agentUrl),
          ...(session ? { SOMMELIER_SESSION: session } : {}),
        },
      }).then(({ stdout }) => JSON.parse(stdout));

    const started = await run(["start"]);
    expect(started).toMatchObject({ ok: true });
    const session = started.session as string;

    try {
      const sendUserMessage = (content: string) =>
        addin.send(
          JSON.stringify({
            protocolVersion: PROTOCOL_VERSION,
            type: "event",
            event: "pair.chat.message",
            data: {
              messageId: `message-${content.length}`,
              sender: "user",
              content,
              occurredAt: new Date().toISOString(),
            },
          }),
        );

      sendUserMessage("Which sheets exist?");
      expect(await run(["next", "--timeout", "10000"], session)).toMatchObject({
        type: "message",
        message: { sender: "user", content: "Which sheets exist?" },
      });

      expect(
        await run(["request", "--method", "excel.sheet.list", "--params", "{}"], session),
      ).toMatchObject({ ok: true, result: { sheets: [{ id: "Sheet1" }] } });

      expect(await run(["reply", "--content", "Only Sheet1 is visible."], session)).toMatchObject({
        ok: true,
      });

      const waiting = run(["next", "--timeout", "10000"], session);
      setTimeout(() => sendUserMessage("Read A1 next."), 300);
      expect(await waiting).toMatchObject({
        type: "message",
        message: { content: "Read A1 next." },
      });

      expect(agentMessages).toEqual(["Only Sheet1 is visible."]);
      expect(agentConnections).toBe(1);
    } finally {
      await run(["stop"], session).catch(() => undefined);
      addin.close();
    }
  }, 30_000);

  it("fails a second bridge start with the reason the relay gave", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { addinUrl: string; agentUrl: string };
    const origin = `ws://${server.host}:${server.port}`;
    const local = (url: string) => url.replace("wss://pair.example.test", origin);
    const addin = new WebSocket(local(pairing.addinUrl));
    await once(addin, "open");

    const script = fileURLToPath(
      new URL("../../../skills/sommelier/scripts/session.mjs", import.meta.url),
    );
    const start = () =>
      execFileAsync(process.execPath, [script, "start"], {
        env: { ...process.env, SOMMELIER_URL: local(pairing.agentUrl) },
      });

    const first = JSON.parse((await start()).stdout) as { session: string };
    try {
      await expect(start()).rejects.toThrow(/already connected/i);
    } finally {
      await execFileAsync(process.execPath, [script, "stop"], {
        env: { ...process.env, SOMMELIER_SESSION: first.session },
      }).catch(() => undefined);
      addin.close();
    }
  }, 30_000);

  it("supports the standalone Agent Skill RPC client", async () => {
    server = await createPairServer({
      allowUnencryptedLegacy: true,
      port: 0,
      publicOrigin: "https://pair.example.test",
    });
    const pairing = (await (
      await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" })
    ).json()) as { addinUrl: string; agentUrl: string };
    const origin = `ws://${server.host}:${server.port}`;
    const localUrl = (url: string) => url.replace("wss://pair.example.test", origin);
    const addin = new WebSocket(localUrl(pairing.addinUrl));
    await once(addin, "open");
    const respond = (data: RawData): void => {
      const request = JSON.parse(String(data)) as { type: string; id: string; method: string };
      if (request.type !== "request") return;
      addin.off("message", respond);
      addin.send(
        JSON.stringify({
          protocolVersion: PROTOCOL_VERSION,
          type: "response",
          id: request.id,
          method: request.method,
          result: { sheets: [] },
        }),
      );
    };
    addin.on("message", respond);

    const { stdout } = await execFileAsync(
      process.execPath,
      [
        fileURLToPath(new URL("../../../skills/sommelier/scripts/rpc.mjs", import.meta.url)),
        "--method",
        "excel.sheet.list",
        "--params",
        "{}",
      ],
      { env: { ...process.env, SOMMELIER_URL: localUrl(pairing.agentUrl) } },
    );
    expect(JSON.parse(stdout)).toMatchObject({ type: "response", result: { sheets: [] } });
    addin.close();
  });
});

it("serves the standalone offline documentation module alongside the bridge", async () => {
  server = await createPairServer({
    port: 0,
    publicOrigin: "https://pair.example.test",
    bridgeScript: fileURLToPath(
      new URL("../../../skills/sommelier/scripts/session.mjs", import.meta.url),
    ),
  });
  const response = await fetch(`http://${server.host}:${server.port}/agent/lib/excel-docs.mjs`);
  expect(response.status).toBe(200);
  expect(response.headers.get("content-type")).toContain("text/javascript");
  expect(response.headers.get("cache-control")).toBe("no-cache");
  expect(await response.text()).toContain("searchExcelDocs");
});
