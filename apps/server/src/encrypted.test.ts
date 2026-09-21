import { once } from "node:events";
import {
  createEncryptedSocket,
  createPairIdentity,
  pairConnectionUrl,
  relaySocketUrl,
} from "@lucamattiazzi/sommelier-transport";
import { afterEach, expect, it } from "vitest";
import WebSocket from "ws";
import { createPairServer, type PairServer } from "./index.js";

let server: PairServer | undefined;
afterEach(async () => {
  await server?.close();
  server = undefined;
});

it("reconnects a stable UID and relays only opaque encrypted frames", async () => {
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  const identity = createPairIdentity();
  const origin = `http://${server.host}:${server.port}`;
  const addinUrl = await pairConnectionUrl(origin, identity, "addin");
  const agentUrl = await pairConnectionUrl(origin, identity, "agent");
  const open = (url: string) => {
    const native = new globalThis.WebSocket(relaySocketUrl(url));
    const secure = createEncryptedSocket(native, url);
    const authenticated = new Promise<void>((resolve) =>
      secure.addEventListener("message", (event) => {
        if ("data" in event && JSON.parse(String(event.data)).data?.connected) resolve();
      }),
    );
    return { native, secure, authenticated };
  };
  const addin = open(addinUrl);
  const agent = open(agentUrl);
  await Promise.all([addin.authenticated, agent.authenticated]);
  const response = new Promise<string>((resolve) =>
    agent.secure.addEventListener("message", (e) => {
      if ("data" in e && String(e.data).includes("synthetic cells")) resolve(String(e.data));
    }),
  );
  addin.secure.send(JSON.stringify({ value: "synthetic cells" }));
  expect(await response).toContain("synthetic cells");
  const closed = new Promise<void>((resolve) =>
    addin.native.addEventListener("close", () => resolve()),
  );
  addin.secure.close();
  await closed;
  const next = open(addinUrl);
  await next.authenticated;
  expect(agent.native.readyState).toBe(1);
  next.secure.close();
  agent.secure.close();
});

it("rejects plaintext on encrypted sessions and prevents duplicate role takeover", async () => {
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  const url = relaySocketUrl(
    await pairConnectionUrl(`http://${server.host}:${server.port}`, createPairIdentity(), "agent"),
  );
  const agent = new WebSocket(url);
  await once(agent, "open");
  const duplicate = new WebSocket(url);
  expect((await once(duplicate, "close"))[0]).toBe(1008);
  const closed = once(agent, "close");
  agent.send(
    JSON.stringify({
      protocolVersion: "0.2",
      type: "request",
      method: "excel.range.read",
      id: "1",
      params: {},
    }),
  );
  expect((await closed)[0]).toBe(1007);
});

it("does not offer a plaintext pairing downgrade by default", async () => {
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  expect(
    (await fetch(`http://${server.host}:${server.port}/api/pair/sessions`, { method: "POST" }))
      .status,
  ).toBe(410);
});

it("remembers a named bridge and reconnects without another URL", async () => {
  const { execFile } = await import("node:child_process");
  const { promisify } = await import("node:util");
  const { mkdtemp, rm, stat } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const { join } = await import("node:path");
  const run = promisify(execFile);
  const directory = await mkdtemp(join(tmpdir(), "pair-profile-test-"));
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  const identity = createPairIdentity();
  const origin = `http://${server.host}:${server.port}`;
  const url = await pairConnectionUrl(origin, identity, "agent");
  const env: NodeJS.ProcessEnv = { ...process.env, SOMMELIER_HOME: directory };
  delete env.SOMMELIER_URL;
  delete env.SOMMELIER_SESSION;
  const script = "skills/sommelier/scripts/session.mjs";
  const invoke = (args: string[], pairUrl?: string) =>
    run(process.execPath, [script, ...args], {
      env: { ...env, ...(pairUrl ? { SOMMELIER_URL: pairUrl } : {}) },
    }).then(({ stdout }) => JSON.parse(stdout));
  try {
    const first = await invoke(["start", "--name", "desk"], url);
    expect(first.ok).toBe(true);
    expect(await invoke(["list"])).toMatchObject({ terminals: [{ name: "desk" }] });
    expect((await stat(join(directory, "desk.json"))).mode & 0o777).toBe(0o600);
    const relayPort = server.port;
    await server.close();
    server = undefined;
    await new Promise((resolve) => setTimeout(resolve, 800));
    server = await createPairServer({ port: relayPort, publicOrigin: "https://pair.example.test" });
    await expect
      .poll(async () => (await invoke(["status", "--name", "desk"])).state, { timeout: 5_000 })
      .toBe("open");
    const occupied = new WebSocket(relaySocketUrl(url));
    expect((await once(occupied, "close"))[0]).toBe(1008);

    await invoke(["stop", "--name", "desk"]);
    const second = await invoke(["start", "--name", "desk"]);
    expect(second.ok).toBe(true);
    expect(second.session).not.toBe(first.session);
    expect(await invoke(["status", "--name", "desk"])).toMatchObject({
      ok: true,
      taskPaneConnected: false,
    });
    await invoke(["stop", "--name", "desk"]);
    await invoke(["forget", "--name", "desk"]);
    expect(await invoke(["list"])).toMatchObject({ terminals: [] });
  } finally {
    await invoke(["stop", "--name", "desk"]).catch(() => undefined);
    await rm(directory, { recursive: true, force: true });
  }
}, 15_000);

it("serves sideload manifests as XML without immutable caching", async () => {
  const { mkdtemp, writeFile, rm } = await import("node:fs/promises");
  const { join } = await import("node:path");
  const { tmpdir } = await import("node:os");
  const directory = await mkdtemp(join(tmpdir(), "pair-assets-"));
  try {
    await writeFile(join(directory, "manifest.xml"), "<OfficeApp />");
    server = await createPairServer({
      port: 0,
      publicOrigin: "https://pair.example.test",
      staticDirectory: directory,
    });
    const response = await fetch(`http://${server.host}:${server.port}/manifest.xml`);
    expect(response.headers.get("content-type")).toContain("application/xml");
    expect(response.headers.get("cache-control")).toBe("no-cache");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

it("allows a separately trusted pane origin and rejects other browser origins", async () => {
  server = await createPairServer({
    port: 0,
    publicOrigin: "https://relay.example.test",
    addinOrigin: "https://trusted-pane.example.test",
  });
  const origin = `http://${server.host}:${server.port}`;
  const config = await fetch(`${origin}/api/config`, {
    headers: { origin: "https://trusted-pane.example.test" },
  });
  expect(config.headers.get("access-control-allow-origin")).toBe(
    "https://trusted-pane.example.test",
  );
  const denied = await fetch(`${origin}/api/config`, {
    headers: { origin: "https://untrusted.example.test" },
  });
  expect(denied.headers.get("access-control-allow-origin")).toBeNull();
  const url = relaySocketUrl(await pairConnectionUrl(origin, createPairIdentity(), "addin"));
  const bad = new WebSocket(url, { origin: "https://untrusted.example.test" });
  expect((await once(bad, "close"))[0]).toBe(1008);
  const good = new WebSocket(url, { origin: "https://trusted-pane.example.test" });
  await once(good, "open");
  good.close();
});

it.each([
  ["uid", "invalid", "Invalid terminal UID."],
  ["role", "invalid", "Invalid terminal role."],
  ["token", "", "Missing or invalid relay token. Copy the complete pairing URL."],
])("identifies invalid %s without exposing credentials", async (parameter, value, reason) => {
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  const url = new URL(
    relaySocketUrl(
      await pairConnectionUrl(
        `http://${server.host}:${server.port}`,
        createPairIdentity(),
        "agent",
      ),
    ),
  );
  url.searchParams.set(parameter, value);
  const socket = new WebSocket(url);
  const [code, message] = await once(socket, "close");
  expect(code).toBe(1008);
  expect(String(message)).toBe(reason);
});

it("distinguishes a rejected browser origin from invalid credentials", async () => {
  server = await createPairServer({ port: 0, publicOrigin: "https://pair.example.test" });
  const url = relaySocketUrl(
    await pairConnectionUrl(`http://${server.host}:${server.port}`, createPairIdentity(), "agent"),
  );
  const socket = new WebSocket(url, { origin: "https://untrusted.example.test" });
  const [code, message] = await once(socket, "close");
  expect(code).toBe(1008);
  expect(String(message)).toBe("Origin is not allowed.");
});

it.each(["uid", "token"])(
  "rejects a missing %s locally before saving a bridge profile",
  async (parameter) => {
    const { execFile } = await import("node:child_process");
    const { promisify } = await import("node:util");
    const { mkdtemp, readdir, rm } = await import("node:fs/promises");
    const { tmpdir } = await import("node:os");
    const { join } = await import("node:path");
    const directory = await mkdtemp(join(tmpdir(), "pair-invalid-url-"));
    const url = new URL(
      await pairConnectionUrl("https://pair.example.test", createPairIdentity(), "agent"),
    );
    url.searchParams.delete(parameter);
    try {
      await expect(
        promisify(execFile)(
          process.execPath,
          ["skills/sommelier/scripts/session.mjs", "start", "--name", "desk"],
          {
            env: { ...process.env, SOMMELIER_HOME: directory, SOMMELIER_URL: url.toString() },
          },
        ),
      ).rejects.toMatchObject({
        stderr: expect.stringContaining(
          parameter === "uid" ? "Invalid terminal UID" : "Missing or invalid relay token",
        ),
      });
      expect(await readdir(directory)).toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
);
