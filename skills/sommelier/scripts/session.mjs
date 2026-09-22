#!/usr/bin/env node

import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { connect, createServer } from "node:net";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";

const PROTOCOL_VERSION = "0.2";
const MAX_CONTENT_LENGTH = 32_000;
const CLOSED_SESSION_GRACE_MS = 5 * 60_000;
const MAX_QUEUED_MESSAGES = 1_000;

function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index < 0 ? fallback : process.argv[index + 1];
}

function fail(message) {
  process.stderr.write(`${message}\n`);
  process.exit(1);
}

function positiveInteger(value, name) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) fail(`${name} must be a positive integer.`);
  return parsed;
}

function readLine(stream) {
  return new Promise((resolve, reject) => {
    const reader = createInterface({ input: stream });
    reader.once("line", (line) => {
      resolve(line.trim());
      reader.close();
    });
    reader.once("close", () => resolve(""));
    stream.once("error", reject);
  });
}

const profileDirectory =
  process.env.SOMMELIER_HOME ??
  process.env.AI_CDL_PAIR_HOME ??
  (existsSync(join(homedir(), ".ai-cdl-pair")) && !existsSync(join(homedir(), ".sommelier"))
    ? join(homedir(), ".ai-cdl-pair")
    : join(homedir(), ".sommelier"));
function profileName() {
  const name = argument("--name");
  if (name !== undefined && !/^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$/.test(name))
    fail("Use a terminal name of 1–64 letters, numbers, underscores or hyphens.");
  return name;
}
function readProfile(name) {
  if (!name) return undefined;
  try {
    return JSON.parse(readFileSync(join(profileDirectory, `${name}.json`), "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return undefined;
    throw error;
  }
}
function saveProfile(name, value) {
  mkdirSync(profileDirectory, { recursive: true, mode: 0o700 });
  chmodSync(profileDirectory, 0o700);
  const temp = join(profileDirectory, `${name}.${randomUUID()}.tmp`);
  writeFileSync(temp, JSON.stringify(value), { mode: 0o600, flag: "wx" });
  renameSync(temp, join(profileDirectory, `${name}.json`));
}

function pairUrl() {
  const url =
    process.env.SOMMELIER_URL ?? process.env.AI_CDL_PAIR_URL ?? readProfile(profileName())?.url;
  if (!url) fail("SOMMELIER_URL is required.");
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return fail("SOMMELIER_URL must be a valid WebSocket URL.");
  }
  if (parsed.protocol !== "ws:" && parsed.protocol !== "wss:") {
    fail("SOMMELIER_URL must use ws:// or wss://.");
  }
  if (parsed.protocol === "ws:" && !["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname))
    fail("Remote relays require WSS.");
  if (parsed.pathname === "/connect") {
    if (!/^[\w-]{43}$/.test(parsed.hash.slice(1)) || parsed.searchParams.get("role") !== "agent")
      fail("An encrypted agent URL is required.");
    if (!/^[a-f\d-]{36}$/i.test(parsed.searchParams.get("uid") ?? ""))
      fail("Invalid terminal UID. Copy the complete pairing URL.");
    if (!/^[\w-]{43}$/.test(parsed.searchParams.get("token") ?? ""))
      fail(
        "Missing or invalid relay token. Copy the complete pairing URL, including token and #secret.",
      );
  }
  return url;
}

function controlPathForNewSession() {
  if (process.platform === "win32") return `\\\\.\\pipe\\sommelier-${randomUUID()}`;
  const directory = mkdtempSync(join(tmpdir(), "sommelier-"), { mode: 0o700 });
  return join(directory, "control.sock");
}

function removeControlPath(controlPath) {
  if (process.platform === "win32") return;
  rmSync(controlPath, { force: true });
  rmSync(join(controlPath, ".."), { force: true, recursive: true });
}

/** Own the single agent WebSocket and expose it to short-lived local commands. */
async function runDaemon(controlPath) {
  const url = await readLine(process.stdin);
  if (!url) fail("The Sommelier URL was not provided on stdin.");

  const watchers = new Map();
  function trace(kind, message) {
    const metadata = {
      kind,
      occurredAt: new Date().toISOString(),
      ...(typeof message.id === "string" ? { id: message.id } : {}),
      ...(typeof message.method === "string" ? { method: message.method } : {}),
      ...(typeof message.event === "string" ? { event: message.event } : {}),
      ...(message.type === "response"
        ? { outcome: "success" }
        : message.type === "error"
          ? { outcome: "error" }
          : {}),
      ...(typeof message.error?.code === "string" ? { errorCode: message.error.code } : {}),
      ...(["user", "agent"].includes(message.data?.sender) ? { sender: message.data.sender } : {}),
      ...(typeof message.data?.messageId === "string" ? { messageId: message.data.messageId } : {}),
      ...(pending.has(message.id)
        ? { durationMs: Date.now() - pending.get(message.id).startedAt }
        : {}),
    };
    for (const [connection, includeContent] of watchers) {
      if (connection.destroyed) {
        watchers.delete(connection);
        continue;
      }
      connection.write(
        `${JSON.stringify({ ...metadata, ...(includeContent ? { payload: message } : {}) })}\n`,
      );
      if (connection.writableLength > 1_048_576) connection.destroy();
    }
  }
  const inbox = [];
  const waiters = [];
  const pending = new Map();
  let state = "connecting";
  let peerConnected = false;
  let peerGeneration = 0;
  let turnGeneration;
  let handshakeSent = false;
  let ready = false;
  let closeInfo;
  let exitTimer;
  let readyTimer;

  const closedMessage = () =>
    closeInfo?.reason
      ? `Sommelier connection closed (${closeInfo.code}): ${closeInfo.reason}`
      : `Sommelier connection closed${closeInfo ? ` (code ${closeInfo.code})` : ""}.`;

  const handshake = (payload) => {
    if (handshakeSent) return;
    handshakeSent = true;
    process.stdout.write(`${JSON.stringify(payload)}\n`);
  };

  const shutdown = (code) => {
    try {
      control.close();
    } catch {}
    removeControlPath(controlPath);
    process.exit(code);
  };

  const encrypted = new URL(url).pathname === "/connect";
  const encryption = encrypted ? await import("./lib/encrypted-socket.mjs") : undefined;
  let socket;
  let retryTimer;
  let retryDelay = 500;
  let stopping = false;
  const connectRelay = () => {
    const native = new WebSocket(encryption ? encryption.relaySocketUrl(url) : url);
    socket = encryption ? encryption.createEncryptedSocket(native, url) : native;
    attachSocket(socket);
  };
  const attachSocket = (current) => {
    const socket = current;
    let closed = false;

    const confirmReady = () => {
      if (state !== "open") return;
      clearTimeout(readyTimer);
      ready = true;
      handshake({ ok: true });
    };

    socket.addEventListener("open", () => {
      if (closed) return;
      trace("connection", { state: "open" });
      state = "open";
      retryDelay = 500;
      // The relay closes a duplicate role right after the handshake, so give it a moment to object.
      readyTimer = setTimeout(confirmReady, 300);
    });

    const onClose = (event) => {
      if (closed) return;
      closed = true;
      trace("connection", { state: "closed" });
      state = "closed";
      peerConnected = false;
      peerGeneration++;
      inbox.length = 0;
      clearTimeout(readyTimer);
      closeInfo = { code: event.code, reason: String(event.reason ?? "") };
      for (const waiter of waiters.splice(0)) waiter.settle({ ok: true, type: "closed" });
      for (const request of pending.values()) {
        clearTimeout(request.timer);
        request.settle({ ok: false, error: { message: closedMessage() } });
      }
      pending.clear();
      handshake({ ok: false, error: closedMessage() });
      // A session that never became usable has no session path anyone can hold, so it need not linger.
      if (!ready) shutdown(1);
      if (encrypted && !stopping && event.code !== 1008 && event.code !== 1007) {
        state = "connecting";
        retryTimer = setTimeout(connectRelay, retryDelay);
        retryDelay = Math.min(retryDelay * 2, 30_000);
      } else exitTimer = setTimeout(() => shutdown(0), CLOSED_SESSION_GRACE_MS);
    };
    socket.addEventListener("close", onClose);
    socket.addEventListener("error", () => {
      if (closed) return;
      if (state === "connecting" && !ready) {
        handshake({ ok: false, error: "Could not connect to the Sommelier URL." });
        shutdown(1);
      }
      // Node 22 may emit only error when a reconnect is refused, without a close event.
      onClose({ code: 1006, reason: "Connection failed." });
      try {
        socket.close();
      } catch {}
    });

    socket.addEventListener("message", (event) => {
      if (closed) return;
      confirmReady();
      let message;
      try {
        message = JSON.parse(String(event.data));
      } catch {
        return;
      }
      trace(message.type === "event" ? "event" : "rpc.response", message);
      if (message.type === "event") {
        if (message.event === "pair.peer.changed" && message.data?.role === "addin") {
          peerConnected = Boolean(message.data.connected);
          if (!peerConnected) {
            peerGeneration++;
            inbox.length = 0;
            for (const request of pending.values()) {
              clearTimeout(request.timer);
              request.settle({
                ok: false,
                error: { message: "Workbook disconnected. Check its state before retrying." },
              });
            }
            pending.clear();
          }
        }
        if (message.event === "pair.chat.message" && message.data?.sender === "user") {
          const waiter = waiters.shift();
          if (waiter) waiter.settle({ ok: true, type: "message", message: message.data });
          else {
            inbox.push(message.data);
            if (inbox.length > MAX_QUEUED_MESSAGES) inbox.shift();
          }
        }
        return;
      }
      if (message.type !== "response" && message.type !== "error") return;
      const request = pending.get(message.id);
      if (!request) return;
      pending.delete(message.id);
      clearTimeout(request.timer);
      request.settle(
        message.type === "error"
          ? { ok: false, error: message.error }
          : { ok: true, result: message.result },
      );
    });
  };
  connectRelay();

  const commands = {
    begin: () => {
      if (!peerConnected) return { ok: false, error: { message: "Workbook is not connected." } };
      turnGeneration = peerGeneration;
      return { ok: true };
    },
    status: () => ({
      ok: true,
      state,
      taskPaneConnected: peerConnected,
      queuedMessages: inbox.length,
      waitingConsumers: waiters.length,
      pendingRequests: pending.size,
      ...(closeInfo ? { closedBecause: closedMessage() } : {}),
    }),

    next: (input, settle, connection) => {
      if (inbox.length > 0) return { ok: true, type: "message", message: inbox.shift() };
      if (state === "closed") return { ok: true, type: "closed" };
      const timeoutMs = input.timeoutMs ?? 120_000;
      const waiter = { settle: (payload) => settle(payload) };
      waiters.push(waiter);
      const cleanup = () => {
        clearTimeout(timer);
        const index = waiters.indexOf(waiter);
        if (index >= 0) waiters.splice(index, 1);
        connection.off("close", cleanup);
      };
      const timer = setTimeout(() => {
        cleanup();
        settle({ ok: true, type: "idle" });
      }, timeoutMs);
      connection.once("close", cleanup);
      waiter.settle = (payload) => {
        cleanup();
        settle(payload);
      };
      return undefined;
    },

    request: (input, settle) => {
      if (turnGeneration !== undefined && turnGeneration !== peerGeneration)
        return {
          ok: false,
          error: {
            message:
              "The workbook connection changed during this turn. Start a new task after reconnecting.",
          },
        };
      if (state !== "open") return { ok: false, error: { message: closedMessage() } };
      if (encrypted && !peerConnected)
        return {
          ok: false,
          error: { message: "Open Excel and reconnect the authorized terminal first." },
        };
      if (typeof input.method !== "string" || !input.method) {
        return { ok: false, error: { message: "A method is required." } };
      }
      const params = input.params ?? {};
      if (!params || typeof params !== "object" || Array.isArray(params)) {
        return { ok: false, error: { message: "Params must be a JSON object." } };
      }
      const id = randomUUID();
      const timer = setTimeout(() => {
        pending.delete(id);
        settle({ ok: false, error: { message: `${input.method} timed out.` } });
      }, input.timeoutMs ?? 60_000);
      pending.set(id, { settle, timer, startedAt: Date.now() });
      try {
        trace("rpc.request", { id, method: input.method, params });
        socket.send(
          JSON.stringify({
            protocolVersion: PROTOCOL_VERSION,
            type: "request",
            id,
            method: input.method,
            params,
          }),
        );
      } catch (error) {
        pending.delete(id);
        clearTimeout(timer);
        return { ok: false, error: { message: String(error) } };
      }
      return undefined;
    },

    reply: (input) => {
      if (state !== "open") return { ok: false, error: { message: closedMessage() } };
      if (encrypted && !peerConnected)
        return { ok: false, error: { message: "Task pane is not connected." } };
      const content = typeof input.content === "string" ? input.content.trim() : "";
      if (!content) return { ok: false, error: { message: "Reply content is required." } };
      if (content.length > MAX_CONTENT_LENGTH) {
        return { ok: false, error: { message: "Reply content is too long." } };
      }
      const messageId = randomUUID();
      trace("event", { event: "pair.chat.message", data: { sender: "agent", content, messageId } });
      socket.send(
        JSON.stringify({
          protocolVersion: PROTOCOL_VERSION,
          type: "event",
          event: "pair.chat.message",
          data: {
            messageId,
            sender: "agent",
            content,
            occurredAt: new Date().toISOString(),
          },
        }),
      );
      return { ok: true, messageId };
    },
  };

  const control = createServer((connection) => {
    connection.on("error", () => connection.destroy());
    const reader = createInterface({ input: connection });
    reader.on("error", () => {});
    reader.once("line", (line) => {
      let input;
      try {
        input = JSON.parse(line);
      } catch {
        connection.end(
          `${JSON.stringify({ ok: false, error: { message: "Invalid command." } })}\n`,
        );
        return;
      }
      if (input.command === "watch") {
        watchers.set(connection, input.includeContent === true);
        connection.once("close", () => watchers.delete(connection));
        return;
      }
      const respond = (payload) => {
        if (connection.writableEnded) return;
        connection.end(`${JSON.stringify(payload)}\n`);
      };
      if (input.command === "stop") {
        stopping = true;
        clearTimeout(retryTimer);
        respond({ ok: true });
        connection.on("close", () => {
          clearTimeout(exitTimer);
          try {
            socket.close();
          } catch {}
          shutdown(0);
        });
        return;
      }
      const handler = Object.hasOwn(commands, input.command) ? commands[input.command] : undefined;
      if (!handler) {
        respond({ ok: false, error: { message: `Unknown command: ${input.command}.` } });
        return;
      }
      const immediate = handler(input, respond, connection);
      if (immediate) respond(immediate);
    });
  });

  control.on("error", (error) => {
    handshake({ ok: false, error: `Could not open the local control socket: ${error.message}` });
    process.exit(1);
  });
  control.listen(controlPath, () => {
    if (process.platform !== "win32") chmodSync(controlPath, 0o600);
  });

  for (const signal of ["SIGINT", "SIGTERM"]) {
    process.on(signal, () => {
      stopping = true;
      clearTimeout(retryTimer);
      try {
        socket.close();
      } catch {}
      shutdown(0);
    });
  }
}

/** Send one command to a running session daemon. */
function send(controlPath, payload, timeoutMs = 3_600_000) {
  return new Promise((resolve, reject) => {
    const connection = connect(controlPath);
    const timer = setTimeout(() => {
      connection.destroy();
      reject(new Error("The Sommelier session did not respond."));
    }, timeoutMs);
    connection.on("close", () => {
      clearTimeout(timer);
      reject(new Error("The Sommelier bridge closed the local request without a response."));
    });
    connection.on("error", () => {
      clearTimeout(timer);
      reject(new Error("No running Sommelier session. Run `session.mjs start` first."));
    });
    connection.write(`${JSON.stringify(payload)}\n`);
    const reader = createInterface({ input: connection });
    reader.on("error", () => {});
    reader.once("line", (line) => {
      clearTimeout(timer);
      connection.end();
      try {
        resolve(JSON.parse(line));
      } catch {
        reject(new Error("The Sommelier session returned an invalid response."));
      }
    });
  });
}

async function start() {
  const url = pairUrl();
  const name = profileName();
  const existing = readProfile(name);
  if (existing?.session) {
    const status = await send(existing.session, { command: "status" }, 2_000).catch(
      () => undefined,
    );
    if (status?.ok && status.state !== "closed") {
      if (existing.url !== url)
        fail(
          "This name already identifies another running terminal. Stop it or use a different name.",
        );
      process.stdout.write(
        `${JSON.stringify({ ok: true, session: existing.session, reused: true })}\n`,
      );
      return;
    }
  }
  const controlPath = controlPathForNewSession();
  const environment = { ...process.env };
  delete environment.SOMMELIER_URL;
  delete environment.AI_CDL_PAIR_URL;
  const child = spawn(
    process.execPath,
    [fileURLToPath(import.meta.url), "--daemon", "--control", controlPath],
    { detached: true, stdio: ["pipe", "pipe", "ignore"], env: environment },
  );
  child.stdin.end(`${url}\n`);
  const handshake = await new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error("Starting the Sommelier session timed out.")),
      30_000,
    );
    const reader = createInterface({ input: child.stdout });
    reader.once("line", (line) => {
      clearTimeout(timer);
      try {
        resolve(JSON.parse(line));
      } catch {
        reject(new Error("The Sommelier session produced an unreadable handshake."));
      }
    });
    child.once("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once("exit", () => {
      clearTimeout(timer);
      reject(new Error("The Sommelier session exited before it was ready."));
    });
  });
  if (!handshake.ok) fail(handshake.error ?? "The Sommelier session could not start.");
  child.stdout.destroy();
  child.unref();
  if (name) saveProfile(name, { url, session: controlPath });
  process.stdout.write(`${JSON.stringify({ ok: true, session: controlPath })}\n`);
}

async function main() {
  if (process.argv.includes("--daemon")) {
    await runDaemon(argument("--control"));
    return;
  }

  const command = process.argv[2];
  if (command === "docs-search" || command === "docs-get") {
    const { searchExcelDocs, getExcelDoc } = await import("./lib/excel-docs.mjs");
    const result =
      command === "docs-search"
        ? searchExcelDocs({ query: argument("--query"), limit: Number(argument("--limit", "5")) })
        : getExcelDoc({ id: argument("--id") });
    process.stdout.write(`${JSON.stringify({ ok: true, result })}\n`);
    return;
  }
  const name = profileName();
  if (command === "list") {
    let names = [];
    try {
      names = readdirSync(profileDirectory)
        .filter((file) => /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}\.json$/.test(file))
        .map((file) => file.slice(0, -5));
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
    const terminals = await Promise.all(
      names.map(async (name) => {
        const profile = readProfile(name);
        const status = profile?.session
          ? await send(profile.session, { command: "status" }, 2_000).catch(() => undefined)
          : undefined;
        return {
          name,
          state: status?.state ?? "stopped",
          taskPaneConnected: status?.taskPaneConnected ?? false,
        };
      }),
    );
    process.stdout.write(`${JSON.stringify({ ok: true, terminals })}\n`);
    return;
  }
  if (command === "forget") {
    if (!name) fail("--name is required.");
    const profile = readProfile(name);
    if (profile?.session)
      await send(profile.session, { command: "stop" }, 2_000).catch(() => undefined);
    rmSync(join(profileDirectory, `${name}.json`), { force: true });
    process.stdout.write(`${JSON.stringify({ ok: true })}\n`);
    return;
  }
  if (command === "start") {
    await start();
    return;
  }

  const controlPath = argument(
    "--session",
    name
      ? readProfile(name)?.session
      : (process.env.SOMMELIER_SESSION ?? process.env.AI_CDL_PAIR_SESSION),
  );
  if (!controlPath) fail("Set SOMMELIER_SESSION or pass --session with the path from `start`.");

  let payload;
  switch (command) {
    case "status":
      payload = { command: "status" };
      break;
    case "next":
      payload = {
        command: "next",
        timeoutMs: positiveInteger(argument("--timeout", "120000"), "--timeout"),
      };
      break;
    case "request": {
      const method = argument("--method");
      if (!method) fail("--method is required.");
      let params;
      try {
        params = JSON.parse(argument("--params", "{}"));
      } catch {
        fail("--params must be valid JSON.");
      }
      payload = {
        command: "request",
        method,
        params,
        timeoutMs: positiveInteger(argument("--timeout", "60000"), "--timeout"),
      };
      break;
    }
    case "reply": {
      const content = argument("--content");
      if (!content) fail("--content is required.");
      payload = { command: "reply", content };
      break;
    }
    case "stop":
      payload = { command: "stop" };
      break;
    default:
      fail("Usage: session.mjs docs-search|docs-get|start|status|next|request|reply|stop");
  }

  const response = await send(controlPath, payload);
  process.stdout.write(`${JSON.stringify(response)}\n`);
  if (response.ok === false) process.exitCode = 2;
}

await main().catch((error) => fail(error instanceof Error ? error.message : String(error)));
