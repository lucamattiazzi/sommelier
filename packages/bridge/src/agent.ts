#!/usr/bin/env node
import { execFile, spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { chmod, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { createInterface } from "node:readline/promises";
import { fileURLToPath } from "node:url";
import { parseArgs, promisify } from "node:util";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { bridgeCommand, subscribeBridge } from "./bridge.js";
import { CodexAdapter } from "./codex.js";
import { startCodexProcess } from "./line-rpc.js";
import { createExcelMcpServer } from "./mcp.js";
import { OpenCodeAdapter } from "./opencode.js";
import { startOpenCodeProcess } from "./opencode-process.js";
import {
  lockProfile,
  profileDirectory,
  profileNameSchema,
  readBinding,
  saveBinding,
} from "./profiles.js";

const run = promisify(execFile);
const self = fileURLToPath(import.meta.url);
const sessionScript = fileURLToPath(new URL("./skill/scripts/session.mjs", import.meta.url));
const { values: flags, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    name: { type: "string", default: "desk" },
    directory: { type: "string" },
    server: { type: "string" },
    session: { type: "string" },
    out: { type: "string" },
    channel: { type: "boolean" },
    "include-content": { type: "boolean" },
    "new-session": { type: "boolean" },
    help: { type: "boolean" },
  },
});
const name = profileNameSchema.parse(flags.name);
const invoke = async (...args: string[]) => {
  const { stdout } = await run(process.execPath, [sessionScript, ...args, "--name", name], {
    env: process.env,
    timeout: 35000,
  });
  return JSON.parse(stdout) as unknown;
};
async function pathForBridge(): Promise<string> {
  const profile = z
    .object({ session: z.string(), url: z.string().url() })
    .parse(JSON.parse(await readFile(join(profileDirectory(), `${name}.json`), "utf8")));
  const url = new URL(profile.url);
  if (
    url.pathname !== "/connect" ||
    url.searchParams.get("role") !== "agent" ||
    !/^[\w-]{43}$/.test(url.hash.slice(1))
  )
    throw new Error("Adapters require an encrypted Sommelier profile.");
  return profile.session;
}
async function startBridge(): Promise<string> {
  await invoke("start");
  return pathForBridge();
}
const messageSchema = z.object({
  type: z.enum(["message", "idle", "closed"]),
  message: z.object({ content: z.string().min(1).max(32000), messageId: z.string() }).optional(),
});
async function main() {
  const command = positionals[0];
  if (!command || flags.help) {
    console.log(
      "Usage: sommelier pair|list|status|stop|forget|codex|opencode|claude|mcp|trace --name desk\nFirst: pair --name desk (paste the private URL once). Then: codex|opencode|claude --name desk.\nOptions: --session ID --new-session --directory PATH; OpenCode starts automatically; optional --server attaches an existing loopback API; trace --out FILE [--include-content].\nMCP-only: mcp --name desk [--channel for Claude].",
    );
    return;
  }
  if (command === "pair") {
    const reader = createInterface({ input: process.stdin, output: process.stderr });
    const url = (
      await reader.question("Paste the private connection URL from Excel (stored locally): ")
    ).trim();
    reader.close();
    const parsed = new URL(url);
    if (
      parsed.pathname !== "/connect" ||
      parsed.searchParams.get("role") !== "agent" ||
      !/^[\w-]{43}$/.test(parsed.hash.slice(1))
    )
      throw new Error("Use an encrypted agent connection URL from the TaskPane.");
    process.env.SOMMELIER_URL = url;
    try {
      await invoke("start");
    } finally {
      delete process.env.SOMMELIER_URL;
    }
    console.log(
      `Saved ${name}. Run sommelier codex|opencode|claude --name ${name}, then reconnect in Excel.`,
    );
    return;
  }
  if (["list", "status", "stop", "forget"].includes(command)) {
    if (command === "forget") {
      const unlock = await lockProfile(name);
      try {
        await invoke("forget");
        await rm(join(profileDirectory(), "adapters", `${name}.json`), { force: true });
      } finally {
        await unlock();
      }
    } else console.log(JSON.stringify(await invoke(command), null, 2));
    return;
  }
  if (command === "trace") {
    if (!flags.out)
      throw new Error("trace requires --out FILE. Content recording requires --include-content.");
    const { open } = await import("node:fs/promises");
    const file = await open(resolve(flags.out), "wx", 0o600);
    const path = await pathForBridge();
    let writes = Promise.resolve();
    let count = 0;
    const off = subscribeBridge(
      path,
      (event) => {
        if (++count > 1000) {
          off();
          console.error("Trace writer could not keep up; subscription stopped.");
          return;
        }
        writes = writes
          .then(async () => {
            await file.write(`${JSON.stringify(event)}\n`);
          })
          .catch(() => {
            off();
            console.error("Trace write failed; subscription stopped.");
          })
          .finally(() => {
            count--;
          });
      },
      {
        includeContent: flags["include-content"] === true,
        onError: (error) => {
          console.error(error.message);
          process.exitCode = 1;
        },
      },
    );
    const close = async () => {
      off();
      await writes;
      await file.close();
    };
    process.once("SIGINT", () => {
      void close();
    });
    process.once("SIGTERM", () => {
      void close();
    });
    return;
  }
  if (command === "mcp") {
    const path = await pathForBridge();
    const abort = new AbortController();
    const call = (payload: Parameters<typeof bridgeCommand>[1]) =>
      bridgeCommand(path, payload, 125000, abort.signal);
    const unlock = flags.channel ? await lockProfile(name) : undefined;
    const server = createExcelMcpServer(call, flags.channel === true);
    const transport = new StdioServerTransport();
    let closed = false;
    server.onclose = () => {
      closed = true;
      abort.abort();
      void unlock?.();
    };
    const initialized = new Promise<void>((resolve) => {
      server.oninitialized = resolve;
    });
    await server.connect(transport);
    if (flags.channel) {
      await initialized;
      try {
        while (!closed) {
          const next = messageSchema.parse(await call({ command: "next", timeoutMs: 30000 }));
          if (next.type === "closed") break;
          if (next.type === "message" && next.message) {
            await call({ command: "begin" });
            await server.notification({
              method: "notifications/claude/channel",
              params: {
                content: next.message.content,
                meta: { messageId: next.message.messageId, profile: name },
              },
            });
          }
        }
      } finally {
        await server.close();
        await unlock?.();
      }
    }
    return;
  }
  if (!["codex", "opencode", "claude"].includes(command))
    throw new Error("Unknown command. Use --help.");
  const harness = z.enum(["codex", "opencode", "claude"]).parse(command);
  const binding = await readBinding(name);
  if (binding && binding.harness !== harness && !flags["new-session"])
    throw new Error(
      `This profile belongs to ${binding.harness}. Use that adapter or --new-session.`,
    );
  const directory = resolve(flags.directory ?? binding?.directory ?? process.cwd());
  if (binding && binding.directory !== directory && !flags["new-session"])
    throw new Error(
      "The saved conversation belongs to another directory. Use --new-session to change it.",
    );
  const nativeId = flags.session ?? (flags["new-session"] ? undefined : binding?.nativeId);
  const path = await startBridge();
  const environment = { SOMMELIER_HOME: profileDirectory() };
  const mcpCommand = [process.execPath, self, "mcp", "--name", name];
  if (harness === "claude") {
    const native = nativeId ?? randomUUID();
    const dir = join(profileDirectory(), "adapters");
    await mkdir(dir, { recursive: true, mode: 0o700 });
    const config = join(dir, `${name}.mcp.json`);
    await writeFile(
      config,
      JSON.stringify({
        mcpServers: {
          ai_cdl_pair: {
            command: process.execPath,
            args: [...mcpCommand.slice(1), "--channel"],
            env: environment,
          },
        },
      }),
      { mode: 0o600 },
    );
    await chmod(config, 0o600);
    console.error(
      "Claude Channels is a research preview. Accept Claude's channel consent to enable Excel input. Keep this session open.",
    );
    const child = spawn(
      "claude",
      [
        "--mcp-config",
        config,
        "--dangerously-load-development-channels",
        "server:ai_cdl_pair",
        ...(nativeId ? ["--resume", native] : ["--session-id", native]),
      ],
      { cwd: directory, stdio: "inherit" },
    );
    await new Promise<void>((resolve, reject) => {
      child.once("spawn", resolve);
      child.once("error", reject);
    });
    await saveBinding(name, { harness, directory, nativeId: native });
    const exit = await new Promise<number>((resolve) =>
      child.once("exit", (code) => resolve(code ?? 1)),
    );
    await invoke("stop").catch(() => {});
    process.exitCode = exit;
    return;
  }
  const unlock = await lockProfile(name);
  const call = (payload: Parameters<typeof bridgeCommand>[1]) => bridgeCommand(path, payload);
  let close: () => void | Promise<void> = () => {};
  const startup = new AbortController();
  let stopping = false;
  const stop = () => {
    stopping = true;
    startup.abort();
    void close();
  };
  process.once("SIGINT", stop);
  process.once("SIGTERM", stop);
  try {
    let adapter: CodexAdapter | OpenCodeAdapter;
    let origin: string | undefined;
    if (harness === "codex") {
      const rpc = startCodexProcess();
      close = () => rpc.close();
      adapter = new CodexAdapter(rpc, { directory, command: mcpCommand, environment });
    } else {
      origin = flags.server ?? (flags["new-session"] ? undefined : binding?.origin);
      let endpoint: { origin: string; password?: string; username?: string };
      if (origin) {
        endpoint = {
          origin,
          ...(process.env.OPENCODE_SERVER_PASSWORD
            ? { password: process.env.OPENCODE_SERVER_PASSWORD }
            : {}),
          ...(process.env.OPENCODE_SERVER_USERNAME
            ? { username: process.env.OPENCODE_SERVER_USERNAME }
            : {}),
        };
      } else {
        const owned = await startOpenCodeProcess(directory, startup.signal);
        close = owned.close;
        endpoint = owned;
      }
      adapter = new OpenCodeAdapter({
        ...endpoint,
        directory,
        name,
        command: mcpCommand,
        environment,
      });
    }
    const session = await adapter.connect(nativeId);
    await saveBinding(name, {
      harness,
      directory,
      nativeId: session,
      ...(origin ? { origin } : {}),
    });
    console.error(
      `Excel ${name} → ${harness} ${session}. Waiting for TaskPane messages. Ctrl+C stops the adapter; the association is remembered.`,
    );
    while (!stopping) {
      const next = messageSchema.parse(await call({ command: "next", timeoutMs: 1000 }));
      if (next.type === "closed") break;
      if (next.type !== "message" || !next.message) continue;
      try {
        await call({ command: "begin" });
        const text = await adapter.prompt(next.message.content);
        if (text.trim())
          for (let offset = 0; offset < text.length; offset += 32000)
            await call({ command: "reply", content: text.slice(offset, offset + 32000) });
      } catch (error) {
        console.error(error instanceof Error ? error.message : "Agent turn failed.");
        await call({
          command: "reply",
          content:
            "The agent could not finish this request. Check the local adapter terminal and the workbook before retrying.",
        }).catch(() => {});
        // A timeout may leave a native turn running. Stop rather than start overlapping turns.
        throw error;
      }
    }
  } finally {
    await close();
    await invoke("stop").catch(() => {});
    await unlock();
  }
}
void main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : "Sommelier adapter failed.");
  process.exitCode = 1;
});
