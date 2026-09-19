import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { z } from "zod";
import type { HarnessRpc } from "./codex.js";

/** Codex's newline-delimited app-server protocol. No shell interpolation. */
export function startCodexProcess(): HarnessRpc & { close(): void } {
  const child = spawn("codex", ["app-server"], { stdio: ["pipe", "pipe", "inherit"] });
  const reader = createInterface({ input: child.stdout });
  const pending = new Map<
    number,
    {
      resolve(value: unknown): void;
      reject(error: Error): void;
      timer: ReturnType<typeof setTimeout>;
    }
  >();
  const listeners = new Set<(method: string, params: unknown) => void>();
  let nextId = 1;
  let closed = false;
  const write = (value: unknown) => child.stdin.write(`${JSON.stringify(value)}\n`);
  child.stdin.on("error", () => {});
  const end = () => {
    closed = true;
    for (const entry of pending.values()) {
      clearTimeout(entry.timer);
      entry.reject(new Error("Codex App Server exited."));
    }
    pending.clear();
    for (const listener of listeners) listener("pair.process.closed", {});
  };
  child.once("error", end);
  child.once("exit", end);
  reader.on("line", (line) => {
    const parsed = z
      .object({
        id: z.union([z.number(), z.string()]).optional(),
        method: z.string().optional(),
        params: z.unknown().optional(),
        result: z.unknown().optional(),
        error: z.object({ message: z.string() }).optional(),
      })
      .safeParse(
        (() => {
          try {
            return JSON.parse(line);
          } catch {
            return null;
          }
        })(),
      );
    if (!parsed.success) return;
    const message = parsed.data;
    if (message.method && message.id !== undefined) {
      // Native filesystem/command approvals must never inherit the Excel-only approval waiver.
      if (
        ["item/commandExecution/requestApproval", "item/fileChange/requestApproval"].includes(
          message.method,
        )
      ) {
        console.error(
          "Sommelier declined a native command/file approval. Excel approvals remain in the TaskPane.",
        );
        write({ id: message.id, result: { decision: "decline" } });
      } else if (message.method === "mcpServer/elicitation/request") {
        write({ id: message.id, result: { action: "cancel", content: null, _meta: null } });
      } else {
        console.error(`Unsupported Codex interaction: ${message.method}`);
        write({
          id: message.id,
          error: {
            code: -32601,
            message: "This Sommelier adapter does not handle this interactive request.",
          },
        });
      }
    } else if (message.method) {
      for (const listener of listeners) listener(message.method, message.params);
    } else if (typeof message.id === "number") {
      const entry = pending.get(message.id);
      if (!entry) return;
      clearTimeout(entry.timer);
      pending.delete(message.id);
      if (message.error) entry.reject(new Error(message.error.message));
      else entry.resolve(message.result);
    }
  });
  return {
    request(method, params) {
      if (closed) return Promise.reject(new Error("Codex App Server is closed."));
      const id = nextId++;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          reject(new Error(`Codex ${method} timed out.`));
        }, 125_000);
        pending.set(id, { resolve, reject, timer });
        write({ id, method, params });
      });
    },
    notify(method, params) {
      write({ method, params });
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    close() {
      reader.close();
      child.kill();
      end();
    },
  };
}
