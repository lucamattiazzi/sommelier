import { execFile } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { createPairIdentity, pairConnectionUrl } from "@lucamattiazzi/sommelier-transport";
import { expect, it } from "vitest";
import { z } from "zod";
import { createPairServer } from "../../../apps/server/src/index.js";
import { CodexAdapter } from "./codex.js";
import { startCodexProcess } from "./line-rpc.js";
import { OpenCodeAdapter } from "./opencode.js";
import { startOpenCodeProcess } from "./opencode-process.js";

// Explicit opt-in: starts installed harnesses and creates/archives a synthetic Codex thread.
// Model calls require a second opt-in: PAIR_NATIVE_TURN=codex|opencode|1. Data is synthetic.
it.skipIf(process.env.PAIR_NATIVE_SMOKE !== "1")(
  "loads native Excel tools, restores OpenCode and optionally restores a completed Codex thread",
  async () => {
    const dir = await mkdtemp(join(tmpdir(), "pair-native-"));
    const relay = await createPairServer({ port: 0, publicOrigin: "https://example.test" });
    const environment = { SOMMELIER_HOME: dir };
    const invoke = (command: string, url?: string) =>
      promisify(execFile)(
        process.execPath,
        ["skills/sommelier/scripts/session.mjs", command, "--name", "desk"],
        { env: { ...process.env, ...environment, ...(url ? { SOMMELIER_URL: url } : {}) } },
      );
    let rpc = startCodexProcess();
    let threadId: string | undefined;
    let openCode: Awaited<ReturnType<typeof startOpenCodeProcess>> | undefined;
    try {
      await invoke(
        "start",
        await pairConnectionUrl(
          `http://${relay.host}:${relay.port}`,
          createPairIdentity(),
          "agent",
        ),
      );
      const command = [
        process.execPath,
        resolve("packages/bridge/dist/agent.js"),
        "mcp",
        "--name",
        "desk",
      ];
      const codex = new CodexAdapter(rpc, { directory: dir, command, environment });
      threadId = await codex.connect();
      expect(threadId).toBeTruthy();
      const status = await rpc.request("mcpServerStatus/list", { threadId });
      expect(JSON.stringify(status).includes("excel_chart_create")).toBe(true);
      if (["1", "codex"].includes(process.env.PAIR_NATIVE_TURN ?? "")) {
        let calledGuide = false;
        const off = rpc.subscribe((method, params) => {
          const result = z
            .object({
              item: z.object({
                type: z.literal("mcpToolCall"),
                server: z.literal("ai_cdl_pair"),
                tool: z.literal("excel_guide"),
                status: z.literal("completed"),
              }),
            })
            .safeParse(params);
          if (method === "item/completed" && result.success) calledGuide = true;
        });
        expect(
          await codex.prompt(
            "Use only excel_guide, then reply exactly PAIR_SMOKE_OK. This is a synthetic integration test; do not use other tools.",
          ),
        ).toContain("PAIR_SMOKE_OK");
        off();
        expect(calledGuide).toBe(true);
      }
      if (["1", "codex"].includes(process.env.PAIR_NATIVE_TURN ?? "")) {
        // Codex persists its rollout after the first turn, not for an empty thread.
        rpc.close();
        rpc = startCodexProcess();
        const resumed = new CodexAdapter(rpc, { directory: dir, command, environment });
        expect(await resumed.connect(threadId)).toBe(threadId);
      }
      openCode = await startOpenCodeProcess(dir);
      const opencode = new OpenCodeAdapter({
        origin: openCode.origin,
        password: openCode.password,
        username: openCode.username,
        directory: dir,
        name: "desk",
        command,
        environment,
      });
      expect((await fetch(`${openCode.origin}/global/health`)).status).toBe(401);
      const openCodeSession = await opencode.connect();
      expect(openCodeSession).toBeTruthy();
      if (["1", "opencode"].includes(process.env.PAIR_NATIVE_TURN ?? ""))
        expect(
          await opencode.prompt(
            "Use only the Pair excel_guide tool, then reply exactly PAIR_SMOKE_OK. This is a synthetic integration test; do not use other tools.",
          ),
        ).toContain("PAIR_SMOKE_OK");
      await openCode.close();
      openCode = await startOpenCodeProcess(dir);
      const restoredOpenCode = new OpenCodeAdapter({
        origin: openCode.origin,
        password: openCode.password,
        username: openCode.username,
        directory: dir,
        name: "desk",
        command,
        environment,
      });
      expect(await restoredOpenCode.connect(openCodeSession)).toBe(openCodeSession);
      await fetch(`${openCode.origin}/session/${encodeURIComponent(openCodeSession)}`, {
        method: "DELETE",
        headers: {
          Authorization: `Basic ${Buffer.from(`${openCode.username}:${openCode.password}`).toString("base64")}`,
        },
      });
      console.log(
        `Native harness smoke passed (model turns: ${process.env.PAIR_NATIVE_TURN ?? "none"}).`,
      );
    } finally {
      if (threadId) await rpc.request("thread/archive", { threadId }).catch(() => {});
      rpc.close();
      await openCode?.close();
      await invoke("stop").catch(() => {});
      await relay.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
  120000,
);
