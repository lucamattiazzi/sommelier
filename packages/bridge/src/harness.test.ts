import { expect, it, vi } from "vitest";
import { CodexAdapter, type HarnessRpc } from "./codex.js";
import { callExcelTool, createExcelMcpServer, excelTools } from "./mcp.js";
import { OpenCodeAdapter } from "./opencode.js";

it("exposes validated read/write/chart tools and local documentation", async () => {
  const call = vi.fn(async () => ({
    ok: true,
    result: { mode: "whole-workbook", pinnedRanges: [] },
  }));
  expect(excelTools.some((t) => t.name === "excel_chart_create")).toBe(true);
  await expect(
    callExcelTool(call, "excel_range_read", {
      range: { sheetId: "S", address: "A1" },
      unexpected: true,
    }),
  ).rejects.toThrow();
  expect(call).not.toHaveBeenCalled();
  expect(await callExcelTool(call, "excel_context_get", {})).toMatchObject({
    mode: "whole-workbook",
  });
  expect(await callExcelTool(call, "excel_guide", {})).toContain("untrusted");
  expect(createExcelMcpServer(call)).toBeDefined();
});

it("OpenCode registers MCP, binds a selected session and submits without replay", async () => {
  const fetcher = vi.fn(async (url: string | URL | Request, options?: RequestInit) => {
    const path = new URL(url instanceof Request ? url.url : url).pathname;
    if (path === "/session/status") return Response.json({});
    if (path === "/mcp") return Response.json({ pair_desk: { status: "connected" } });
    if (options?.method === "POST" && path.endsWith("/message"))
      return Response.json({ parts: [{ type: "text", text: "Done" }] });
    return Response.json({ id: "known" });
  });
  const adapter = new OpenCodeAdapter({
    origin: "http://127.0.0.1:4000",
    directory: "/tmp/project",
    name: "desk",
    command: ["node", "/test/agent.mjs", "mcp"],
    fetch: fetcher,
  });
  expect(await adapter.connect("known")).toBe("known");
  expect(await adapter.prompt("Read selection")).toBe("Done");
  expect(
    fetcher.mock.calls.filter(([u]) => String(u).includes("/session/known/message")),
  ).toHaveLength(1);
  expect(
    () =>
      new OpenCodeAdapter({
        origin: "http://example.com",
        directory: "/tmp",
        name: "desk",
        command: [],
      }),
  ).toThrow(/loopback/);
});

it("Codex resumes the stored thread with MCP configured and waits for matching completion", async () => {
  let listener: (method: string, params: unknown) => void = () => {};
  const request = vi.fn(async (method: string) => {
    if (method === "thread/resume") return { thread: { id: "known" } };
    if (method === "turn/start") {
      queueMicrotask(() => {
        listener("item/completed", {
          threadId: "other",
          item: { type: "agentMessage", text: "Wrong" },
        });
        listener("item/completed", {
          threadId: "known",
          turnId: "turn1",
          item: { type: "agentMessage", text: "Done" },
        });
        listener("turn/completed", {
          threadId: "known",
          turn: { id: "turn1", status: "completed", error: null },
        });
      });
      return { turn: { id: "turn1" } };
    }
    return {};
  });
  const rpc: HarnessRpc = {
    request,
    notify: vi.fn(),
    subscribe: (l) => {
      listener = l;
      return () => {};
    },
  };
  const adapter = new CodexAdapter(rpc, {
    directory: "/tmp",
    command: ["node", "/test/agent.mjs", "mcp"],
  });
  expect(await adapter.connect("known")).toBe("known");
  expect(request.mock.calls.map(([method]) => method)).toEqual(["initialize", "thread/resume"]);
  expect(request).toHaveBeenCalledWith(
    "thread/resume",
    expect.objectContaining({
      config: expect.objectContaining({
        "mcp_servers.ai_cdl_pair": expect.objectContaining({
          default_tools_approval_mode: "approve",
        }),
      }),
    }),
  );
  expect(await adapter.prompt("Read selection")).toBe("Done");
});
