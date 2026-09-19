import { z } from "zod";
import { EXCEL_GUIDE } from "./mcp.js";

export interface HarnessRpc {
  request(method: string, params: unknown): Promise<unknown>;
  notify(method: string, params: unknown): void;
  subscribe(listener: (method: string, params: unknown) => void): () => void;
}
export interface CodexAdapterOptions {
  readonly directory: string;
  readonly command: readonly string[];
  readonly environment?: Record<string, string>;
}
export class CodexAdapter {
  #thread: string | undefined;
  constructor(
    readonly rpc: HarnessRpc,
    readonly options: CodexAdapterOptions,
  ) {}
  async connect(threadId?: string): Promise<string> {
    await this.rpc.request("initialize", {
      clientInfo: { name: "sommelier", title: "Sommelier", version: "0.2.0-beta.1" },
    });
    this.rpc.notify("initialized", {});
    const [command, ...args] = this.options.command;
    if (!command) throw new Error("MCP command is required.");
    const params = {
      cwd: this.options.directory,
      developerInstructions: EXCEL_GUIDE,
      config: {
        "mcp_servers.ai_cdl_pair": {
          command,
          args,
          env: this.options.environment ?? {},
          enabled: true,
          required: true,
          tool_timeout_sec: 125,
          default_tools_approval_mode: "approve",
        },
      },
      ...(threadId ? { threadId } : {}),
    };
    const result = z
      .object({ thread: z.object({ id: z.string().min(1) }) })
      .parse(await this.rpc.request(threadId ? "thread/resume" : "thread/start", params));
    this.#thread = result.thread.id;
    return result.thread.id;
  }
  async prompt(content: string): Promise<string> {
    const threadId = this.#thread;
    if (!threadId) throw new Error("Codex thread is not connected.");
    let turnId: string | undefined;
    const events: { method: string; params: unknown }[] = [];
    let settle: (error?: Error, text?: string) => void = () => {};
    const completed = new Promise<string>((resolve, reject) => {
      settle = (error, text) => (error ? reject(error) : resolve(text ?? ""));
    });
    // Attach rejection handling before awaiting turn/start.
    void completed.catch(() => {});
    const texts: string[] = [];
    const receive = (method: string, input: unknown) => {
      if (method === "pair.process.closed") {
        settle(new Error("Codex exited before completing the turn."));
        return;
      }
      if (!turnId) {
        events.push({ method, params: input });
        return;
      }
      const params = z
        .object({
          threadId: z.string(),
          turnId: z.string().optional(),
          item: z.object({ type: z.string(), text: z.string().optional() }).optional(),
          turn: z.object({ id: z.string(), status: z.string() }).optional(),
        })
        .safeParse(input);
      if (!params.success || params.data.threadId !== threadId) return;
      const p = params.data;
      if (
        method === "item/completed" &&
        p.turnId === turnId &&
        p.item?.type === "agentMessage" &&
        p.item.text
      )
        texts.push(p.item.text);
      if (method === "turn/completed" && p.turn?.id === turnId) {
        if (p.turn.status === "completed") settle(undefined, texts.join("\n"));
        else settle(new Error(`Codex turn ${p.turn.status}. Check the workbook before retrying.`));
      }
      if (method === "pair.process.closed")
        settle(new Error("Codex exited before completing the turn."));
    };
    const off = this.rpc.subscribe(receive);
    const timer = setTimeout(
      () =>
        settle(
          new Error("Codex turn timed out. Do not retry a mutation without inspecting Excel."),
        ),
      600_000,
    );
    try {
      const result = z.object({ turn: z.object({ id: z.string() }) }).parse(
        await this.rpc.request("turn/start", {
          threadId,
          input: [{ type: "text", text: content }],
        }),
      );
      turnId = result.turn.id;
      for (const event of events) receive(event.method, event.params);
      return await completed;
    } finally {
      clearTimeout(timer);
      off();
    }
  }
}
