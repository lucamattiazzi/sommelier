import {
  type AgentAdapter,
  type AgentEvent,
  type AgentTurnRequest,
  agentEventSchema,
  type JsonValue,
  PROTOCOL_VERSION,
  type ToolResultMessage,
} from "@lucamattiazzi/sommelier-core";

/** One adapter contract check. */
export interface ContractCaseResult {
  readonly name: string;
  readonly passed: boolean;
  readonly message: string;
}

/** Complete reusable agent-adapter contract report. */
export interface AdapterContractReport {
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly passed: boolean;
  readonly cases: readonly ContractCaseResult[];
}

/** Runtime limits for adapter contract cases. */
export interface AdapterContractOptions {
  /** Maximum duration of each case. Defaults to 30 seconds. */
  readonly timeoutMs?: number;
}

const capabilities = [
  {
    name: "excel.read_range",
    description: "read",
    classification: "read" as const,
    inputSchema: { type: "object" as const },
  },
  {
    name: "excel.write_range",
    description: "write",
    classification: "write" as const,
    inputSchema: { type: "object" as const },
  },
];

function request(prompt: string, toolResults: readonly ToolResultMessage[] = []): AgentTurnRequest {
  return {
    protocolVersion: PROTOCOL_VERSION,
    sessionId: "contract-session",
    turnId: `contract-${prompt}`,
    iteration: toolResults.length + 1,
    userMessage: {
      protocolVersion: PROTOCOL_VERSION,
      type: "user_message",
      text: `Sommelier contract: ${prompt}`,
      selectionIncluded: false,
    },
    tools: capabilities,
    toolResults,
  };
}

async function collect(
  adapter: AgentAdapter,
  input: AgentTurnRequest,
  signal: AbortSignal,
): Promise<AgentEvent[]> {
  const output: AgentEvent[] = [];
  for await (const event of adapter.runTurn(input, { signal }))
    output.push(agentEventSchema.parse(event));
  return output;
}

/** Run the vendor-neutral protocol contract against an adapter or HTTP endpoint wrapper. */
export async function runAdapterContract(
  adapter: AgentAdapter,
  options: AdapterContractOptions = {},
): Promise<AdapterContractReport> {
  const timeoutMs = options.timeoutMs ?? 30_000;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new Error("SOMMELIER_CONTRACT_TIMEOUT_INVALID: timeoutMs must be a positive number.");
  }
  const cases: ContractCaseResult[] = [];
  const check = async (name: string, execute: (signal: AbortSignal) => Promise<boolean>) => {
    const controller = new AbortController();
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => {
        const error = new Error(`Contract case timed out after ${timeoutMs}ms.`);
        controller.abort(error);
        reject(error);
      }, timeoutMs);
    });
    try {
      const passed = await Promise.race([execute(controller.signal), timeout]);
      cases.push({ name, passed, message: passed ? "passed" : "unexpected event sequence" });
    } catch (error) {
      cases.push({
        name,
        passed: false,
        message: error instanceof Error ? error.message : String(error),
      });
    } finally {
      if (timer) clearTimeout(timer);
    }
  };

  await check("assistant-only completion", async (signal) => {
    const events = await collect(adapter, request("assistant-only"), signal);
    return (
      events.some((event) => event.type === "assistant_message") &&
      events.some((event) => event.type === "completion")
    );
  });
  await check("one read tool call", async (signal) =>
    (await collect(adapter, request("read-tool"), signal)).some(
      (event) => event.type === "tool_call" && event.name === "excel.read_range",
    ),
  );
  await check("one write tool call", async (signal) =>
    (await collect(adapter, request("mutation-tool"), signal)).some(
      (event) => event.type === "tool_call" && event.name === "excel.write_range",
    ),
  );
  await check("multiple sequential tool calls", async (signal) => {
    const initial = await collect(adapter, request("sequential-tools"), signal);
    const toolCall = initial.find((event) => event.type === "tool_call");
    if (!toolCall || toolCall.type !== "tool_call") return false;
    const result: ToolResultMessage = {
      protocolVersion: PROTOCOL_VERSION,
      type: "tool_result",
      toolCallId: toolCall.id,
      status: "success",
      result: { ok: true },
    };
    const events = await collect(adapter, request("sequential-tools", [result]), signal);
    return events.some((event) => event.type === "tool_call" || event.type === "completion");
  });
  await check("invalid tool name", async (signal) =>
    (await collect(adapter, request("invalid-tool"), signal)).some(
      (event) => event.type === "tool_call" && event.name === "excel.not_a_tool",
    ),
  );
  await check("invalid arguments", async (signal) =>
    (await collect(adapter, request("invalid-arguments"), signal)).some(
      (event) => event.type === "tool_call" && event.arguments === null,
    ),
  );
  await check("tool error propagation", async (signal) => {
    const initial = await collect(adapter, request("tool-error"), signal);
    const toolCall = initial.find((event) => event.type === "tool_call");
    if (!toolCall || toolCall.type !== "tool_call") return false;
    const result: ToolResultMessage = {
      protocolVersion: PROTOCOL_VERSION,
      type: "tool_result",
      toolCallId: toolCall.id,
      status: "error",
      error: { code: "TEST", message: "synthetic", context: {} },
    };
    return (await collect(adapter, request("tool-error", [result]), signal)).some(
      (event) => event.type === "assistant_message",
    );
  });
  await check("cancellation", async () => {
    const controller = new AbortController();
    controller.abort(new Error("contract cancellation"));
    try {
      await collect(adapter, request("cancellation"), controller.signal);
      return false;
    } catch {
      return true;
    }
  });
  await check("timeout", async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new Error("contract timeout")), 25);
    try {
      await collect(adapter, request("timeout"), controller.signal);
      return false;
    } catch {
      return true;
    } finally {
      clearTimeout(timer);
    }
  });
  await check("malformed response", async (signal) => {
    try {
      await collect(adapter, request("malformed"), signal);
      return false;
    } catch {
      return true;
    }
  });
  return { protocolVersion: PROTOCOL_VERSION, passed: cases.every((entry) => entry.passed), cases };
}

/** Assert a JSON-compatible range matrix with a compact diff on failure. */
export function assertRangeValues(
  actual: readonly (readonly JsonValue[])[],
  expected: readonly (readonly JsonValue[])[],
): void {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `Range values differ.\nExpected: ${JSON.stringify(expected)}\nActual: ${JSON.stringify(actual)}`,
    );
  }
}
