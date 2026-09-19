import {
  type AgentAdapter,
  type AgentEvent,
  type AgentTurnRequest,
  PROTOCOL_VERSION,
} from "@lucamattiazzi/sommelier-core";
import { describe, expect, it } from "vitest";
import { runAdapterContract } from "./contract.js";

describe("runAdapterContract", () => {
  it("uses the configured case timeout", async () => {
    const adapter: AgentAdapter = {
      async *runTurn(_request: AgentTurnRequest, options: { signal: AbortSignal }) {
        await new Promise<void>((_, reject) => {
          options.signal.addEventListener("abort", () => reject(options.signal.reason), {
            once: true,
          });
        });
      },
    };

    const startedAt = Date.now();
    const report = await runAdapterContract(adapter, { timeoutMs: 10 });

    expect(Date.now() - startedAt).toBeLessThan(1_000);
    expect(report.cases[0]).toMatchObject({
      name: "assistant-only completion",
      passed: false,
      message: "Contract case timed out after 10ms.",
    });
  });

  it("rejects non-positive timeouts", async () => {
    const adapter = { async *runTurn() {} } satisfies AgentAdapter;
    await expect(runAdapterContract(adapter, { timeoutMs: 0 })).rejects.toThrow(
      "SOMMELIER_CONTRACT_TIMEOUT_INVALID",
    );
  });

  it("echoes endpoint-generated tool call ids when continuing a turn", async () => {
    const continuedIds: string[] = [];
    const turnIds = new Map<string, string>();
    const complete = (): AgentEvent[] => [
      {
        protocolVersion: PROTOCOL_VERSION,
        type: "assistant_message",
        content: "Synthetic contract response",
      },
      { protocolVersion: PROTOCOL_VERSION, type: "completion" },
    ];
    const adapter: AgentAdapter = {
      async *runTurn(request: AgentTurnRequest, options: { signal: AbortSignal }) {
        if (options.signal.aborted) throw options.signal.reason;
        const prompt = request.userMessage.text.replace("Sommelier contract: ", "");
        const priorTurnId = turnIds.get(prompt);
        if (priorTurnId && priorTurnId !== request.turnId) {
          throw new Error(`Continuation changed turn id for ${prompt}.`);
        }
        turnIds.set(prompt, request.turnId);
        if (prompt === "timeout") {
          await new Promise<void>((_, reject) =>
            options.signal.addEventListener("abort", () => reject(options.signal.reason), {
              once: true,
            }),
          );
          return;
        }
        if (prompt === "malformed") {
          yield { protocolVersion: "9.0", type: "completion" } as unknown as AgentEvent;
          return;
        }
        if (prompt === "read-tool") {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call",
            id: "model-read-71",
            name: "excel.read_range",
            arguments: {},
          };
          return;
        }
        if (prompt === "mutation-tool") {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call",
            id: "model-write-72",
            name: "excel.write_range",
            arguments: {},
          };
          return;
        }
        if (prompt === "sequential-tools" || prompt === "tool-error") {
          if (request.toolResults.length === 0) {
            yield {
              protocolVersion: PROTOCOL_VERSION,
              type: "tool_call",
              id: prompt === "sequential-tools" ? "model-sequential-73" : "model-error-74",
              name: "excel.read_range",
              arguments: {},
            };
          } else {
            continuedIds.push(request.toolResults[0]?.toolCallId ?? "missing");
            yield* complete();
          }
          return;
        }
        if (prompt === "invalid-tool") {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call",
            id: "model-invalid-75",
            name: "excel.not_a_tool",
            arguments: {},
          };
          return;
        }
        if (prompt === "invalid-arguments") {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call",
            id: "model-invalid-arguments-76",
            name: "excel.read_range",
            arguments: null,
          };
          return;
        }
        yield* complete();
      },
    };

    const report = await runAdapterContract(adapter);

    expect(report.passed).toBe(true);
    expect(continuedIds).toEqual(["model-sequential-73", "model-error-74"]);
    expect(new Set(turnIds.values()).size).toBe(turnIds.size);
  });
});
