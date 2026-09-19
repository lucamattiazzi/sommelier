#!/usr/bin/env node

import { HttpAgentAdapter } from "@lucamattiazzi/sommelier-agent-http";
import {
  type AgentAdapter,
  type AgentEvent,
  type AgentTurnRequest,
  PROTOCOL_VERSION,
} from "@lucamattiazzi/sommelier-core";
import { Command, InvalidArgumentError } from "commander";
import { runAdapterContract } from "./contract.js";

function positiveNumber(value: string): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new InvalidArgumentError("Timeout must be a positive number of milliseconds.");
  }
  return parsed;
}

const program = new Command()
  .name("sommelier-contract")
  .description("Validate an Sommelier agent adapter endpoint.")
  .version("0.1.0")
  .option("--endpoint <url>", "HTTP endpoint")
  .option("--scripted", "use the built-in deterministic protocol adapter")
  .option(
    "--timeout-ms <milliseconds>",
    "timeout for each endpoint contract case",
    positiveNumber,
    30_000,
  )
  .option("--json", "machine-readable output")
  .action(
    async (options: {
      endpoint?: string;
      scripted?: boolean;
      timeoutMs: number;
      json?: boolean;
    }) => {
      const adapter = options.endpoint
        ? new HttpAgentAdapter({ endpoint: options.endpoint, timeoutMs: options.timeoutMs })
        : options.scripted
          ? new ContractAdapter()
          : undefined;
      if (!adapter) {
        throw new Error("SOMMELIER_CONTRACT_TARGET_REQUIRED: Pass --endpoint <url> or --scripted.");
      }
      const report = await runAdapterContract(adapter, { timeoutMs: options.timeoutMs });
      console.log(
        options.json
          ? JSON.stringify(report, null, 2)
          : report.cases
              .map((entry) => `${entry.passed ? "PASS" : "FAIL"} ${entry.name}: ${entry.message}`)
              .join("\n"),
      );
      if (!report.passed) process.exitCode = 1;
    },
  );

class ContractAdapter implements AgentAdapter {
  async *runTurn(
    request: AgentTurnRequest,
    options: { signal: AbortSignal },
  ): AsyncIterable<AgentEvent> {
    if (options.signal.aborted) throw options.signal.reason;
    const prompt = request.userMessage.text;
    if (prompt.endsWith("malformed")) {
      yield { protocolVersion: "9.0", type: "completion" } as unknown as AgentEvent;
      return;
    }
    if (prompt.endsWith("timeout")) {
      await new Promise<void>((_, reject) =>
        options.signal.addEventListener("abort", () => reject(options.signal.reason), {
          once: true,
        }),
      );
    }
    if (prompt.endsWith("read-tool")) {
      yield {
        protocolVersion: PROTOCOL_VERSION,
        type: "tool_call",
        id: "contract-read",
        name: "excel.read_range",
        arguments: {},
      };
    } else if (prompt.endsWith("sequential-tools") || prompt.endsWith("tool-error")) {
      if (request.toolResults.length === 0) {
        yield {
          protocolVersion: PROTOCOL_VERSION,
          type: "tool_call",
          id: prompt.endsWith("tool-error") ? "contract-error" : "contract-sequential",
          name: "excel.read_range",
          arguments: {},
        };
      } else {
        yield {
          protocolVersion: PROTOCOL_VERSION,
          type: "assistant_message",
          content: "Contract response",
        };
        yield { protocolVersion: PROTOCOL_VERSION, type: "completion" };
      }
    } else if (prompt.endsWith("mutation-tool")) {
      yield {
        protocolVersion: PROTOCOL_VERSION,
        type: "tool_call",
        id: "contract-write",
        name: "excel.write_range",
        arguments: {},
      };
    } else if (prompt.endsWith("invalid-tool")) {
      yield {
        protocolVersion: PROTOCOL_VERSION,
        type: "tool_call",
        id: "bad",
        name: "excel.not_a_tool",
        arguments: {},
      };
    } else if (prompt.endsWith("invalid-arguments")) {
      yield {
        protocolVersion: PROTOCOL_VERSION,
        type: "tool_call",
        id: "bad-args",
        name: "excel.read_range",
        arguments: null,
      };
    } else {
      yield {
        protocolVersion: PROTOCOL_VERSION,
        type: "assistant_message",
        content: "Contract response",
      };
      yield { protocolVersion: PROTOCOL_VERSION, type: "completion" };
    }
  }
}

program.parseAsync().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
