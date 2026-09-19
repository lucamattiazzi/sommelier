import { describe, expect, it } from "vitest";
import { z } from "zod";
import { PROTOCOL_VERSION } from "./protocol.js";
import { createAgentSession } from "./session.js";
import { defineTool } from "./tool.js";

describe("createAgentSession", () => {
  it("executes reads and completes an agent loop", async () => {
    const agent = {
      async *runTurn(request: { iteration: number }) {
        if (request.iteration === 1) {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call" as const,
            id: "call-1",
            name: "test.read",
            arguments: {},
          };
        } else {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "assistant_message" as const,
            content: "Done",
          };
          yield { protocolVersion: PROTOCOL_VERSION, type: "completion" as const };
        }
      },
    };
    const tool = defineTool({
      name: "test.read",
      description: "read",
      classification: "read",
      input: z.object({}),
      execute: async () => ({ ok: true }),
    });
    const session = createAgentSession({ agent, tools: [tool] });
    const eventTypes: string[] = [];
    session.subscribe((event) => eventTypes.push(event.type));
    await expect(session.sendMessage({ text: "test" })).resolves.toMatchObject({
      assistantMessage: "Done",
      toolCallCount: 1,
    });
    expect(eventTypes[0]).toBe("session.started");
  });

  it("waits for write approval", async () => {
    let wrote = false;
    const agent = {
      async *runTurn(request: { iteration: number }) {
        if (request.iteration === 1) {
          yield {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_call" as const,
            id: "write-1",
            name: "test.write",
            arguments: {},
          };
        } else yield { protocolVersion: PROTOCOL_VERSION, type: "completion" as const };
      },
    };
    const tool = defineTool({
      name: "test.write",
      description: "write",
      classification: "write",
      input: z.object({}),
      execute: async () => {
        wrote = true;
        return { ok: true };
      },
      preview: async () => ({ summary: "Write", affectedCells: 1, warnings: [] }),
    });
    const session = createAgentSession({ agent, tools: [tool] });
    session.subscribe((event) => {
      if (event.type === "approval.requested") void session.approve(event.payload.requestId);
    });
    await session.sendMessage({ text: "write" });
    expect(wrote).toBe(true);
  });

  it("enforces tool-call budgets", async () => {
    const agent = {
      async *runTurn() {
        yield {
          protocolVersion: PROTOCOL_VERSION,
          type: "tool_call" as const,
          id: crypto.randomUUID(),
          name: "test.read",
          arguments: {},
        };
      },
    };
    const tool = defineTool({
      name: "test.read",
      description: "read",
      classification: "read",
      input: z.object({}),
      execute: async () => ({}),
    });
    const session = createAgentSession({ agent, tools: [tool], limits: { maxToolCalls: 1 } });
    await expect(session.sendMessage({ text: "loop" })).rejects.toMatchObject({
      code: "SOMMELIER_BUDGET_EXCEEDED",
    });
  });

  it("terminates the turn when returned-byte budgets are exceeded", async () => {
    const agent = {
      async *runTurn() {
        yield {
          protocolVersion: PROTOCOL_VERSION,
          type: "tool_call" as const,
          id: "large-read",
          name: "test.read",
          arguments: {},
        };
      },
    };
    const tool = defineTool({
      name: "test.read",
      description: "read",
      classification: "read",
      input: z.object({}),
      execute: async () => ({ content: "larger than one byte" }),
    });
    const session = createAgentSession({ agent, tools: [tool], limits: { maxBytesReturned: 1 } });
    await expect(session.sendMessage({ text: "read" })).rejects.toMatchObject({
      code: "SOMMELIER_BUDGET_EXCEEDED",
    });
    expect(session.state).toBe("failed");
  });
});
