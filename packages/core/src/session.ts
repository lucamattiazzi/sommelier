import { asSommelierError, SommelierError } from "./errors.js";
import type { SessionEvent, SessionEventBody, TraceRedaction, TraceSink } from "./events.js";
import { redactSessionEvent } from "./events.js";
import type { ApprovalPolicy } from "./policy.js";
import { defaultApprovalPolicy } from "./policy.js";
import type {
  AgentAdapter,
  AgentEvent,
  JsonValue,
  ToolCapability,
  ToolResultMessage,
  UserProtocolMessage,
} from "./protocol.js";
import { PROTOCOL_VERSION } from "./protocol.js";
import { type SessionState, SessionStateMachine } from "./state.js";
import type { ToolDefinition, ToolPreview } from "./tool.js";

/** Resource limits enforced by a session. */
export interface SessionLimits {
  readonly maxAgentIterations: number;
  readonly maxToolCalls: number;
  readonly maxCellsRead: number;
  readonly maxCellsWritten: number;
  readonly maxBytesReturned: number;
  readonly toolTimeoutMs: number;
  readonly turnTimeoutMs: number;
}

/** Input for a new user turn. */
export interface UserMessageInput {
  readonly text: string;
  readonly context?: { readonly includeSelection?: boolean };
  readonly limits?: Partial<SessionLimits>;
}

/** Successful turn output. */
export interface TurnResult {
  readonly sessionId: string;
  readonly turnId: string;
  readonly state: "completed";
  readonly assistantMessage: string;
  readonly toolCallCount: number;
}

/** Construction options for an agent session. */
export interface CreateAgentSessionOptions {
  readonly agent: AgentAdapter;
  readonly tools: readonly ToolDefinition[];
  readonly policy?: ApprovalPolicy;
  readonly limits?: Partial<SessionLimits>;
  readonly traceSinks?: readonly TraceSink[];
  readonly traceRedaction?: TraceRedaction;
  readonly createId?: () => string;
  readonly now?: () => Date;
}

/** Public session controller. */
export interface AgentSession {
  readonly id: string;
  readonly state: SessionState;
  sendMessage(input: UserMessageInput): Promise<TurnResult>;
  cancel(reason?: string): void;
  approve(requestId: string): Promise<void>;
  reject(requestId: string, reason?: string): Promise<void>;
  subscribe(listener: (event: SessionEvent) => void): () => void;
}

const defaultLimits: SessionLimits = {
  maxAgentIterations: 10,
  maxToolCalls: 20,
  maxCellsRead: 10_000,
  maxCellsWritten: 1_000,
  maxBytesReturned: 1_000_000,
  toolTimeoutMs: 15_000,
  turnTimeoutMs: 60_000,
};

interface PendingApproval {
  readonly resolve: (resolution: { approved: boolean; reason?: string }) => void;
}

class AgentSessionImpl implements AgentSession {
  readonly id: string;
  readonly #agent: AgentAdapter;
  readonly #tools: Map<string, ToolDefinition>;
  readonly #policy: ApprovalPolicy;
  readonly #baseLimits: SessionLimits;
  readonly #traceSinks: readonly TraceSink[];
  readonly #redaction: TraceRedaction;
  readonly #createId: () => string;
  readonly #now: () => Date;
  readonly #machine = new SessionStateMachine();
  readonly #listeners = new Set<(event: SessionEvent) => void>();
  readonly #approvals = new Map<string, PendingApproval>();
  readonly #startedEvent: SessionEvent;
  #turnAbort: AbortController | undefined;
  #turnId: string | undefined;
  #cancelReason = "The turn was cancelled by the host application.";

  constructor(options: CreateAgentSessionOptions) {
    this.#createId = options.createId ?? (() => crypto.randomUUID());
    this.#now = options.now ?? (() => new Date());
    this.id = this.#createId();
    this.#agent = options.agent;
    this.#tools = new Map(options.tools.map((tool) => [tool.name, tool]));
    if (this.#tools.size !== options.tools.length) {
      throw new SommelierError({
        code: "SOMMELIER_DUPLICATE_TOOL",
        message: "Tool names must be unique.",
      });
    }
    this.#policy = options.policy ?? defaultApprovalPolicy();
    this.#baseLimits = { ...defaultLimits, ...options.limits };
    this.#traceSinks = options.traceSinks ?? [];
    this.#redaction = options.traceRedaction ?? "metadata";
    this.#startedEvent = this.#emit({
      type: "session.started",
      payload: { state: this.#machine.state },
    });
  }

  get state(): SessionState {
    return this.#machine.state;
  }

  subscribe(listener: (event: SessionEvent) => void): () => void {
    this.#listeners.add(listener);
    listener(this.#startedEvent);
    return () => this.#listeners.delete(listener);
  }

  cancel(reason = "The turn was cancelled by the host application."): void {
    this.#cancelReason = reason;
    this.#turnAbort?.abort(new SommelierError({ code: "SOMMELIER_CANCELLED", message: reason }));
    for (const approval of this.#approvals.values()) approval.resolve({ approved: false, reason });
    this.#approvals.clear();
  }

  async approve(requestId: string): Promise<void> {
    this.#resolveApproval(requestId, { approved: true });
  }

  async reject(
    requestId: string,
    reason = "The user rejected this workbook mutation.",
  ): Promise<void> {
    this.#resolveApproval(requestId, { approved: false, reason });
  }

  async sendMessage(input: UserMessageInput): Promise<TurnResult> {
    if (
      [
        "running_agent",
        "executing_read_tool",
        "awaiting_approval",
        "executing_write_tool",
      ].includes(this.state)
    ) {
      throw new SommelierError({
        code: "SOMMELIER_ACTIVE_TURN",
        message: "This session already has an active turn.",
        suggestedAction: "Wait for the active turn or cancel it before sending another message.",
      });
    }

    const turnId = this.#createId();
    this.#turnId = turnId;
    this.#machine.transition("running_agent");
    const controller = new AbortController();
    this.#turnAbort = controller;
    const limits = { ...this.#baseLimits, ...input.limits };
    const turnTimer = setTimeout(
      () =>
        controller.abort(
          new SommelierError({ code: "SOMMELIER_TIMEOUT", message: "The turn timed out." }),
        ),
      limits.turnTimeoutMs,
    );
    const userMessage: UserProtocolMessage = {
      protocolVersion: PROTOCOL_VERSION,
      type: "user_message",
      text: input.text,
      selectionIncluded: input.context?.includeSelection ?? false,
    };
    this.#emit({
      type: "user.message",
      payload: { text: input.text, includeSelection: userMessage.selectionIncluded },
    });

    let toolCalls = 0;
    let cellsRead = 0;
    let cellsWritten = 0;
    let bytesReturned = 0;
    let assistantMessage = "";
    const toolResults: ToolResultMessage[] = [];

    try {
      for (let iteration = 1; iteration <= limits.maxAgentIterations; iteration += 1) {
        this.#throwIfAborted(controller.signal);
        this.#emit({ type: "agent.started", payload: { iteration } });
        let requestedTool = false;
        let completed = false;
        const events = this.#agent.runTurn(
          {
            protocolVersion: PROTOCOL_VERSION,
            sessionId: this.id,
            turnId,
            iteration,
            userMessage,
            tools: [...this.#tools.values()].map<ToolCapability>((tool) => ({
              name: tool.name,
              description: tool.description,
              classification: tool.classification,
              inputSchema: tool.inputSchema,
            })),
            toolResults,
          },
          { signal: controller.signal },
        );

        for await (const event of events) {
          this.#throwIfAborted(controller.signal);
          if (event.protocolVersion !== PROTOCOL_VERSION) {
            throw new SommelierError({
              code: "SOMMELIER_AGENT_PROTOCOL_INVALID",
              message: `Unsupported agent protocol version: ${event.protocolVersion}.`,
              suggestedAction: `Configure the adapter to use protocol ${PROTOCOL_VERSION}.`,
            });
          }
          if (event.type === "assistant_delta") {
            this.#emit({ type: "assistant.delta", payload: { delta: event.delta } });
          } else if (event.type === "assistant_message") {
            assistantMessage = event.content;
            this.#emit({ type: "assistant.message", payload: { content: event.content } });
          } else if (event.type === "tool_call") {
            requestedTool = true;
            toolCalls += 1;
            if (toolCalls > limits.maxToolCalls)
              this.#budgetError("tool calls", limits.maxToolCalls);
            const result = await this.#executeTool(event, controller.signal, limits, {
              cellsRead,
              cellsWritten,
              bytesReturned,
            });
            cellsRead += result.read;
            cellsWritten += result.written;
            bytesReturned += result.bytes;
            toolResults.push(result.message);
          } else {
            completed = true;
          }
        }

        if (completed && !requestedTool) {
          this.#machine.transition("completed");
          this.#emit({ type: "turn.completed", payload: { assistantMessage } });
          return {
            sessionId: this.id,
            turnId,
            state: "completed",
            assistantMessage,
            toolCallCount: toolCalls,
          };
        }
      }
      this.#budgetError("agent iterations", limits.maxAgentIterations);
    } catch (error) {
      const structured = this.#normalizeTurnError(error, controller.signal);
      if (structured.code === "SOMMELIER_CANCELLED" || structured.code === "SOMMELIER_TIMEOUT") {
        this.#machine.transition("cancelled");
        this.#emit({ type: "turn.cancelled", payload: { reason: structured.message } });
      } else {
        this.#machine.transition("failed");
        this.#emit({ type: "turn.failed", payload: { error: structured.toJSON() } });
      }
      throw structured;
    } finally {
      clearTimeout(turnTimer);
      this.#turnAbort = undefined;
      this.#turnId = undefined;
      this.#approvals.clear();
    }
    throw new SommelierError({
      code: "SOMMELIER_UNREACHABLE",
      message: "The turn ended unexpectedly.",
    });
  }

  async #executeTool(
    event: Extract<AgentEvent, { type: "tool_call" }>,
    turnSignal: AbortSignal,
    limits: SessionLimits,
    usage: {
      readonly cellsRead: number;
      readonly cellsWritten: number;
      readonly bytesReturned: number;
    },
  ): Promise<{ message: ToolResultMessage; read: number; written: number; bytes: number }> {
    const safeArguments = JSON.parse(JSON.stringify(event.arguments)) as JsonValue;
    this.#emit(
      { type: "tool.requested", payload: { name: event.name, arguments: safeArguments } },
      event.id,
    );
    const tool = this.#tools.get(event.name);
    if (!tool) {
      const error = new SommelierError({
        code: "SOMMELIER_TOOL_NOT_FOUND",
        message: `The agent requested unavailable tool ${event.name}.`,
        context: { toolName: event.name },
        suggestedAction: "Use one of the tool declarations supplied with the turn request.",
      });
      this.#emit(
        { type: "tool.failed", payload: { name: event.name, error: error.toJSON() } },
        event.id,
      );
      return { message: this.#errorResult(event.id, error), read: 0, written: 0, bytes: 0 };
    }
    const parsed = tool.input.safeParse(event.arguments);
    if (!parsed.success) {
      const error = new SommelierError({
        code: "SOMMELIER_TOOL_ARGUMENTS_INVALID",
        message: `Invalid arguments for ${event.name}.`,
        context: {
          issues: parsed.error.issues
            .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
            .join("; "),
        },
        suggestedAction: "Generate arguments that conform to the advertised JSON Schema.",
      });
      this.#emit(
        { type: "tool.failed", payload: { name: event.name, error: error.toJSON() } },
        event.id,
      );
      return { message: this.#errorResult(event.id, error), read: 0, written: 0, bytes: 0 };
    }
    this.#emit({ type: "tool.validated", payload: { name: event.name } }, event.id);
    const estimate = tool.estimateCells?.(parsed.data) ?? { read: 0, written: 0 };
    if (usage.cellsRead + estimate.read > limits.maxCellsRead)
      this.#budgetError("cells read", limits.maxCellsRead);
    if (usage.cellsWritten + estimate.written > limits.maxCellsWritten) {
      this.#budgetError("cells written", limits.maxCellsWritten);
    }
    const context = {
      signal: turnSignal,
      sessionId: this.id,
      turnId: this.#turnId ?? "",
      toolCallId: event.id,
    };
    const preview = tool.preview
      ? await tool.preview(parsed.data, context)
      : {
          summary: `${tool.name} will modify the workbook.`,
          affectedCells: estimate.written,
          warnings: [],
        };
    const decision = await this.#policy.decide({
      toolName: tool.name,
      classification: tool.classification,
      ...(tool.classification === "read" ? {} : { preview }),
    });
    if (decision.action === "deny") {
      const error = new SommelierError({
        code: "SOMMELIER_POLICY_DENIED",
        message: decision.reason,
        context: { toolName: tool.name },
      });
      return {
        message: this.#errorResult(event.id, error, "rejected"),
        read: 0,
        written: 0,
        bytes: 0,
      };
    }
    if (decision.action === "require_approval") {
      this.#machine.transition("awaiting_approval");
      const resolution = await this.#requestApproval(preview, event.id, turnSignal);
      if (!resolution.approved) {
        this.#machine.transition("running_agent");
        return {
          message: {
            protocolVersion: PROTOCOL_VERSION,
            type: "tool_result",
            toolCallId: event.id,
            status: "rejected",
            result: { message: resolution.reason ?? "The mutation was rejected." },
          },
          read: 0,
          written: 0,
          bytes: 0,
        };
      }
    }
    this.#machine.transition(
      tool.classification === "read" ? "executing_read_tool" : "executing_write_tool",
    );
    this.#emit({ type: "tool.started", payload: { name: tool.name } }, event.id);
    try {
      const result = await this.#withTimeout(
        tool.execute(parsed.data, context),
        limits.toolTimeoutMs,
        turnSignal,
      );
      const bytes = new TextEncoder().encode(JSON.stringify(result)).byteLength;
      if (usage.bytesReturned + bytes > limits.maxBytesReturned)
        this.#budgetError("bytes returned", limits.maxBytesReturned);
      this.#emit({ type: "tool.completed", payload: { name: tool.name, result } }, event.id);
      this.#machine.transition("running_agent");
      return {
        message: {
          protocolVersion: PROTOCOL_VERSION,
          type: "tool_result",
          toolCallId: event.id,
          status: "success",
          result,
        },
        read: estimate.read,
        written: estimate.written,
        bytes,
      };
    } catch (error) {
      const structured = asSommelierError(error);
      if (
        structured.code === "SOMMELIER_BUDGET_EXCEEDED" ||
        structured.code === "SOMMELIER_TIMEOUT"
      ) {
        throw structured;
      }
      this.#emit(
        { type: "tool.failed", payload: { name: tool.name, error: structured.toJSON() } },
        event.id,
      );
      this.#machine.transition("running_agent");
      return { message: this.#errorResult(event.id, structured), read: 0, written: 0, bytes: 0 };
    }
  }

  async #requestApproval(preview: ToolPreview, toolCallId: string, signal: AbortSignal) {
    const requestId = this.#createId();
    const resolution = new Promise<{ approved: boolean; reason?: string }>((resolve) => {
      this.#approvals.set(requestId, { resolve });
    });
    this.#emit({ type: "approval.requested", payload: { requestId, preview } }, toolCallId);
    const abort = new Promise<never>((_, reject) => {
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    });
    const result = await Promise.race([resolution, abort]);
    this.#approvals.delete(requestId);
    return result;
  }

  #resolveApproval(requestId: string, resolution: { approved: boolean; reason?: string }): void {
    const pending = this.#approvals.get(requestId);
    if (!pending) {
      throw new SommelierError({
        code: "SOMMELIER_APPROVAL_NOT_FOUND",
        message: `Approval request ${requestId} is not pending.`,
        suggestedAction: "Resolve only request IDs from approval.requested events.",
      });
    }
    this.#approvals.delete(requestId);
    this.#emit({
      type: "approval.resolved",
      payload: {
        requestId,
        resolution: resolution.approved ? "approved" : "rejected",
        ...(resolution.reason === undefined ? {} : { reason: resolution.reason }),
      },
    });
    pending.resolve(resolution);
  }

  #errorResult(
    toolCallId: string,
    error: SommelierError,
    status: "error" | "rejected" = "error",
  ): ToolResultMessage {
    return {
      protocolVersion: PROTOCOL_VERSION,
      type: "tool_result",
      toolCallId,
      status,
      error: error.toJSON(),
    };
  }

  #budgetError(resource: string, limit: number): never {
    throw new SommelierError({
      code: "SOMMELIER_BUDGET_EXCEEDED",
      message: `The turn exceeded its ${resource} budget of ${limit}.`,
      context: { resource, limit },
      suggestedAction: "Request a smaller range or increase the explicit session limit.",
    });
  }

  async #withTimeout<T>(promise: Promise<T>, timeoutMs: number, signal: AbortSignal): Promise<T> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(
        () =>
          reject(
            new SommelierError({ code: "SOMMELIER_TIMEOUT", message: "Tool execution timed out." }),
          ),
        timeoutMs,
      );
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    });
    try {
      return await Promise.race([promise, timeout]);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  #throwIfAborted(signal: AbortSignal): void {
    if (signal.aborted)
      throw (
        signal.reason ??
        new SommelierError({ code: "SOMMELIER_CANCELLED", message: this.#cancelReason })
      );
  }

  #normalizeTurnError(error: unknown, signal: AbortSignal): SommelierError {
    if (signal.aborted) {
      const reason = signal.reason;
      if (reason instanceof SommelierError) return reason;
      return new SommelierError({ code: "SOMMELIER_CANCELLED", message: this.#cancelReason });
    }
    return asSommelierError(error);
  }

  #emit(body: SessionEventBody, toolCallId?: string): SessionEvent {
    const event: SessionEvent = {
      ...body,
      protocolVersion: PROTOCOL_VERSION,
      eventId: this.#createId(),
      timestamp: this.#now().toISOString(),
      sessionId: this.id,
      ...(this.#turnId === undefined ? {} : { turnId: this.#turnId }),
      ...(toolCallId === undefined ? {} : { toolCallId }),
    };
    for (const listener of this.#listeners) listener(event);
    const persisted = redactSessionEvent(event, this.#redaction);
    for (const sink of this.#traceSinks) void sink.write(persisted);
    return event;
  }
}

/** Create an isolated session controller with safe approval defaults. */
export function createAgentSession(options: CreateAgentSessionOptions): AgentSession {
  return new AgentSessionImpl(options);
}
