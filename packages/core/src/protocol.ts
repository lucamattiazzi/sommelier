import { z } from "zod";

/** Current wire and trace protocol version. */
export const PROTOCOL_VERSION = "0.1" as const;

/** JSON-compatible value accepted by protocol boundaries. */
export type JsonValue =
  | null
  | boolean
  | number
  | string
  | JsonValue[]
  | { [key: string]: JsonValue };

/** Portable JSON Schema fragment used to advertise tool contracts. */
export type JsonSchema = Readonly<Record<string, JsonValue>>;

/** Tool classification used by policy decisions. */
export type ToolClassification = "read" | "write" | "structural";

/** Agent-visible tool declaration. */
export interface ToolCapability {
  readonly name: string;
  readonly description: string;
  readonly inputSchema: JsonSchema;
  readonly classification: ToolClassification;
}

/** A user message kept distinct from workbook/tool data. */
export interface UserProtocolMessage {
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly type: "user_message";
  readonly text: string;
  readonly selectionIncluded: boolean;
}

/** Result returned to an agent after a tool call. */
export interface ToolResultMessage {
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly type: "tool_result";
  readonly toolCallId: string;
  readonly status: "success" | "rejected" | "error";
  readonly result?: JsonValue;
  readonly error?: SerializedSommelierError;
}

/** Input supplied to an adapter for one agent iteration. */
export interface AgentTurnRequest {
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly sessionId: string;
  readonly turnId: string;
  readonly iteration: number;
  readonly userMessage: UserProtocolMessage;
  readonly tools: readonly ToolCapability[];
  readonly toolResults: readonly ToolResultMessage[];
}

/** Options every agent adapter must honor. */
export interface AgentRunOptions {
  readonly signal: AbortSignal;
}

/** Events that an adapter can yield during one iteration. */
export type AgentEvent =
  | {
      readonly protocolVersion: typeof PROTOCOL_VERSION;
      readonly type: "assistant_delta";
      readonly delta: string;
    }
  | {
      readonly protocolVersion: typeof PROTOCOL_VERSION;
      readonly type: "assistant_message";
      readonly content: string;
    }
  | {
      readonly protocolVersion: typeof PROTOCOL_VERSION;
      readonly type: "tool_call";
      readonly id: string;
      readonly name: string;
      readonly arguments: unknown;
    }
  | { readonly protocolVersion: typeof PROTOCOL_VERSION; readonly type: "completion" };

/** Vendor-neutral adapter implemented by an organization's existing agent transport. */
export interface AgentAdapter {
  runTurn(request: AgentTurnRequest, options: AgentRunOptions): AsyncIterable<AgentEvent>;
}

/** Safe, JSON-serializable public error representation. */
export interface SerializedSommelierError {
  readonly code: string;
  readonly message: string;
  readonly context: Readonly<Record<string, JsonValue>>;
  readonly probableCause?: string;
  readonly suggestedAction?: string;
  readonly documentationUrl?: string;
}

const protocolVersionSchema = z.literal(PROTOCOL_VERSION);

/** Runtime schema for a tool call returned by an untrusted agent. */
export const toolCallSchema = z.object({
  protocolVersion: protocolVersionSchema,
  type: z.literal("tool_call"),
  id: z.string().min(1).max(200),
  name: z.string().min(1).max(200),
  arguments: z.unknown(),
});

/** Runtime schema for agent events at transport boundaries. */
export const agentEventSchema = z.discriminatedUnion("type", [
  z.object({
    protocolVersion: protocolVersionSchema,
    type: z.literal("assistant_delta"),
    delta: z.string(),
  }),
  z.object({
    protocolVersion: protocolVersionSchema,
    type: z.literal("assistant_message"),
    content: z.string(),
  }),
  toolCallSchema,
  z.object({ protocolVersion: protocolVersionSchema, type: z.literal("completion") }),
]);

/** Convert a Zod schema to portable JSON Schema. */
export function toJsonSchema(schema: z.ZodType): JsonSchema {
  return z.toJSONSchema(schema, { target: "draft-7" }) as JsonSchema;
}
