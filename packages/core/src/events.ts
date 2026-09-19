import type { JsonValue, PROTOCOL_VERSION, SerializedSommelierError } from "./protocol.js";
import type { SessionState } from "./state.js";
import type { ToolPreview } from "./tool.js";

/** Common metadata included on every observable session event. */
export interface SessionEventEnvelope {
  readonly protocolVersion: typeof PROTOCOL_VERSION;
  readonly eventId: string;
  readonly timestamp: string;
  readonly sessionId: string;
  readonly turnId?: string;
  readonly toolCallId?: string;
}

/** Event-specific type and payload before common metadata is attached. */
export type SessionEventBody =
  | { readonly type: "session.started"; readonly payload: { readonly state: SessionState } }
  | {
      readonly type: "user.message";
      readonly payload: { readonly text: string; readonly includeSelection: boolean };
    }
  | { readonly type: "agent.started"; readonly payload: { readonly iteration: number } }
  | { readonly type: "assistant.delta"; readonly payload: { readonly delta: string } }
  | { readonly type: "assistant.message"; readonly payload: { readonly content: string } }
  | {
      readonly type: "tool.requested";
      readonly payload: { readonly name: string; readonly arguments: JsonValue };
    }
  | { readonly type: "tool.validated"; readonly payload: { readonly name: string } }
  | { readonly type: "tool.started"; readonly payload: { readonly name: string } }
  | {
      readonly type: "tool.completed";
      readonly payload: { readonly name: string; readonly result: JsonValue };
    }
  | {
      readonly type: "tool.failed";
      readonly payload: { readonly name: string; readonly error: SerializedSommelierError };
    }
  | {
      readonly type: "approval.requested";
      readonly payload: { readonly requestId: string; readonly preview: ToolPreview };
    }
  | {
      readonly type: "approval.resolved";
      readonly payload: {
        readonly requestId: string;
        readonly resolution: "approved" | "rejected";
        readonly reason?: string;
      };
    }
  | { readonly type: "turn.completed"; readonly payload: { readonly assistantMessage: string } }
  | { readonly type: "turn.cancelled"; readonly payload: { readonly reason: string } }
  | {
      readonly type: "turn.failed";
      readonly payload: { readonly error: SerializedSommelierError };
    };

/** JSON-serializable session event stream. */
export type SessionEvent = SessionEventEnvelope & SessionEventBody;

/** Sink for persisted or remote traces. */
export interface TraceSink {
  write(event: SessionEvent): void | Promise<void>;
}

/** Trace redaction strategy. */
export type TraceRedaction = "none" | "metadata";

/** Redact workbook-bearing payloads while retaining correlation metadata. */
export function redactSessionEvent(event: SessionEvent, mode: TraceRedaction): SessionEvent {
  if (mode === "none") return event;
  if (event.type === "tool.requested") {
    return { ...event, payload: { ...event.payload, arguments: { redacted: true } } };
  }
  if (event.type === "tool.completed") {
    return { ...event, payload: { ...event.payload, result: { redacted: true } } };
  }
  return event;
}
