import type { JsonValue, SerializedSommelierError } from "./protocol.js";

/** Stable Sommelier error codes. */
export type SommelierErrorCode =
  | "SOMMELIER_ACTIVE_TURN"
  | "SOMMELIER_AGENT_PROTOCOL_INVALID"
  | "SOMMELIER_APPROVAL_NOT_FOUND"
  | "SOMMELIER_BUDGET_EXCEEDED"
  | "SOMMELIER_CANCELLED"
  | "SOMMELIER_INVALID_STATE_TRANSITION"
  | "SOMMELIER_POLICY_DENIED"
  | "SOMMELIER_TIMEOUT"
  | "SOMMELIER_TOOL_ARGUMENTS_INVALID"
  | "SOMMELIER_TOOL_EXECUTION_FAILED"
  | "SOMMELIER_TOOL_NOT_FOUND";

/** Options for a structured, safe public error. */
export interface SommelierErrorOptions {
  readonly code: SommelierErrorCode | (string & {});
  readonly message: string;
  readonly context?: Readonly<Record<string, JsonValue>>;
  readonly probableCause?: string;
  readonly suggestedAction?: string;
  readonly documentationUrl?: string;
  readonly cause?: unknown;
}

/** Base error carrying a stable code and remediation fields. */
export class SommelierError extends Error {
  readonly code: string;
  readonly context: Readonly<Record<string, JsonValue>>;
  readonly probableCause: string | undefined;
  readonly suggestedAction: string | undefined;
  readonly documentationUrl: string | undefined;

  constructor(options: SommelierErrorOptions) {
    super(options.message, { cause: options.cause });
    this.name = "SommelierError";
    this.code = options.code;
    this.context = options.context ?? {};
    this.probableCause = options.probableCause;
    this.suggestedAction = options.suggestedAction;
    this.documentationUrl = options.documentationUrl;
  }

  /** Return a JSON-safe form that deliberately omits the cause chain. */
  toJSON(): SerializedSommelierError {
    return {
      code: this.code,
      message: this.message,
      context: this.context,
      ...(this.probableCause === undefined ? {} : { probableCause: this.probableCause }),
      ...(this.suggestedAction === undefined ? {} : { suggestedAction: this.suggestedAction }),
      ...(this.documentationUrl === undefined ? {} : { documentationUrl: this.documentationUrl }),
    };
  }
}

/** Convert an unknown failure into a safe structured error. */
export function asSommelierError(
  error: unknown,
  fallbackCode = "SOMMELIER_TOOL_EXECUTION_FAILED",
): SommelierError {
  if (error instanceof SommelierError) return error;
  return new SommelierError({
    code: fallbackCode,
    message: error instanceof Error ? error.message : "An unknown Sommelier error occurred.",
    suggestedAction: "Inspect the associated trace event and retry the operation.",
    cause: error,
  });
}
