import { connect } from "node:net";
import { createInterface } from "node:readline";
import { type JsonValue, jsonValueSchema } from "@lucamattiazzi/sommelier-protocol";
import { z } from "zod";

const responseSchema = z
  .object({ ok: z.boolean(), error: z.object({ message: z.string() }).optional() })
  .catchall(jsonValueSchema);
export type BridgeResponse = z.infer<typeof responseSchema>;

/** One local request. Never retries a workbook mutation. */
export function bridgeCommand(
  path: string,
  payload: Record<string, JsonValue>,
  timeoutMs = 125_000,
  signal?: AbortSignal,
): Promise<BridgeResponse> {
  return new Promise((resolve, reject) => {
    const socket = connect({ path, ...(signal ? { signal } : {}) });
    const reader = createInterface({ input: socket });
    const timer = setTimeout(
      () => finish(new Error("Sommelier request timed out. Inspect the workbook before retrying.")),
      timeoutMs,
    );
    function finish(error?: Error, result?: BridgeResponse) {
      clearTimeout(timer);
      reader.close();
      socket.destroy();
      if (error) reject(error);
      else if (result) resolve(result);
    }
    reader.on("error", () => finish(new Error("Bridge request cancelled or disconnected.")));
    socket.once("error", () =>
      finish(new Error("Bridge unavailable. Reconnect the named profile.")),
    );
    socket.once("end", () => finish(new Error("Bridge closed without a response.")));
    socket.once("connect", () => socket.write(`${JSON.stringify(payload)}\n`));
    reader.once("line", (line) => {
      try {
        const result = responseSchema.parse(JSON.parse(line));
        if (!result.ok) finish(new Error(result.error?.message ?? "Bridge request failed."));
        else finish(undefined, result);
      } catch {
        finish(new Error("Invalid bridge response."));
      }
    });
  });
}

export const bridgeTraceSchema = z.object({
  kind: z.enum(["rpc.request", "rpc.response", "event", "connection"]),
  occurredAt: z.string(),
  id: z.string().optional(),
  method: z.string().optional(),
  event: z.string().optional(),
  sender: z.enum(["user", "agent"]).optional(),
  messageId: z.string().optional(),
  outcome: z.enum(["success", "error"]).optional(),
  errorCode: z.string().optional(),
  durationMs: z.number().nonnegative().optional(),
  payload: jsonValueSchema.optional(),
});
export type BridgeTrace = z.infer<typeof bridgeTraceSchema>;
export interface BridgeSubscriptionOptions {
  /** Explicit opt-in: payloads may contain cells, prompts and responses. */
  readonly includeContent?: boolean;
  readonly onError?: (error: Error) => void;
}
/** Subscribe locally without consuming chat. Listener failures cannot break the bridge. */
export function subscribeBridge(
  path: string,
  listener: (event: BridgeTrace) => void,
  options: BridgeSubscriptionOptions = {},
): () => void {
  const socket = connect(path);
  const reader = createInterface({ input: socket });
  socket.once("connect", () =>
    socket.write(
      `${JSON.stringify({ command: "watch", includeContent: options.includeContent === true })}\n`,
    ),
  );
  reader.on("error", () => {});
  socket.on("error", () => options.onError?.(new Error("Trace bridge unavailable.")));
  reader.on("line", (line) => {
    try {
      const event = bridgeTraceSchema.parse(JSON.parse(line));
      listener(event);
    } catch {
      /* Isolate invalid events and consumer exceptions. */
    }
  });
  return () => {
    reader.close();
    socket.destroy();
  };
}
