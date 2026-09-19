import {
  type AgentAdapter,
  type AgentEvent,
  type AgentRunOptions,
  type AgentTurnRequest,
  agentEventSchema,
  PROTOCOL_VERSION,
  SommelierError,
} from "@lucamattiazzi/sommelier-core";
import { z } from "zod";

/** HTTP adapter construction options. */
export interface HttpAgentAdapterOptions {
  readonly endpoint: string;
  readonly getAccessToken?: () => Promise<string | undefined>;
  readonly timeoutMs?: number;
  readonly headers?: Readonly<Record<string, string>>;
  readonly fetch?: typeof globalThis.fetch;
}

const responseSchema = z.object({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  events: z.array(agentEventSchema),
});

/** Stateless JSON transport for any agent implementing the documented Sommelier protocol. */
export class HttpAgentAdapter implements AgentAdapter {
  readonly #options: HttpAgentAdapterOptions;

  constructor(options: HttpAgentAdapterOptions) {
    const endpoint = new URL(options.endpoint);
    if (!["http:", "https:"].includes(endpoint.protocol)) {
      throw new SommelierError({
        code: "SOMMELIER_HTTP_ENDPOINT_INVALID",
        message: "Agent endpoint must use HTTP or HTTPS.",
      });
    }
    this.#options = options;
  }

  async *runTurn(request: AgentTurnRequest, options: AgentRunOptions): AsyncIterable<AgentEvent> {
    const timeoutMs = this.#options.timeoutMs ?? 30_000;
    const timeoutController = new AbortController();
    const timeout = setTimeout(
      () =>
        timeoutController.abort(
          new SommelierError({
            code: "SOMMELIER_HTTP_TIMEOUT",
            message: `Agent request timed out after ${timeoutMs}ms.`,
          }),
        ),
      timeoutMs,
    );
    const signal = AbortSignal.any([options.signal, timeoutController.signal]);
    try {
      const token = await this.#options.getAccessToken?.();
      if (signal.aborted) throw signal.reason;
      const response = await (this.#options.fetch ?? globalThis.fetch)(this.#options.endpoint, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json",
          "x-ai-cdl-protocol-version": PROTOCOL_VERSION,
          ...this.#options.headers,
          ...(token ? { authorization: `Bearer ${token}` } : {}),
        },
        body: JSON.stringify(request),
        signal,
      });
      if (!response.ok) {
        const requestId = response.headers.get("x-request-id");
        throw new SommelierError({
          code: "SOMMELIER_HTTP_STATUS_ERROR",
          message: `Agent endpoint returned HTTP ${response.status}.`,
          context: { status: response.status, ...(requestId ? { requestId } : {}) },
          probableCause:
            response.status === 401 || response.status === 403
              ? "The endpoint rejected the supplied credentials."
              : "The agent endpoint could not process the turn.",
          suggestedAction:
            "Check endpoint logs using the safe request ID; tokens and response bodies are not included in this error.",
        });
      }
      let payload: unknown;
      try {
        payload = await response.json();
      } catch (error) {
        throw new SommelierError({
          code: "SOMMELIER_HTTP_MALFORMED_RESPONSE",
          message: "Agent endpoint returned invalid JSON.",
          suggestedAction: "Return a JSON object containing protocolVersion and events.",
          cause: error,
        });
      }
      const parsed = responseSchema.safeParse(payload);
      if (!parsed.success) {
        throw new SommelierError({
          code: "SOMMELIER_HTTP_PROTOCOL_INVALID",
          message: "Agent endpoint response does not match protocol 0.1.",
          context: {
            issues: parsed.error.issues
              .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
              .join("; "),
          },
          suggestedAction: "Validate the endpoint with `sommelier-contract --endpoint <url>`.",
        });
      }
      for (const event of parsed.data.events) yield event;
    } catch (error) {
      if (signal.aborted) {
        const reason = signal.reason;
        if (reason instanceof Error) throw reason;
        throw new SommelierError({
          code: "SOMMELIER_HTTP_CANCELLED",
          message: "Agent request was cancelled.",
        });
      }
      if (error instanceof SommelierError) throw error;
      throw new SommelierError({
        code: "SOMMELIER_HTTP_NETWORK_ERROR",
        message: "Could not reach the agent endpoint.",
        probableCause: "The endpoint is unavailable, blocked by CORS, or has a TLS/network error.",
        suggestedAction: "Verify the URL, certificate, CORS policy, and network connectivity.",
        cause: error,
      });
    } finally {
      clearTimeout(timeout);
    }
  }
}

/** JSON Schema for the HTTP response envelope. */
export const httpAgentResponseJsonSchema = z.toJSONSchema(responseSchema, { target: "draft-7" });
