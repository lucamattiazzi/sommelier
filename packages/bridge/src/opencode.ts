import { z } from "zod";
import { EXCEL_GUIDE } from "./mcp.js";

export interface OpenCodeAdapterOptions {
  readonly origin: string;
  readonly directory: string;
  readonly name: string;
  readonly command: readonly string[];
  readonly environment?: Record<string, string>;
  readonly password?: string;
  readonly username?: string;
  readonly fetch?: typeof fetch;
}
const sessionSchema = z.object({ id: z.string().min(1) });
export class OpenCodeAdapter {
  #session: string | undefined;
  readonly #origin: URL;
  constructor(readonly options: OpenCodeAdapterOptions) {
    this.#origin = new URL(options.origin);
    if (
      this.#origin.protocol !== "http:" ||
      !["127.0.0.1", "localhost", "[::1]"].includes(this.#origin.hostname) ||
      this.#origin.username ||
      this.#origin.password
    )
      throw new Error(
        "OpenCode must use a local loopback HTTP server; remote plaintext agent traffic is not allowed.",
      );
  }
  async #request(path: string, body?: unknown): Promise<unknown> {
    const url = new URL(path, this.#origin);
    url.searchParams.set("directory", this.options.directory);
    const response = await (this.options.fetch ?? fetch)(url, {
      method: body === undefined ? "GET" : "POST",
      redirect: "error",
      headers: {
        "Content-Type": "application/json",
        ...(this.options.password
          ? {
              Authorization: `Basic ${Buffer.from(`${this.options.username ?? "opencode"}:${this.options.password}`).toString("base64")}`,
            }
          : {}),
      },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      signal: AbortSignal.timeout(600_000),
    });
    if (!response.ok)
      throw new Error(
        `OpenCode request failed (${response.status}). Check the local server. The request was not retried.`,
      );
    return response.json();
  }
  async connect(sessionId?: string): Promise<string> {
    const key = `pair_${this.options.name}`;
    const statuses = z.record(z.string(), z.object({ status: z.string() })).parse(
      await this.#request("/mcp", {
        name: key,
        config: {
          type: "local",
          command: this.options.command,
          enabled: true,
          timeout: 125000,
          environment: this.options.environment ?? {},
        },
      }),
    );
    if (statuses[key]?.status !== "connected")
      throw new Error("OpenCode could not load the Excel MCP tools.");
    const session = sessionSchema.parse(
      await this.#request(
        sessionId ? `/session/${encodeURIComponent(sessionId)}` : "/session",
        sessionId ? undefined : { title: `Excel · ${this.options.name}` },
      ),
    );
    this.#session = session.id;
    return session.id;
  }
  async prompt(content: string): Promise<string> {
    if (!this.#session) throw new Error("OpenCode session is not connected.");
    const statuses = z
      .record(z.string(), z.object({ type: z.string() }))
      .parse(await this.#request("/session/status"));
    if (statuses[this.#session] && statuses[this.#session]?.type !== "idle")
      throw new Error(
        "OpenCode session is busy. Send the message again after its current turn finishes.",
      );
    const result = z
      .object({
        info: z.object({ error: z.unknown().optional() }).optional(),
        parts: z.array(z.object({ type: z.string(), text: z.string().optional() })),
      })
      .parse(
        await this.#request(`/session/${encodeURIComponent(this.#session)}/message`, {
          system: EXCEL_GUIDE,
          parts: [{ type: "text", text: content }],
        }),
      );
    if (result.info?.error) {
      const failure = z
        .object({
          name: z.string().regex(/^[A-Za-z0-9_]+$/),
          data: z.object({ statusCode: z.number().optional() }).optional(),
        })
        .safeParse(result.info.error);
      const reason = failure.success
        ? `${failure.data.name}${failure.data.data?.statusCode ? `, HTTP ${failure.data.data.statusCode}` : ""}`
        : "provider error";
      throw new Error(
        `OpenCode turn failed (${reason}). Check the local agent/provider configuration; no request was retried.`,
      );
    }
    return result.parts
      .filter((p) => p.type === "text")
      .map((p) => p.text ?? "")
      .join("\n");
  }
}
