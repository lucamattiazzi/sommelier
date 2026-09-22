import { randomBytes, randomUUID, timingSafeEqual } from "node:crypto";
import { readFile, stat } from "node:fs/promises";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { extname, resolve, sep } from "node:path";
import { PROTOCOL_VERSION, parseProtocolMessage } from "@lucamattiazzi/sommelier-protocol";
import { type RawData, WebSocket, WebSocketServer } from "ws";
import { createTerminalRegistry } from "./registry.js";

export interface PairServerOptions {
  /** For migration tests and trusted loopback development only. Disabled by default. */
  readonly allowUnencryptedLegacy?: boolean;
  readonly host?: string;
  readonly port?: number;
  readonly publicOrigin: string;
  /** Independently hosted, trusted pane origin; defaults to publicOrigin. */
  readonly addinOrigin?: string;
  readonly agentOrigin?: string;
  readonly staticDirectory?: string;
  readonly bridgeScript?: string;
  readonly sessionTtlMs?: number;
  readonly maxActiveSessions?: number;
  readonly maxPayloadBytes?: number;
  readonly createId?: () => string;
  readonly createToken?: () => string;
  readonly now?: () => Date;
}

export interface PairServer {
  readonly host: string;
  readonly port: number;
  close(): Promise<void>;
}

type PairRole = "addin" | "agent";

interface PairSession {
  readonly id: string;
  readonly addinToken: string;
  readonly agentToken: string;
  readonly expiresAt: number;
  addin?: WebSocket;
  agent?: WebSocket;
}

const contentTypes: Readonly<Record<string, string>> = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".xml": "application/xml; charset=utf-8",
};

function secureHeaders(response: ServerResponse): void {
  response.setHeader("x-content-type-options", "nosniff");
  response.setHeader("referrer-policy", "no-referrer");
  response.setHeader("permissions-policy", "camera=(), microphone=(), geolocation=()");
  response.setHeader(
    "content-security-policy",
    "default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self' https: wss:; frame-ancestors https://*.office.com https://*.officeapps.live.com https://*.cloud.microsoft https://*.microsoft365.com https://onedrive.live.com https://*.sharepoint.com",
  );
}

function json(response: ServerResponse, status: number, value: unknown): void {
  secureHeaders(response);
  response.setHeader("cache-control", "no-store");
  response.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(value));
}

function httpOrigin(socketUrl: URL): string {
  const result = new URL(socketUrl);
  result.protocol = result.protocol === "wss:" ? "https:" : "http:";
  return result.origin;
}

function socketOrigin(publicOrigin: URL): string {
  const result = new URL(publicOrigin);
  result.protocol = result.protocol === "https:" ? "wss:" : "ws:";
  return result.origin;
}

function sameToken(actual: string, expected: string): boolean {
  const actualBytes = Buffer.from(actual);
  const expectedBytes = Buffer.from(expected);
  return actualBytes.length === expectedBytes.length && timingSafeEqual(actualBytes, expectedBytes);
}

function listen(
  server: ReturnType<typeof createServer>,
  port: number,
  host: string,
): Promise<void> {
  return new Promise((resolvePromise, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => {
      server.off("error", reject);
      resolvePromise();
    });
  });
}

async function serveStatic(
  request: IncomingMessage,
  response: ServerResponse,
  directory: string,
): Promise<boolean> {
  if (request.method !== "GET" && request.method !== "HEAD") return false;
  const pathname = decodeURIComponent(new URL(request.url ?? "/", "http://localhost").pathname);
  const relative = pathname === "/" ? "index.html" : pathname.slice(1);
  let file = resolve(directory, relative);
  const root = resolve(directory);
  if (file !== root && !file.startsWith(`${root}${sep}`)) return false;
  try {
    if (!(await stat(file)).isFile()) return false;
  } catch {
    if (extname(relative)) return false;
    file = resolve(root, "index.html");
  }
  try {
    const contents = await readFile(file);
    secureHeaders(response);
    response.writeHead(200, {
      "cache-control": /[-.][A-Za-z0-9_-]{8,}\.(js|css)$/.test(file)
        ? "public, max-age=31536000, immutable"
        : "no-cache",
      "content-type": contentTypes[extname(file)] ?? "application/octet-stream",
    });
    response.end(request.method === "HEAD" ? undefined : contents);
    return true;
  } catch {
    return false;
  }
}

/** Create the production Sommelier static host and authenticated WebSocket rendezvous relay. */
export async function createPairServer(options: PairServerOptions): Promise<PairServer> {
  const host = options.host ?? "127.0.0.1";
  const publicOrigin = new URL(options.publicOrigin);
  if (!new Set(["http:", "https:"]).has(publicOrigin.protocol)) {
    throw new Error("PAIR_PUBLIC_ORIGIN must use HTTP or HTTPS.");
  }
  const now = options.now ?? (() => new Date());
  const createId = options.createId ?? randomUUID;
  const createToken = options.createToken ?? (() => randomBytes(32).toString("base64url"));
  const sessionTtlMs = options.sessionTtlMs ?? 10 * 60_000;
  const maxActiveSessions = options.maxActiveSessions ?? 1_000;
  const relayOrigin = socketOrigin(publicOrigin);
  const agentOrigin = options.agentOrigin ? new URL(options.agentOrigin) : new URL(relayOrigin);
  if (!new Set(["ws:", "wss:"]).has(agentOrigin.protocol)) {
    throw new Error("PAIR_AGENT_ORIGIN must use WS or WSS.");
  }
  const bridgeScript = resolve(options.bridgeScript ?? "skills/sommelier/scripts/session.mjs");
  const sessions = new Map<string, PairSession>();
  const addinOrigin = new URL(options.addinOrigin ?? publicOrigin.origin).origin;
  const registry = createTerminalRegistry(addinOrigin, maxActiveSessions, sessionTtlMs);
  const sockets = new WebSocketServer({
    noServer: true,
    maxPayload: options.maxPayloadBytes ?? 1_400_000,
  });

  function attach(socket: WebSocket, session: PairSession, role: PairRole): void {
    const existing = session[role];
    if (existing?.readyState === WebSocket.OPEN) {
      socket.close(1008, `The ${role} role is already connected.`);
      return;
    }
    session[role] = socket;
    const peerRole: PairRole = role === "addin" ? "agent" : "addin";
    const peer = session[peerRole];
    if (peer?.readyState === WebSocket.OPEN) {
      peer.send(
        JSON.stringify({
          protocolVersion: PROTOCOL_VERSION,
          type: "event",
          event: "pair.peer.changed",
          data: { role, connected: true },
        }),
      );
      socket.send(
        JSON.stringify({
          protocolVersion: PROTOCOL_VERSION,
          type: "event",
          event: "pair.peer.changed",
          data: { role: peerRole, connected: true },
        }),
      );
    }
    socket.on("message", (data: RawData, isBinary: boolean) => {
      if (isBinary) {
        socket.close(1003, "Binary messages are unsupported.");
        return;
      }
      try {
        const message = parseProtocolMessage(JSON.parse(data.toString()));
        if (message.type === "event" && message.event === "pair.peer.changed") {
          throw new Error("Sommelier presence events are server-owned.");
        }
        const peer = session[peerRole];
        if (peer?.readyState === WebSocket.OPEN) peer.send(JSON.stringify(message));
      } catch {
        socket.close(1007, "Invalid protocol message.");
      }
    });
    socket.on("close", () => {
      if (session[role] !== socket) return;
      delete session[role];
      const peer = session[peerRole];
      if (peer?.readyState === WebSocket.OPEN) {
        peer.send(
          JSON.stringify({
            protocolVersion: PROTOCOL_VERSION,
            type: "event",
            event: "pair.peer.changed",
            data: { role, connected: false },
          }),
        );
      }
    });
  }

  const server = createServer((request, response) => {
    void (async () => {
      const url = new URL(request.url ?? "/", publicOrigin);
      if (request.method === "GET" && url.pathname === "/healthz") {
        json(response, 200, { status: "ok" });
        return;
      }
      if (
        request.method === "GET" &&
        [
          "/agent/session.mjs",
          "/agent/lib/encrypted-socket.mjs",
          "/agent/lib/excel-docs.mjs",
        ].includes(url.pathname)
      ) {
        try {
          const contents = await readFile(
            url.pathname === "/agent/session.mjs"
              ? bridgeScript
              : resolve(bridgeScript, "..", url.pathname.slice("/agent/".length)),
          );
          secureHeaders(response);
          response.writeHead(200, {
            "cache-control": "no-cache",
            "content-type": "text/javascript; charset=utf-8",
          });
          response.end(contents);
        } catch {
          json(response, 404, { error: "The agent bridge script is unavailable." });
        }
        return;
      }
      if (request.method === "GET" && url.pathname === "/api/config") {
        response.setHeader("vary", "Origin");
        if (request.headers.origin === addinOrigin)
          response.setHeader("access-control-allow-origin", addinOrigin);
        json(response, 200, {
          protocolVersion: "0.2",
          relayPath: "/connect",
          encryption: "ai-cdl-e2ee-v1",
          agentOrigin: agentOrigin.origin,
        });
        return;
      }
      if (request.method === "POST" && url.pathname === "/api/pair/sessions") {
        if (!options.allowUnencryptedLegacy) {
          json(response, 410, { error: "Use encrypted terminal pairing at /connect." });
          return;
        }
        for (const [id, session] of sessions) {
          if (session.expiresAt <= now().getTime()) sessions.delete(id);
        }
        if (sessions.size >= maxActiveSessions) {
          json(response, 503, { error: "Sommelier relay capacity reached. Try again later." });
          return;
        }
        const session: PairSession = {
          id: createId(),
          addinToken: createToken(),
          agentToken: createToken(),
          expiresAt: now().getTime() + sessionTtlMs,
        };
        sessions.set(session.id, session);
        const connectionUrl = (origin: string, role: PairRole, token: string) => {
          const target = new URL("/pair", origin);
          target.searchParams.set("session", session.id);
          target.searchParams.set("role", role);
          target.searchParams.set("token", token);
          return target.toString();
        };
        json(response, 201, {
          sessionId: session.id,
          expiresAt: new Date(session.expiresAt).toISOString(),
          addinUrl: connectionUrl(relayOrigin, "addin", session.addinToken),
          agentUrl: connectionUrl(agentOrigin.origin, "agent", session.agentToken),
          bridgeUrl: new URL("/agent/session.mjs", httpOrigin(agentOrigin)).toString(),
        });
        return;
      }
      if (
        options.staticDirectory &&
        (await serveStatic(request, response, options.staticDirectory))
      ) {
        return;
      }
      json(response, 404, { error: "Not found" });
    })().catch(() => json(response, 500, { error: "Internal server error" }));
  });

  server.on("upgrade", (request, socket, head) => {
    const url = new URL(request.url ?? "/", publicOrigin);
    if (url.pathname === "/connect") {
      sockets.handleUpgrade(request, socket, head, (webSocket) =>
        registry.attach(webSocket, request, url),
      );
      return;
    }
    const sessionId = url.searchParams.get("session");
    const role = url.searchParams.get("role");
    const token = url.searchParams.get("token");
    const session = sessionId ? sessions.get(sessionId) : undefined;
    const validRole = role === "addin" || role === "agent";
    const expectedToken = validRole && session ? session[`${role}Token`] : undefined;
    if (
      !options.allowUnencryptedLegacy ||
      url.pathname !== "/pair" ||
      !session ||
      !validRole ||
      !token ||
      !expectedToken ||
      session.expiresAt <= now().getTime() ||
      !sameToken(token, expectedToken)
    ) {
      sockets.handleUpgrade(request, socket, head, (webSocket) =>
        webSocket.close(1008, "Invalid or expired pairing credentials."),
      );
      return;
    }
    sockets.handleUpgrade(request, socket, head, (webSocket) => {
      attach(webSocket, session, role);
    });
  });

  await listen(server, options.port ?? 3000, host);
  const address = server.address();
  if (!address || typeof address === "string")
    throw new Error("Sommelier server address is unavailable.");

  return {
    host,
    port: address.port,
    async close() {
      registry.close();
      for (const socket of sockets.clients) socket.terminate();
      await new Promise<void>((resolvePromise, reject) =>
        sockets.close((error) => (error ? reject(error) : resolvePromise())),
      );
      await new Promise<void>((resolvePromise, reject) =>
        server.close((error) => (error ? reject(error) : resolvePromise())),
      );
    },
  };
}
