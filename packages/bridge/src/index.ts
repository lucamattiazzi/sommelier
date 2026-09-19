import { createServer, type Server } from "node:http";
import { parseProtocolMessage } from "@lucamattiazzi/sommelier-protocol";
import { WebSocket, WebSocketServer } from "ws";

export type PairRole = "addin" | "agent";

export interface PairRelayServerOptions {
  readonly host?: string;
  readonly port?: number;
}

export interface PairRelayServer {
  readonly host: string;
  readonly port: number;
  close(): Promise<void>;
}

interface PairPeers {
  addin?: WebSocket;
  agent?: WebSocket;
}

function listen(server: Server, port: number, host: string): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => {
      server.off("error", reject);
      resolve();
    });
  });
}

/** Start a loopback WebSocket relay that only pairs and forwards validated envelopes. */
export async function createPairRelayServer(
  options: PairRelayServerOptions = {},
): Promise<PairRelayServer> {
  const host = options.host ?? "127.0.0.1";
  const server = createServer();
  const sockets = new WebSocketServer({ server, path: "/pair" });
  const sessions = new Map<string, PairPeers>();

  sockets.on("connection", (socket, request) => {
    const url = new URL(request.url ?? "/pair", `http://${request.headers.host ?? host}`);
    const sessionId = url.searchParams.get("session");
    const role = url.searchParams.get("role");
    if (!sessionId || (role !== "addin" && role !== "agent")) {
      socket.close(1008, "A session and valid role are required.");
      return;
    }
    const peers = sessions.get(sessionId) ?? {};
    if (peers[role]?.readyState === WebSocket.OPEN) {
      socket.close(1008, `The ${role} role is already connected.`);
      return;
    }
    peers[role] = socket;
    sessions.set(sessionId, peers);

    socket.on("message", (data, isBinary) => {
      if (isBinary) return;
      try {
        const message = parseProtocolMessage(JSON.parse(data.toString()));
        const peer = role === "addin" ? peers.agent : peers.addin;
        if (peer?.readyState === WebSocket.OPEN) peer.send(JSON.stringify(message));
      } catch {
        socket.close(1007, "Invalid protocol message.");
      }
    });
    socket.on("close", () => {
      if (peers[role] === socket) delete peers[role];
      if (!peers.addin && !peers.agent) sessions.delete(sessionId);
    });
  });

  await listen(server, options.port ?? 4310, host);
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Relay address is unavailable.");

  return {
    host,
    port: address.port,
    async close() {
      for (const socket of sockets.clients) socket.terminate();
      await new Promise<void>((resolve, reject) =>
        sockets.close((error) => (error ? reject(error) : resolve())),
      );
      await new Promise<void>((resolve, reject) =>
        server.close((error) => (error ? reject(error) : resolve())),
      );
    },
  };
}

export * from "./bridge.js";
