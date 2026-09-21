import { timingSafeEqual } from "node:crypto";
import type { IncomingMessage } from "node:http";
import { type PairRole, parseSecureFrame } from "@lucamattiazzi/sommelier-transport";
import { WebSocket } from "ws";

interface Connection {
  readonly token: string;
  touched: number;
  addin?: WebSocket;
  agent?: WebSocket;
}

/** In-memory rendezvous only. Names and trust records belong to endpoints, never this registry. */
export function createTerminalRegistry(publicOrigin: string, capacity: number, idleTtlMs: number) {
  const connections = new Map<string, Connection>();
  const seen = new WeakSet<WebSocket>();
  const same = (a: string, b: string) =>
    a.length === b.length && timingSafeEqual(Buffer.from(a), Buffer.from(b));
  const presence = (socket: WebSocket | undefined, role: PairRole, connected: boolean) => {
    if (socket?.readyState === WebSocket.OPEN)
      socket.send(
        JSON.stringify({
          protocolVersion: "0.2",
          type: "event",
          event: "pair.peer.changed",
          data: { role, connected },
        }),
      );
  };
  const sweep = () => {
    const now = Date.now();
    for (const [id, connection] of connections) {
      for (const role of ["addin", "agent"] as const) {
        const socket = connection[role];
        if (!socket) continue;
        if (seen.has(socket)) {
          seen.delete(socket);
          socket.ping();
        } else socket.terminate();
      }
      if (!connection.addin && !connection.agent && now - connection.touched > idleTtlMs)
        connections.delete(id);
    }
  };
  const timer = setInterval(sweep, 30_000);
  timer.unref();
  return {
    attach(socket: WebSocket, request: IncomingMessage, url: URL) {
      const id = url.searchParams.get("uid") ?? "";
      const token = url.searchParams.get("token") ?? "";
      const role = url.searchParams.get("role");
      const origin = request.headers.origin;
      if (!/^[a-f\d-]{36}$/i.test(id)) {
        socket.close(1008, "Invalid terminal UID.");
        return;
      }
      if (role !== "addin" && role !== "agent") {
        socket.close(1008, "Invalid terminal role.");
        return;
      }
      if (!/^[\w-]{43}$/.test(token)) {
        socket.close(1008, "Missing or invalid relay token. Copy the complete pairing URL.");
        return;
      }
      if (origin && origin !== publicOrigin) {
        socket.close(1008, "Origin is not allowed.");
        return;
      }
      let connection = connections.get(id);
      if (!connection) {
        if (connections.size >= capacity) {
          socket.close(1013, "Terminal registry is full.");
          return;
        }
        connection = { token, touched: Date.now() };
        connections.set(id, connection);
      }
      if (!same(connection.token, token)) {
        socket.close(1008, "Invalid terminal capability.");
        return;
      }
      if (connection[role] && connection[role]?.readyState !== WebSocket.CLOSED) {
        socket.close(1008, `The ${role} role is already connected.`);
        return;
      }
      const current = connection;
      const peer: PairRole = role === "addin" ? "agent" : "addin";
      current[role] = socket;
      current.touched = Date.now();
      seen.add(socket);
      socket.on("pong", () => seen.add(socket));
      if (current[peer]?.readyState === WebSocket.OPEN) {
        presence(current[peer], role, true);
        presence(socket, peer, true);
      }
      socket.on("message", (data, binary) => {
        if (binary) {
          socket.close(1003, "Text frames required.");
          return;
        }
        try {
          const frame = parseSecureFrame(data.toString());
          const target = current[peer];
          if (target?.readyState === WebSocket.OPEN) {
            if (target.bufferedAmount > 2_800_000) {
              socket.close(1013, "Peer cannot keep up.");
              return;
            }
            target.send(JSON.stringify(frame));
          }
        } catch {
          socket.close(1007, "Invalid encrypted frame.");
        }
      });
      socket.on("error", () => socket.terminate());
      socket.on("close", () => {
        if (current[role] !== socket) return;
        delete current[role];
        current.touched = Date.now();
        presence(current[peer], role, false);
      });
    },
    close() {
      clearInterval(timer);
      connections.clear();
    },
  };
}
