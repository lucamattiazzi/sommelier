import type { JsonSocket } from "./index.js";
import { createSecureChannel, type SecureChannel } from "./secure.js";

/** Wrap an existing socket. Construct it with the fragment removed from the URL. */
export function createEncryptedSocket(socket: JsonSocket, pairUrl: string): JsonSocket {
  const url = new URL(pairUrl);
  const secret = url.hash.slice(1);
  const role = url.searchParams.get("role");
  const context = url.searchParams.get("uid") ?? url.searchParams.get("session");
  if (!secret || !context || (role !== "addin" && role !== "agent"))
    throw new Error("Invalid encrypted Sommelier URL.");
  const events = new EventTarget();
  const peer = role === "addin" ? "agent" : "addin";
  let channel: SecureChannel | undefined;
  let incoming = Promise.resolve();
  let outgoing = Promise.resolve();
  let handshakeTimer: ReturnType<typeof setTimeout> | undefined;
  const deliver = (data: string) => events.dispatchEvent(new MessageEvent("message", { data }));
  const presence = (connected: boolean) =>
    deliver(
      JSON.stringify({
        protocolVersion: "0.2",
        type: "event",
        event: "pair.peer.changed",
        data: { role: peer, connected },
      }),
    );
  const fail = () => {
    clearTimeout(handshakeTimer);
    channel = undefined;
    socket.close(1008, "Encrypted channel authentication failed.");
  };
  for (const type of ["open", "error", "close"]) {
    socket.addEventListener(type, (event) => {
      if (type !== "open") {
        clearTimeout(handshakeTimer);
        channel = undefined;
      }
      const forwarded =
        type === "close"
          ? Object.assign(new Event(type), {
              code: "code" in event ? event.code : 1006,
              reason: "reason" in event ? event.reason : "Connection closed",
            })
          : new Event(type);
      events.dispatchEvent(forwarded);
    });
  }
  socket.addEventListener("message", (event) => {
    if (!("data" in event) || typeof event.data !== "string") {
      fail();
      return;
    }
    const data = event.data;
    incoming = incoming
      .then(async () => {
        const message: unknown = JSON.parse(data);
        if (
          message &&
          typeof message === "object" &&
          "event" in message &&
          message.event === "pair.peer.changed"
        ) {
          const info = "data" in message ? message.data : undefined;
          if (
            !info ||
            typeof info !== "object" ||
            !("role" in info) ||
            info.role !== peer ||
            !("connected" in info) ||
            typeof info.connected !== "boolean"
          )
            throw new Error("Invalid presence event.");
          clearTimeout(handshakeTimer);
          channel = undefined;
          presence(false);
          if (info.connected) {
            const current = createSecureChannel({
              secret,
              role,
              context,
              send: (frame) => {
                if (channel === current) socket.send(frame);
              },
              receive: (plaintext) => {
                if (channel === current) deliver(plaintext);
              },
              onAuthenticated: () => {
                if (channel === current) {
                  clearTimeout(handshakeTimer);
                  presence(true);
                }
              },
            });
            channel = current;
            handshakeTimer = setTimeout(fail, 15_000);
            await current.start();
          }
          return;
        }
        if (!channel) throw new Error("No authenticated peer.");
        await channel.accept(data);
      })
      .catch(fail);
  });
  return {
    get readyState() {
      return socket.readyState;
    },
    send(data) {
      const current = channel;
      // No offline queue: workbook notifications cannot leak into a later connection.
      if (!current?.authenticated) return;
      outgoing = outgoing
        .then(async () => {
          if (channel === current) await current.send(data);
        })
        .catch(fail);
    },
    close(code, reason) {
      clearTimeout(handshakeTimer);
      channel = undefined;
      socket.close(code, reason);
    },
    addEventListener(type, callback) {
      events.addEventListener(type, callback);
    },
    removeEventListener(type, callback) {
      events.removeEventListener(type, callback);
    },
  };
}

/** WebSocket forbids fragments; the secret is consumed locally and never sent to the relay. */
export function relaySocketUrl(pairUrl: string): string {
  const url = new URL(pairUrl);
  url.hash = "";
  return url.toString();
}
