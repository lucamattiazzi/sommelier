import type { AddinController } from "@lucamattiazzi/sommelier-addin-core";
import {
  type PairChatMessageEvent,
  PROTOCOL_VERSION,
  type ProtocolError,
  type ProtocolEvent,
  type ProtocolMethod,
  type ProtocolMethodMap,
  type ProtocolRequest,
} from "@lucamattiazzi/sommelier-protocol";
import type { Transport, TransportUnsubscribe } from "@lucamattiazzi/sommelier-transport";

export interface PairLifecycle {
  start(): Promise<void>;
  stop(): Promise<void>;
}

/** Forward protocol messages between two transports without interpreting them. */
export function createPairRelay(left: Transport, right: Transport): PairLifecycle {
  let unsubscribes: TransportUnsubscribe[] = [];
  return {
    async start() {
      if (unsubscribes.length > 0) return;
      await Promise.all([left.connect(), right.connect()]);
      unsubscribes = [
        left.subscribe((message) => right.send(message)),
        right.subscribe((message) => left.send(message)),
      ];
    },
    async stop() {
      for (const unsubscribe of unsubscribes) unsubscribe();
      unsubscribes = [];
      await Promise.all([left.close(), right.close()]);
    },
  };
}

export type PairChatListener = (message: PairChatMessageEvent) => void;

/** Build a validated chat event for one Sommelier role. */
function chatMessage(
  sender: PairChatMessageEvent["data"]["sender"],
  content: string,
  createId: () => string,
  now: () => Date,
): PairChatMessageEvent {
  const text = content.trim();
  if (!text) throw new Error("Chat content is required.");
  return {
    protocolVersion: PROTOCOL_VERSION,
    type: "event",
    event: "pair.chat.message",
    data: { messageId: createId(), sender, content: text, occurredAt: now().toISOString() },
  };
}

export interface PairAddinSessionOptions {
  readonly transport: Transport;
  readonly controller: AddinController;
  readonly createId?: () => string;
  readonly now?: () => Date;
}

export interface PairAddinSession extends PairLifecycle {
  /** Send a task-pane message to the paired coding agent. */
  sendUserMessage(content: string): PairChatMessageEvent;
  /** Observe agent replies delivered over the Sommelier session. */
  subscribeChat(listener: PairChatListener): () => void;
}

/** Bind a protocol transport to the host-agnostic add-in controller. */
export function createPairAddinSession(options: PairAddinSessionOptions): PairAddinSession {
  const createId = options.createId ?? (() => crypto.randomUUID());
  const now = options.now ?? (() => new Date());
  const chatListeners = new Set<PairChatListener>();
  let unsubscribeTransport: TransportUnsubscribe | undefined;
  let unsubscribeController: (() => void) | undefined;
  let active = false;
  let generation = 0;
  return {
    async start() {
      if (unsubscribeTransport) return;
      const attempt = ++generation;
      active = true;
      unsubscribeTransport = options.transport.subscribe((message) => {
        if (message.type === "event") {
          if (message.event === "pair.chat.message" && message.data.sender === "agent") {
            for (const listener of [...chatListeners]) listener(message);
          }
          return;
        }
        if (message.type !== "request" || !active) return;
        void options.controller
          .handle(message)
          .then((response) => {
            if (active && generation === attempt && options.transport.state === "connected")
              options.transport.send(response);
          })
          .catch(() => {
            /* A response cannot be delivered after disconnect. */
          });
      });
      try {
        await options.transport.connect();
        await options.controller.start();
        if (!active || generation !== attempt) {
          await options.controller.stop();
          await options.transport.close();
          return;
        }
        unsubscribeController = options.controller.subscribe((event) => {
          if (active && options.transport.state === "connected") options.transport.send(event);
        });
      } catch (error) {
        active = false;
        unsubscribeTransport?.();
        unsubscribeTransport = undefined;
        await options.controller.stop();
        await options.transport.close();
        throw error;
      }
    },
    sendUserMessage(content) {
      const message = chatMessage("user", content, createId, now);
      options.transport.send(message);
      return message;
    },
    subscribeChat(listener) {
      chatListeners.add(listener);
      return () => chatListeners.delete(listener);
    },
    async stop() {
      active = false;
      generation += 1;
      unsubscribeTransport?.();
      unsubscribeController?.();
      unsubscribeTransport = undefined;
      unsubscribeController = undefined;
      chatListeners.clear();
      try {
        await options.controller.stop();
      } finally {
        await options.transport.close();
      }
    },
  };
}

export class PairProtocolError extends Error {
  readonly code: string;
  readonly data: ProtocolError["error"]["data"];

  constructor(error: ProtocolError["error"]) {
    super(error.message);
    this.name = "PairProtocolError";
    this.code = error.code;
    this.data = error.data;
  }
}

export interface PairClientOptions {
  readonly transport: Transport;
  readonly createId?: () => string;
  readonly now?: () => Date;
  readonly timeoutMs?: number;
}

export interface PairClient {
  connect(): Promise<void>;
  request<Method extends ProtocolMethod>(
    method: Method,
    params: ProtocolMethodMap[Method]["params"],
  ): Promise<ProtocolMethodMap[Method]["result"]>;
  subscribe(listener: (event: ProtocolEvent) => void): () => void;
  /** Send an assistant reply to the paired task pane. */
  sendAgentMessage(content: string): PairChatMessageEvent;
  /** Observe task-pane messages typed by the user. */
  subscribeChat(listener: PairChatListener): () => void;
  close(): Promise<void>;
}

interface PendingRequest {
  readonly method: ProtocolMethod;
  readonly resolve: (result: unknown) => void;
  readonly reject: (error: Error) => void;
  readonly timer: ReturnType<typeof setTimeout>;
}

/** Typed RPC client used by local Sommelier agents. */
export function createPairClient(options: PairClientOptions): PairClient {
  const createId = options.createId ?? (() => crypto.randomUUID());
  const now = options.now ?? (() => new Date());
  const timeoutMs = options.timeoutMs ?? 30_000;
  const pending = new Map<string, PendingRequest>();
  const listeners = new Set<(event: ProtocolEvent) => void>();
  const chatListeners = new Set<PairChatListener>();
  let unsubscribe: TransportUnsubscribe | undefined;
  let unsubscribeState: TransportUnsubscribe | undefined;
  function rejectPending(reason: string): void {
    for (const request of pending.values()) {
      clearTimeout(request.timer);
      request.reject(new Error(reason));
    }
    pending.clear();
  }

  function receive(message: Parameters<Transport["send"]>[0]): void {
    if (message.type === "event") {
      for (const listener of [...listeners]) listener(message);
      if (message.event === "pair.chat.message" && message.data.sender === "user") {
        for (const listener of [...chatListeners]) listener(message);
      }
      return;
    }
    if (message.type !== "response" && message.type !== "error") return;
    if (!message.id) return;
    const request = pending.get(message.id);
    if (!request) return;
    pending.delete(message.id);
    clearTimeout(request.timer);
    if (message.type === "error") request.reject(new PairProtocolError(message.error));
    else if (message.method !== request.method)
      request.reject(new Error("Sommelier response method does not match the request."));
    else request.resolve(message.result);
  }

  return {
    async connect() {
      if (unsubscribe) return;
      unsubscribe = options.transport.subscribe(receive);
      unsubscribeState = options.transport.subscribeState((state) => {
        if (state === "disconnected")
          rejectPending(
            "Sommelier transport disconnected. Check the workbook before retrying a mutation.",
          );
      });
      try {
        await options.transport.connect();
      } catch (error) {
        unsubscribe();
        unsubscribeState();
        unsubscribe = undefined;
        unsubscribeState = undefined;
        throw error;
      }
    },
    request(method, params) {
      if (options.transport.state !== "connected") {
        return Promise.reject(new Error("Sommelier client is not connected."));
      }
      const id = createId();
      if (pending.has(id)) return Promise.reject(new Error("Duplicate pending request ID."));
      const message = {
        protocolVersion: PROTOCOL_VERSION,
        type: "request",
        id,
        method,
        params,
      } as ProtocolRequest;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          reject(new Error(`Sommelier request timed out: ${id}.`));
        }, timeoutMs);
        pending.set(id, {
          method,
          resolve: (result) => resolve(result as ProtocolMethodMap[typeof method]["result"]),
          reject,
          timer,
        });
        try {
          options.transport.send(message);
        } catch (error) {
          clearTimeout(timer);
          pending.delete(id);
          reject(error instanceof Error ? error : new Error(String(error)));
        }
      });
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    sendAgentMessage(content) {
      const message = chatMessage("agent", content, createId, now);
      options.transport.send(message);
      return message;
    },
    subscribeChat(listener) {
      chatListeners.add(listener);
      return () => chatListeners.delete(listener);
    },
    async close() {
      unsubscribe?.();
      unsubscribe = undefined;
      unsubscribeState?.();
      unsubscribeState = undefined;
      rejectPending("Sommelier client closed.");
      chatListeners.clear();
      await options.transport.close();
    },
  };
}
