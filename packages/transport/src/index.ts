import { type ProtocolMessage, parseProtocolMessage } from "@lucamattiazzi/sommelier-protocol";

/** Lifecycle state shared by all transports. */
export type TransportState = "disconnected" | "connecting" | "connected" | "closing";

export type TransportSubscriber = (message: ProtocolMessage) => void;
export type TransportStateSubscriber = (state: TransportState) => void;
export type TransportUnsubscribe = () => void;

/** Ordered, bidirectional protocol-message transport. */
export interface Transport {
  readonly state: TransportState;
  connect(): Promise<void>;
  send(message: ProtocolMessage): void;
  subscribe(subscriber: TransportSubscriber): TransportUnsubscribe;
  subscribeState(subscriber: TransportStateSubscriber): TransportUnsubscribe;
  close(): Promise<void>;
}

abstract class BaseTransport implements Transport {
  #state: TransportState = "disconnected";
  readonly #messageSubscribers = new Set<TransportSubscriber>();
  readonly #stateSubscribers = new Set<TransportStateSubscriber>();

  get state(): TransportState {
    return this.#state;
  }

  abstract connect(): Promise<void>;
  abstract send(message: ProtocolMessage): void;
  abstract close(): Promise<void>;

  subscribe(subscriber: TransportSubscriber): TransportUnsubscribe {
    this.#messageSubscribers.add(subscriber);
    return () => this.#messageSubscribers.delete(subscriber);
  }

  subscribeState(subscriber: TransportStateSubscriber): TransportUnsubscribe {
    this.#stateSubscribers.add(subscriber);
    return () => this.#stateSubscribers.delete(subscriber);
  }

  protected setState(state: TransportState): void {
    if (this.#state === state) return;
    this.#state = state;
    for (const subscriber of [...this.#stateSubscribers]) subscriber(state);
  }

  protected deliver(input: unknown): void {
    const message = parseProtocolMessage(input);
    for (const subscriber of [...this.#messageSubscribers]) subscriber(message);
  }
}

class InMemoryTransport extends BaseTransport {
  #peer: InMemoryTransport | undefined;

  setPeer(peer: InMemoryTransport): void {
    this.#peer = peer;
  }

  async connect(): Promise<void> {
    if (this.state === "connected") return;
    if (this.state !== "disconnected") {
      throw new Error(`Cannot connect while transport is ${this.state}.`);
    }
    this.setState("connecting");
    this.setState("connected");
  }

  send(message: ProtocolMessage): void {
    if (this.state !== "connected") throw new Error("Transport is not connected.");
    if (this.#peer?.state !== "connected") throw new Error("Peer transport is not connected.");
    this.#peer.deliverToPeer(message);
  }

  async close(): Promise<void> {
    if (this.state === "disconnected") return;
    this.setState("closing");
    this.setState("disconnected");
  }

  private deliverToPeer(input: unknown): void {
    this.deliver(input);
  }
}

/** Create two isolated in-memory endpoints that synchronously preserve send order. */
export function createInMemoryTransportPair(): readonly [Transport, Transport] {
  const left = new InMemoryTransport();
  const right = new InMemoryTransport();
  left.setPeer(right);
  right.setPeer(left);
  return [left, right];
}

/** Structural subset implemented by browser WebSocket and compatible Node WebSockets. */
export interface JsonSocket {
  readonly readyState: number;
  send(data: string): void;
  close(code?: number, reason?: string): void;
  addEventListener(type: string, callback: EventListenerOrEventListenerObject): void;
  removeEventListener(type: string, callback: EventListenerOrEventListenerObject): void;
}

const SOCKET_CONNECTING = 0;
const SOCKET_OPEN = 1;
const SOCKET_CLOSING = 2;
const SOCKET_CLOSED = 3;

class JsonSocketTransport extends BaseTransport {
  readonly #socket: JsonSocket;
  #connectPromise: Promise<void> | undefined;
  #resolveConnect: (() => void) | undefined;
  #rejectConnect: ((error: Error) => void) | undefined;
  #closePromise: Promise<void> | undefined;
  #resolveClose: (() => void) | undefined;

  constructor(socket: JsonSocket) {
    super();
    this.#socket = socket;
    socket.addEventListener("open", this.#handleOpen);
    socket.addEventListener("message", this.#handleMessage);
    socket.addEventListener("error", this.#handleError);
    socket.addEventListener("close", this.#handleClose);
  }

  connect(): Promise<void> {
    if (this.state === "connected") return Promise.resolve();
    if (this.state === "connecting" && this.#connectPromise) return this.#connectPromise;
    if (this.state !== "disconnected") {
      return Promise.reject(new Error(`Cannot connect while transport is ${this.state}.`));
    }

    this.setState("connecting");
    if (this.#socket.readyState === SOCKET_OPEN) {
      this.setState("connected");
      return Promise.resolve();
    }
    if (this.#socket.readyState !== SOCKET_CONNECTING) {
      this.setState("disconnected");
      return Promise.reject(new Error("JSON socket is not open and cannot connect."));
    }

    this.#connectPromise = new Promise<void>((resolve, reject) => {
      this.#resolveConnect = resolve;
      this.#rejectConnect = reject;
    });
    return this.#connectPromise;
  }

  send(message: ProtocolMessage): void {
    if (this.state !== "connected" || this.#socket.readyState !== SOCKET_OPEN) {
      throw new Error("Transport is not connected.");
    }
    this.#socket.send(JSON.stringify(parseProtocolMessage(message)));
  }

  close(): Promise<void> {
    if (this.state === "disconnected") return Promise.resolve();
    if (this.state === "closing" && this.#closePromise) return this.#closePromise;

    this.setState("closing");
    if (this.#socket.readyState === SOCKET_CLOSED) {
      this.#settleClosed();
      return Promise.resolve();
    }

    this.#closePromise = new Promise<void>((resolve) => {
      this.#resolveClose = resolve;
    });
    if (this.#socket.readyState !== SOCKET_CLOSING) this.#socket.close();
    return this.#closePromise;
  }

  readonly #handleOpen: EventListener = () => {
    if (this.state !== "connecting") return;
    this.setState("connected");
    this.#resolveConnect?.();
    this.#clearConnectPromise();
  };

  readonly #handleMessage: EventListener = (event) => {
    if (this.state !== "connected" || !("data" in event) || typeof event.data !== "string") return;
    try {
      this.deliver(JSON.parse(event.data));
    } catch {
      // Untrusted socket data is never delivered unless it is valid JSON and a protocol message.
    }
  };

  readonly #handleError: EventListener = () => {
    if (this.state === "connecting") {
      this.#rejectConnect?.(new Error("JSON socket connection failed."));
      this.#clearConnectPromise();
    }
    if (this.state !== "disconnected") this.#settleClosed();
  };

  readonly #handleClose: EventListener = () => {
    this.#settleClosed();
  };

  #settleClosed(): void {
    if (this.state === "connecting") {
      this.#rejectConnect?.(new Error("JSON socket closed before connecting."));
      this.#clearConnectPromise();
    }
    this.setState("disconnected");
    this.#resolveClose?.();
    this.#resolveClose = undefined;
    this.#closePromise = undefined;
  }

  #clearConnectPromise(): void {
    this.#connectPromise = undefined;
    this.#resolveConnect = undefined;
    this.#rejectConnect = undefined;
  }
}

/** Adapt an existing JSON WebSocket-like connection to the transport contract. */
export function createJsonSocketTransport(socket: JsonSocket): Transport {
  return new JsonSocketTransport(socket);
}

export { createEncryptedSocket, relaySocketUrl } from "./encrypted-socket.js";
export type {
  PairIdentity,
  PairRole,
  SecureChannel,
  SecureChannelOptions,
  SecureFrame,
} from "./secure.js";
export {
  createPairIdentity,
  createSecureChannel,
  pairConnectionUrl,
  parseSecureFrame,
} from "./secure.js";
