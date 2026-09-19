/** Encrypted transport framing is independent of workbook RPC and model providers. */
export type PairRole = "addin" | "agent";
export interface PairIdentity {
  readonly id: string;
  readonly secret: string;
}

const encoder = new TextEncoder();
const bytes = (value: string): Uint8Array<ArrayBuffer> => encoder.encode(value);
const encode = (value: ArrayBuffer | Uint8Array): string => {
  const data = value instanceof Uint8Array ? value : new Uint8Array(value);
  let binary = "";
  for (let offset = 0; offset < data.length; offset += 8192)
    binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "");
};
const decode = (value: string): Uint8Array<ArrayBuffer> => {
  if (!/^[A-Za-z0-9_-]+$/.test(value)) throw new Error("Invalid encrypted frame encoding.");
  return Uint8Array.from(atob(value.replaceAll("-", "+").replaceAll("_", "/")), (c) =>
    c.charCodeAt(0),
  );
};

export function createPairIdentity(): PairIdentity {
  return { id: crypto.randomUUID(), secret: encode(crypto.getRandomValues(new Uint8Array(32))) };
}

async function hmac(secret: string, value: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    decode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  return encode(await crypto.subtle.sign("HMAC", key, bytes(value)));
}

export async function pairConnectionUrl(
  origin: string,
  identity: PairIdentity,
  role: PairRole,
): Promise<string> {
  if (decode(identity.secret).length !== 32)
    throw new Error("Sommelier secrets must contain 32 random bytes.");
  const url = new URL("/connect", origin);
  const local = ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname);
  if (
    !["https:", "wss:"].includes(url.protocol) &&
    !(local && ["http:", "ws:"].includes(url.protocol))
  ) {
    throw new Error("Use HTTPS/WSS, or a loopback development server.");
  }
  url.protocol = ["https:", "wss:"].includes(url.protocol) ? "wss:" : "ws:";
  url.searchParams.set("uid", identity.id);
  url.searchParams.set("role", role);
  // The relay capability cannot be used to derive the end-to-end secret.
  url.searchParams.set("token", await hmac(identity.secret, `ai-cdl-relay-v1:${identity.id}`));
  url.hash = identity.secret;
  return url.toString();
}

export type SecureFrame =
  | { readonly v: 1; readonly type: "hello"; readonly nonce: string }
  | { readonly v: 1; readonly type: "auth"; readonly proof: string }
  | { readonly v: 1; readonly type: "data"; readonly seq: number; readonly ciphertext: string };

/** Validate only public framing; the relay never deserializes application messages. */
export function parseSecureFrame(input: string): SecureFrame {
  if (input.length > 1_400_000) throw new Error("Encrypted frame is too large.");
  const frame: unknown = JSON.parse(input);
  if (!frame || typeof frame !== "object") throw new Error("Invalid encrypted frame.");
  const f = frame as Record<string, unknown>;
  if (f.v === 1) {
    if (f.type === "hello" && typeof f.nonce === "string" && /^[\w-]{43}$/.test(f.nonce))
      return { v: 1, type: "hello", nonce: f.nonce };
    if (f.type === "auth" && typeof f.proof === "string" && /^[\w-]{43}$/.test(f.proof))
      return { v: 1, type: "auth", proof: f.proof };
    if (
      f.type === "data" &&
      typeof f.seq === "number" &&
      Number.isSafeInteger(f.seq) &&
      f.seq >= 0 &&
      typeof f.ciphertext === "string" &&
      /^[\w-]{22,1400000}$/.test(f.ciphertext)
    )
      return { v: 1, type: "data", seq: f.seq, ciphertext: f.ciphertext };
  }
  throw new Error("Only encrypted transport frames are accepted.");
}

export interface SecureChannelOptions {
  readonly secret: string;
  readonly role: PairRole;
  readonly context: string;
  readonly send: (frame: string) => void;
  readonly receive: (plaintext: string) => void;
  readonly onAuthenticated?: () => void;
}
export interface SecureChannel {
  readonly authenticated: boolean;
  start(): Promise<void>;
  accept(frame: string): Promise<void>;
  send(plaintext: string): Promise<void>;
}

/** PSK authentication, fresh challenges, directional AES-GCM keys and ordered replay protection. */
export function createSecureChannel(options: SecureChannelOptions): SecureChannel {
  if (decode(options.secret).length !== 32) throw new Error("Invalid Sommelier secret.");
  const peer: PairRole = options.role === "addin" ? "agent" : "addin";
  let nonce = "";
  let remoteNonce = "";
  let authenticated = false;
  let outgoing = 0;
  let incoming = 0;
  let sendKey: CryptoKey | undefined;
  let receiveKey: CryptoKey | undefined;
  const transcript = () =>
    JSON.stringify([
      "ai-cdl-e2ee-v1",
      options.context,
      options.role === "addin" ? nonce : remoteNonce,
      options.role === "agent" ? nonce : remoteNonce,
    ]);
  const emit = (frame: SecureFrame) => options.send(JSON.stringify(frame));
  const iv = (seq: number) => {
    const buffer = new Uint8Array(12);
    new DataView(buffer.buffer).setBigUint64(4, BigInt(seq));
    return buffer;
  };
  const derive = async (role: PairRole) => {
    const material = await crypto.subtle.importKey("raw", decode(options.secret), "HKDF", false, [
      "deriveKey",
    ]);
    return crypto.subtle.deriveKey(
      {
        name: "HKDF",
        hash: "SHA-256",
        salt: bytes(transcript()),
        info: bytes(`ai-cdl-data:${role}`),
      },
      material,
      { name: "AES-GCM", length: 256 },
      false,
      ["encrypt", "decrypt"],
    );
  };
  return {
    get authenticated() {
      return authenticated;
    },
    async start() {
      if (nonce) throw new Error("Create a fresh channel for each connection.");
      nonce = encode(crypto.getRandomValues(new Uint8Array(32)));
      emit({ v: 1, type: "hello", nonce });
    },
    async accept(input) {
      const frame = parseSecureFrame(input);
      if (!nonce) throw new Error("Channel has not started.");
      if (frame.type === "hello") {
        if (remoteNonce) throw new Error("Unexpected repeated handshake.");
        remoteNonce = frame.nonce;
        emit({
          v: 1,
          type: "auth",
          proof: await hmac(options.secret, `${transcript()}:${options.role}`),
        });
        return;
      }
      if (frame.type === "auth") {
        if (!remoteNonce || authenticated) throw new Error("Unexpected authentication.");
        const key = await crypto.subtle.importKey(
          "raw",
          decode(options.secret),
          { name: "HMAC", hash: "SHA-256" },
          false,
          ["verify"],
        );
        if (
          !(await crypto.subtle.verify(
            "HMAC",
            key,
            decode(frame.proof),
            bytes(`${transcript()}:${peer}`),
          ))
        )
          throw new Error("Terminal authentication failed.");
        [sendKey, receiveKey] = await Promise.all([derive(options.role), derive(peer)]);
        authenticated = true;
        options.onAuthenticated?.();
        return;
      }
      if (!authenticated || !receiveKey || frame.seq !== incoming)
        throw new Error("Unauthenticated or replayed frame.");
      const plaintext = await crypto.subtle.decrypt(
        { name: "AES-GCM", iv: iv(frame.seq), additionalData: bytes(`${transcript()}:${peer}`) },
        receiveKey,
        decode(frame.ciphertext),
      );
      incoming += 1;
      options.receive(new TextDecoder("utf-8", { fatal: true }).decode(plaintext));
    },
    async send(plaintext) {
      if (!authenticated || !sendKey) throw new Error("Terminal is not authenticated.");
      if (bytes(plaintext).length > 1_000_000 || !Number.isSafeInteger(outgoing + 1))
        throw new Error("Message limit exceeded.");
      const seq = outgoing++;
      emit({
        v: 1,
        type: "data",
        seq,
        ciphertext: encode(
          await crypto.subtle.encrypt(
            {
              name: "AES-GCM",
              iv: iv(seq),
              additionalData: bytes(`${transcript()}:${options.role}`),
            },
            sendKey,
            bytes(plaintext),
          ),
        ),
      });
    },
  };
}
