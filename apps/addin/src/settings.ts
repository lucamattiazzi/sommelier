export interface DirectAgentSettings {
  readonly endpoint: string;
  readonly token?: string;
}

/** Validate a direct BYOA endpoint without allowing insecure remote traffic. */
export function parseDirectAgentSettings(endpoint: string, token: string): DirectAgentSettings {
  const url = new URL(endpoint.trim());
  if (url.username || url.password || url.hash)
    throw new Error("Keep credentials out of the endpoint URL; use the bearer token field.");
  const local = url.hostname === "localhost" || url.hostname === "127.0.0.1";
  if (url.protocol !== "https:" && !(local && url.protocol === "http:")) {
    throw new Error("Agent endpoints must use HTTPS. HTTP is allowed only for local development.");
  }
  return { endpoint: url.toString(), ...(token.trim() ? { token: token.trim() } : {}) };
}
