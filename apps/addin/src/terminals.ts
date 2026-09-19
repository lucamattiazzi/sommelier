import type { PairIdentity } from "@lucamattiazzi/sommelier-transport";

export interface SavedTerminal extends PairIdentity {
  readonly name: string;
  readonly agentOrigin: string;
  readonly autoConnect: boolean;
}
export const TERMINALS_KEY = "ai-cdl-pair-terminals-v1";

/** Allowlist persisted fields: workbook data and approval waivers are never saved. */
export function parseSavedTerminals(json: string | null): SavedTerminal[] {
  try {
    const input: unknown = JSON.parse(json ?? "[]");
    if (!Array.isArray(input) || input.length > 20) return [];
    const records: SavedTerminal[] = [];
    for (const item of input as unknown[]) {
      if (!item || typeof item !== "object") continue;
      const r = item as Record<string, unknown>;
      if (
        typeof r.id !== "string" ||
        !/^[a-f\d-]{36}$/i.test(r.id) ||
        typeof r.secret !== "string" ||
        !/^[\w-]{43}$/.test(r.secret) ||
        typeof r.name !== "string" ||
        !r.name.trim() ||
        r.name.length > 64 ||
        typeof r.agentOrigin !== "string" ||
        typeof r.autoConnect !== "boolean"
      )
        continue;
      const url = new URL(r.agentOrigin);
      if (
        url.username ||
        url.password ||
        url.hash ||
        url.search ||
        (url.protocol !== "wss:" &&
          !(url.protocol === "ws:" && ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname)))
      )
        continue;
      if (records.some((record) => record.id === r.id)) continue;
      records.push({
        id: r.id,
        secret: r.secret,
        name: r.name,
        agentOrigin: url.origin,
        autoConnect: r.autoConnect && !records.some((record) => record.autoConnect),
      });
    }
    return records;
  } catch {
    return [];
  }
}

export function rememberTerminal(
  records: readonly SavedTerminal[],
  terminal: SavedTerminal,
): SavedTerminal[] {
  return [
    ...records
      .filter((record) => record.id !== terminal.id)
      .map((record) => (terminal.autoConnect ? { ...record, autoConnect: false } : record)),
    terminal,
  ].slice(-20);
}

export function terminalSetupPrompt(agentUrl: string): string {
  return `Use the installed sommelier skill to associate this Excel task pane with this terminal. Start its scripts/session.mjs with a recognizable --name (for example desk). Pass this URL in SOMMELIER_URL for the first start only:\n\n${agentUrl}\n\nThe URL includes an end-to-end secret. Do not echo it, log it, put it in source files, or send it to another service. The bridge saves it locally with owner-only permissions. On later uses, start with the same --name without another URL. Consume messages with next, use bounded workbook RPC calls, and answer with reply. Always preview mutations and let the task pane handle approval. Follow the skill's protocol reference.`;
}
