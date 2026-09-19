import type { AgentAdapter, AgentEvent, AgentTurnRequest } from "@lucamattiazzi/sommelier-core";

/** Function producing deterministic adapter events from a turn request. */
export type AgentScript = (
  request: AgentTurnRequest,
) => readonly AgentEvent[] | Promise<readonly AgentEvent[]>;

/** Deterministic adapter for unit tests and scripted eval scenarios. */
export class ScriptedAgentAdapter implements AgentAdapter {
  readonly #script: AgentScript;
  readonly requests: AgentTurnRequest[] = [];

  constructor(script: AgentScript) {
    this.#script = script;
  }

  async *runTurn(
    request: AgentTurnRequest,
    options: { signal: AbortSignal },
  ): AsyncIterable<AgentEvent> {
    this.requests.push(request);
    const events = await this.#script(request);
    for (const event of events) {
      if (options.signal.aborted) throw options.signal.reason;
      yield event;
    }
  }
}
