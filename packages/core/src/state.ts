import { SommelierError } from "./errors.js";

/** Explicit session execution states. */
export type SessionState =
  | "idle"
  | "running_agent"
  | "executing_read_tool"
  | "awaiting_approval"
  | "executing_write_tool"
  | "completed"
  | "cancelled"
  | "failed";

const transitions: Readonly<Record<SessionState, readonly SessionState[]>> = {
  idle: ["running_agent"],
  running_agent: ["executing_read_tool", "awaiting_approval", "completed", "cancelled", "failed"],
  executing_read_tool: ["running_agent", "cancelled", "failed"],
  awaiting_approval: ["executing_write_tool", "running_agent", "cancelled", "failed"],
  executing_write_tool: ["running_agent", "cancelled", "failed"],
  completed: ["running_agent"],
  cancelled: ["running_agent"],
  failed: ["running_agent"],
};

/** Small deterministic state machine used by every agent session. */
export class SessionStateMachine {
  #state: SessionState = "idle";
  get state(): SessionState {
    return this.#state;
  }

  transition(next: SessionState): void {
    if (!transitions[this.#state].includes(next)) {
      throw new SommelierError({
        code: "SOMMELIER_INVALID_STATE_TRANSITION",
        message: `Cannot transition an Sommelier session from ${this.#state} to ${next}.`,
        context: { from: this.#state, to: next },
        suggestedAction: "Cancel the active turn before starting another operation.",
      });
    }
    this.#state = next;
  }
}
