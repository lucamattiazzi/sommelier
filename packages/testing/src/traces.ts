import type { SessionEvent, TraceSink } from "@lucamattiazzi/sommelier-core";

/** In-memory trace sink for tests and evals. */
export class TraceCollector implements TraceSink {
  readonly events: SessionEvent[] = [];
  write(event: SessionEvent): void {
    this.events.push(event);
  }
}
