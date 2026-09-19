/** The parts of a keyboard event that decide whether a composer should send. */
export interface ComposerKey {
  readonly key: string;
  readonly shiftKey?: boolean;
  readonly altKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
  readonly isComposing?: boolean;
}

/** Enter sends the message; Shift+Enter, modifier chords, and IME composition insert text. */
export function shouldSendOnKey(event: ComposerKey): boolean {
  if (event.key !== "Enter") return false;
  if (event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) return false;
  return !event.isComposing;
}
