import { describe, expect, it } from "vitest";
import { shouldSendOnKey } from "./composer.js";

describe("composer key handling", () => {
  it("sends on a bare Enter", () => {
    expect(shouldSendOnKey({ key: "Enter" })).toBe(true);
  });

  it("keeps Shift+Enter, modifier chords, and IME composition as text input", () => {
    expect(shouldSendOnKey({ key: "Enter", shiftKey: true })).toBe(false);
    expect(shouldSendOnKey({ key: "Enter", altKey: true })).toBe(false);
    expect(shouldSendOnKey({ key: "Enter", ctrlKey: true })).toBe(false);
    expect(shouldSendOnKey({ key: "Enter", metaKey: true })).toBe(false);
    expect(shouldSendOnKey({ key: "Enter", isComposing: true })).toBe(false);
    expect(shouldSendOnKey({ key: "a" })).toBe(false);
  });
});
