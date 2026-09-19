import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, expect, it, vi } from "vitest";
import { lockProfile, profileDirectory, readBinding, saveBinding } from "./profiles.js";

afterEach(() => vi.unstubAllEnvs());
it("remembers the harness conversation in owner-only files and excludes a second consumer", async () => {
  const dir = await mkdtemp(join(tmpdir(), "pair-binding-"));
  vi.stubEnv("AI_CDL_PAIR_HOME", dir);
  try {
    const binding = {
      harness: "codex" as const,
      nativeId: "thread-one",
      directory: "/tmp/synthetic",
    };
    expect(await readBinding("desk")).toBeUndefined();
    await saveBinding("desk", binding);
    expect(await readBinding("desk")).toEqual(binding);
    expect((await stat(join(dir, "adapters", "desk.json"))).mode & 0o777).toBe(0o600);
    const unlock = await lockProfile("desk");
    await expect(lockProfile("desk")).rejects.toThrow("active adapter");
    await unlock();
    const unlockAgain = await lockProfile("desk");
    await unlockAgain();
    await expect(saveBinding("../escape", binding)).rejects.toThrow();
    await writeFile(join(dir, "adapters", "broken.json"), '{"harness":"codex"}');
    await expect(readBinding("broken")).rejects.toThrow();
    expect(await readFile(join(dir, "adapters", "desk.json"), "utf8")).not.toContain("secret");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

it("prefers the Sommelier profile setting and still accepts the old setting", () => {
  vi.stubEnv("AI_CDL_PAIR_HOME", "/tmp/legacy-synthetic");
  vi.stubEnv("SOMMELIER_HOME", "/tmp/sommelier-synthetic");
  expect(profileDirectory()).toBe("/tmp/sommelier-synthetic");
  vi.stubEnv("SOMMELIER_HOME", undefined);
  expect(profileDirectory()).toBe("/tmp/legacy-synthetic");
});
