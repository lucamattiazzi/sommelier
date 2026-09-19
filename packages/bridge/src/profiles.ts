import { randomUUID } from "node:crypto";
import { existsSync } from "node:fs";
import { chmod, mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { z } from "zod";

export const profileNameSchema = z.string().regex(/^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$/);
export const bindingSchema = z.strictObject({
  harness: z.enum(["codex", "opencode", "claude"]),
  directory: z.string().min(1),
  nativeId: z.string().min(1),
  origin: z.string().optional(),
});
export type HarnessBinding = z.infer<typeof bindingSchema>;
export const profileDirectory = () =>
  process.env.SOMMELIER_HOME ??
  process.env.AI_CDL_PAIR_HOME ??
  (existsSync(join(homedir(), ".ai-cdl-pair")) && !existsSync(join(homedir(), ".sommelier"))
    ? join(homedir(), ".ai-cdl-pair")
    : join(homedir(), ".sommelier"));
export async function readBinding(name: string): Promise<HarnessBinding | undefined> {
  profileNameSchema.parse(name);
  try {
    return bindingSchema.parse(
      JSON.parse(await readFile(join(profileDirectory(), "adapters", `${name}.json`), "utf8")),
    );
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw error;
  }
}
export async function saveBinding(name: string, binding: HarnessBinding): Promise<void> {
  profileNameSchema.parse(name);
  const dir = join(profileDirectory(), "adapters");
  await mkdir(dir, { recursive: true, mode: 0o700 });
  await chmod(dir, 0o700);
  const path = join(dir, `${name}.json`),
    temp = `${path}.${randomUUID()}.tmp`;
  await writeFile(temp, JSON.stringify(bindingSchema.parse(binding)), { mode: 0o600, flag: "wx" });
  await rename(temp, path);
}
/** One chat consumer per profile; MCP tool processes never acquire this lock. */
export async function lockProfile(name: string): Promise<() => Promise<void>> {
  profileNameSchema.parse(name);
  const dir = join(profileDirectory(), "adapters");
  await mkdir(dir, { recursive: true, mode: 0o700 });
  const path = join(dir, `${name}.lock`);
  try {
    await writeFile(path, String(process.pid), { mode: 0o600, flag: "wx" });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const pid = Number(await readFile(path, "utf8"));
    if (!Number.isSafeInteger(pid) || pid <= 0)
      throw new Error("Invalid adapter lock. Inspect it before removing it.");
    try {
      process.kill(pid, 0);
    } catch (failure) {
      if ((failure as NodeJS.ErrnoException).code !== "ESRCH") throw failure;
      await rm(path);
      return lockProfile(name);
    }
    throw new Error(
      "This Sommelier profile already has an active adapter. Stop it before reconnecting elsewhere.",
    );
  }
  return () => rm(path, { force: true });
}
