import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { createJiti } from "jiti";
import { z } from "zod";
import { type ResolvedSommelierConfig, sommelierConfigSchema } from "./config.js";

/** Loaded config plus resolved project paths. */
export interface LoadedProject {
  readonly root: string;
  readonly configPath: string;
  readonly config: ResolvedSommelierConfig;
  readonly appId: string;
}

/** Load, validate, and resolve the generated app identity for an Sommelier project. */
export async function loadProject(cwd = process.cwd(), persistId = true): Promise<LoadedProject> {
  const configPath = resolve(cwd, "sommelier.config.ts");
  try {
    await access(configPath);
  } catch {
    throw new Error(
      "SOMMELIER_CONFIG_NOT_FOUND: Create sommelier.config.ts or run `sommelier-config init`.",
    );
  }
  const jiti = createJiti(configPath, { interopDefault: true });
  const imported = await jiti.import(configPath, { default: true });
  const parsed = sommelierConfigSchema.safeParse(imported);
  if (!parsed.success) {
    const issues = z.prettifyError(parsed.error);
    throw new Error(
      `SOMMELIER_CONFIG_INVALID\n${issues}\nFix the listed paths in sommelier.config.ts.`,
    );
  }
  let appId = parsed.data.app.id;
  if (appId === "generate") {
    const idPath = resolve(cwd, ".sommelier/app-id");
    try {
      appId = (await readFile(idPath, "utf8")).trim();
      z.string().uuid().parse(appId);
    } catch {
      if (!persistId)
        throw new Error(
          "SOMMELIER_APP_ID_MISSING: Generate the persistent app ID before running a drift check.",
        );
      appId = crypto.randomUUID();
      await mkdir(dirname(idPath), { recursive: true });
      await writeFile(idPath, `${appId}\n`, { flag: "wx" }).catch(async (error: unknown) => {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
        appId = (await readFile(idPath, "utf8")).trim();
      });
    }
  }
  return { root: cwd, configPath, config: parsed.data, appId };
}
