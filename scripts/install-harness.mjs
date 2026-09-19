#!/usr/bin/env node
import { access, cp, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const locations = {
  codex: ".agents/skills",
  opencode: ".config/opencode/skills",
  claude: ".claude/skills",
  pi: ".pi/agent/skills",
};
const harness = process.argv[2];
if (!Object.hasOwn(locations, harness)) {
  console.error(
    "Usage: node scripts/install-harness.mjs codex|opencode|claude|pi [--home path] [--replace]",
  );
  process.exit(1);
}
const homeIndex = process.argv.indexOf("--home");
if (homeIndex >= 0 && !process.argv[homeIndex + 1]) throw new Error("--home requires a directory.");
const home = homeIndex < 0 ? homedir() : process.argv[homeIndex + 1];
const source = fileURLToPath(new URL("../skills/sommelier", import.meta.url));
await access(join(source, "scripts/lib/encrypted-socket.mjs")).catch(() => {
  throw new Error("Build the portable bridge first: pnpm bridge:build");
});
const destination = join(home, locations[harness], "sommelier");
const replace = process.argv.includes("--replace");
if (
  !replace &&
  (await access(destination).then(
    () => true,
    () => false,
  ))
)
  throw new Error("Skill already exists. Use --replace to update it.");
await mkdir(join(home, locations[harness]), { recursive: true });
await cp(source, destination, { recursive: true, force: replace, errorOnExist: !replace });
console.log(`Installed ${harness}: ${destination}`);
