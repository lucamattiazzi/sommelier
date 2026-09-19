import { cp, mkdir } from "node:fs/promises";
import { build } from "tsup";

await build({
  entry: ["src/index.ts", "src/cli.ts"],
  format: ["esm", "cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
});
await build({ entry: ["src/agent.ts"], format: ["esm"], sourcemap: true });
await mkdir("dist/skill", { recursive: true });
await cp("../../skills/sommelier", "dist/skill", { recursive: true });
await build({
  entry: { "encrypted-socket": "../transport/src/encrypted-socket.ts" },
  outDir: "dist/skill/scripts/lib",
  outExtension: () => ({ js: ".mjs" }),
  format: ["esm"],
  platform: "neutral",
  target: "es2022",
  silent: true,
});
