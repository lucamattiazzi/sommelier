import { build } from "tsup";

await build({
  entry: { "encrypted-socket": "packages/transport/src/encrypted-socket.ts" },
  outDir: "skills/sommelier/scripts/lib",
  outExtension: () => ({ js: ".mjs" }),
  format: ["esm"],
  platform: "neutral",
  target: "es2022",
  silent: true,
});
