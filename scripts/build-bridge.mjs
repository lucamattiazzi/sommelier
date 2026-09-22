import { build } from "tsup";

await build({
  entry: {
    "encrypted-socket": "packages/transport/src/encrypted-socket.ts",
    "excel-docs": "packages/bridge/src/excel-docs.ts",
  },
  outDir: "skills/sommelier/scripts/lib",
  outExtension: () => ({ js: ".mjs" }),
  format: ["esm"],
  platform: "neutral",
  splitting: false,
  clean: true,
  noExternal: ["zod"],
  target: "es2022",
  silent: true,
});
