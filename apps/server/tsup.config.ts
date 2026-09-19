import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/main.ts"],
  format: ["cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  noExternal: [
    "@lucamattiazzi/sommelier-protocol",
    "@lucamattiazzi/sommelier-transport",
    "ws",
    "zod",
  ],
});
