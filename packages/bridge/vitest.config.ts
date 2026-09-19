import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: { alias: { "@lucamattiazzi/sommelier-protocol": resolve("../protocol/src/index.ts") } },
  test: { include: ["src/**/*.test.ts"] },
});
