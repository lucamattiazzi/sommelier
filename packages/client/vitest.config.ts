import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@lucamattiazzi/sommelier-protocol": resolve("../protocol/src/index.ts"),
      "@lucamattiazzi/sommelier-transport": resolve("../transport/src/index.ts"),
      "@lucamattiazzi/sommelier-excel": resolve("../excel/src/index.ts"),
      "@lucamattiazzi/sommelier-addin-core": resolve("../addin-core/src/index.ts"),
    },
  },
  test: { include: ["src/**/*.test.ts"] },
});
