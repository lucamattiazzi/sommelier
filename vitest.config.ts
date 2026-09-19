import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@lucamattiazzi/sommelier-protocol": resolve("packages/protocol/src/index.ts"),
      "@lucamattiazzi/sommelier-transport": resolve("packages/transport/src/index.ts"),
      "@lucamattiazzi/sommelier-core": resolve("packages/core/src/index.ts"),
      "@lucamattiazzi/sommelier-excel": resolve("packages/excel/src/index.ts"),
      "@lucamattiazzi/sommelier-addin-core": resolve("packages/addin-core/src/index.ts"),
      "@lucamattiazzi/sommelier-client": resolve("packages/client/src/index.ts"),
      "@lucamattiazzi/sommelier": resolve("packages/bridge/src/index.ts"),
      "@lucamattiazzi/sommelier-agent-http": resolve("packages/agent-http/src/index.ts"),
      "@lucamattiazzi/sommelier-testing": resolve("packages/testing/src/index.ts"),
      "@lucamattiazzi/sommelier-config/project": resolve("packages/config/src/project.ts"),
      "@lucamattiazzi/sommelier-config/manifest": resolve("packages/config/src/manifest.ts"),
    },
  },
  test: {
    include: ["packages/**/*.test.ts", "apps/**/*.test.ts"],
    coverage: { provider: "v8", reporter: ["text-summary", "json-summary"] },
  },
});
