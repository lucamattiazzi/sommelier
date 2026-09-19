import { describe, expect, it } from "vitest";
import { sommelierConfigSchema } from "./config.js";

const config = {
  app: {
    id: "generate" as const,
    name: "Synthetic",
    description: "Synthetic add-in",
    version: "0.1.0",
  },
  office: {
    hosts: ["Workbook"] as const,
    permissions: "ReadWriteDocument" as const,
    requirements: { ExcelApi: "1.13" },
  },
  taskpane: {
    developmentUrl: "https://localhost:3000",
    productionUrl: "https://example.com",
  },
  commands: { label: "Open", icon: "./icon.png" },
};

describe("sommelierConfigSchema", () => {
  it("provides backward-compatible manifest branding defaults", () => {
    expect(sommelierConfigSchema.parse(config)).toMatchObject({
      app: { providerName: "Sommelier contributors" },
      commands: { groupLabel: "Sommelier" },
    });
  });

  it("accepts product-specific provider and ribbon labels", () => {
    expect(
      sommelierConfigSchema.parse({
        ...config,
        app: { ...config.app, providerName: "Acme Analytics" },
        commands: { ...config.commands, groupLabel: "Acme Copilot" },
      }),
    ).toMatchObject({
      app: { providerName: "Acme Analytics" },
      commands: { groupLabel: "Acme Copilot" },
    });
  });
});
