import { describe, expect, it } from "vitest";
import { generateManifest } from "./manifest.js";
import type { LoadedProject } from "./project.js";

const project: LoadedProject = {
  root: "/synthetic",
  configPath: "/synthetic/sommelier.config.ts",
  appId: "8d9d84a0-5748-4b14-b45d-851812f6d27a",
  config: {
    app: {
      id: "generate",
      name: "Synthetic & Safe",
      description: "Use <bounded> workbook tools",
      version: "0.1.0",
      locale: "en-US",
      providerName: "Synthetic & Co.",
    },
    office: {
      hosts: ["Workbook"],
      permissions: "ReadWriteDocument",
      requirements: { ExcelApi: "1.13" },
    },
    taskpane: {
      developmentUrl: "https://localhost:3000",
      productionUrl: "https://agent.example.com",
    },
    commands: {
      label: "Open & inspect",
      groupLabel: "Workbook & tools",
      icon: "./assets/add-in.png",
    },
  },
};

describe("generateManifest", () => {
  it("is deterministic, escaped, and uses a valid Office version and configured icon", () => {
    const first = generateManifest(project, "production");
    expect(generateManifest(project, "production")).toBe(first);
    expect(first.startsWith('<?xml version="1.0"')).toBe(true);
    expect(first).toContain("Synthetic &amp; Safe");
    expect(first).toContain("Use &lt;bounded&gt; workbook tools");
    expect(first).toContain("<Version>1.1.0.0</Version>");
    expect(first).toContain("Project version 0.1.0 maps to Office manifest version 1.1.0.0");
    expect(first).toContain("<ProviderName>Synthetic &amp; Co.</ProviderName>");
    expect(first).toContain('id="Group.Label" DefaultValue="Workbook &amp; tools"');
    expect(first).toContain("https://agent.example.com/add-in.png");
  });

  it("does not collapse pre-1.0 and stable versions", () => {
    const prerelease = generateManifest(project, "production");
    const stable = generateManifest(
      {
        ...project,
        config: { ...project.config, app: { ...project.config.app, version: "1.1.0" } },
      },
      "production",
    );

    expect(prerelease).toContain("<Version>1.1.0.0</Version>");
    expect(stable).toContain("<Version>2.1.0.0</Version>");
    expect(prerelease).not.toBe(stable);
  });
});

it("uses configured icon sizes and a real support page", () => {
  const manifest = generateManifest(
    {
      ...project,
      config: {
        ...project.config,
        taskpane: { ...project.config.taskpane, supportPath: "/support.html" },
        commands: {
          ...project.config.commands,
          icons: {
            16: "public/icon-16.png",
            32: "public/icon-32.png",
            64: "public/icon-64.png",
            80: "public/icon-80.png",
          },
        },
      },
    },
    "production",
  );
  expect(manifest).toContain('<IconUrl DefaultValue="https://agent.example.com/icon-32.png"');
  expect(manifest).toContain(
    '<HighResolutionIconUrl DefaultValue="https://agent.example.com/icon-64.png"',
  );
  expect(manifest).toContain('<SupportUrl DefaultValue="https://agent.example.com/support.html"');
  for (const size of [16, 32, 80]) expect(manifest).toContain(`resid="Icon.${size}"`);
});
