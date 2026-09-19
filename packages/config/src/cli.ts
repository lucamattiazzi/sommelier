#!/usr/bin/env node

import { createHash } from "node:crypto";
import { access, mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Command } from "commander";
import {
  generateManifest,
  type ManifestEnvironment,
  officeManifestVersion,
  writeManifest,
} from "./manifest.js";
import { loadProject } from "./project.js";

const program = new Command()
  .name("sommelier-config")
  .description("Generate and package safe Sommelier Excel integrations.")
  .version("0.1.0");
const invocationDirectory = process.env.INIT_CWD ?? process.cwd();

program
  .command("init")
  .description("Create a minimal typed Sommelier config in the current directory.")
  .action(async () => {
    const target = resolve(invocationDirectory, "sommelier.config.ts");
    try {
      await access(target);
      throw new Error(
        "SOMMELIER_INIT_EXISTS: sommelier.config.ts already exists; no files were overwritten.",
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
    await writeFile(
      target,
      `import { defineConfig } from "@lucamattiazzi/sommelier-config/config";\n\nexport default defineConfig({\n  app: { id: "generate", name: "Workbook Agent", description: "Use an existing agent from Excel", version: "0.1.0", providerName: "Your company" },\n  office: { hosts: ["Workbook"], permissions: "ReadWriteDocument", requirements: { ExcelApi: "1.13" } },\n  taskpane: { developmentUrl: "https://localhost:3000", productionUrl: "https://agent.example.com" },\n  commands: { label: "Open Agent", groupLabel: "Workbook Agent", icon: "./public/icon.png" },\n});\n`,
    );
    console.log(`Created ${target}. Review URLs and run sommelier-config manifest generate.`);
  });

const manifest = program.command("manifest").description("Generate Excel add-in manifests.");
manifest
  .command("generate")
  .description("Generate deterministic add-in-only manifest XML.")
  .option("-e, --environment <environment>", "development or production", "development")
  .option("-o, --output <path>", "output path", "manifest.xml")
  .option("--check", "fail when generated output differs")
  .action(
    async (options: { environment: ManifestEnvironment; output: string; check?: boolean }) => {
      const project = await loadProject(invocationDirectory, !options.check);
      const target = await writeManifest(
        project,
        options.environment,
        options.output,
        options.check ?? false,
      );
      console.log(
        `${options.check ? "Manifest is current" : "Generated"}: ${target}\nProject version ${project.config.app.version} maps to Office manifest version ${officeManifestVersion(project.config.app.version)}.`,
      );
    },
  );

async function createDeploymentPackage(): Promise<string> {
  const project = await loadProject(invocationDirectory);
  const directory = resolve(project.root, "dist/sommelier-addin");
  await mkdir(directory, { recursive: true });
  const xml = generateManifest(project, "production");
  await writeFile(resolve(directory, "manifest.xml"), xml);
  const checksum = createHash("sha256").update(xml).digest("hex");
  await writeFile(
    resolve(directory, "checksums.json"),
    `${JSON.stringify({ "manifest.xml": `sha256:${checksum}` }, null, 2)}\n`,
  );
  await writeFile(
    resolve(directory, "metadata.json"),
    `${JSON.stringify(
      {
        schemaVersion: "0.1",
        appId: project.appId,
        appVersion: project.config.app.version,
        generatedAt: new Date(0).toISOString(),
      },
      null,
      2,
    )}\n`,
  );
  await writeFile(
    resolve(directory, "deployment-guide.md"),
    deploymentGuide(project.config.app.name),
  );
  return directory;
}

program
  .command("package")
  .description("Create deterministic deployment artifacts without implying tenant deployment.")
  .action(async () => {
    console.log(`Deployment directory: ${await createDeploymentPackage()}`);
  });

const deploy = program
  .command("deploy")
  .description("Prepare, but do not execute, Microsoft 365 deployment.");
deploy
  .command("prepare")
  .requiredOption("--target <target>", "m365")
  .action(async (options: { target: string }) => {
    if (options.target !== "m365") {
      throw new Error("SOMMELIER_DEPLOY_TARGET_UNSUPPORTED: The MVP supports only --target m365.");
    }
    console.log(`Microsoft 365 deployment artifacts: ${await createDeploymentPackage()}`);
  });

function deploymentGuide(name: string): string {
  return `# Microsoft 365 deployment guide for ${name}\n\n1. Host the production task pane at the HTTPS URL configured in \`sommelier.config.ts\`.\n2. In the Microsoft 365 admin center, open **Settings → Integrated apps → Upload custom apps**.\n3. Upload \`manifest.xml\`, review requested permissions, then select the intended users or groups.\n4. Finish deployment. Propagation commonly takes hours and can take up to 24 hours.\n5. To roll back, remove or edit the app assignment in Integrated apps; do not delete hosted assets until clients have stopped using them.\n\nThe task pane hosting and manifest deployment are separate. Sommelier never requests tenant administrator credentials.\n`;
}

program.parseAsync().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
