import { execFileSync, spawnSync } from "node:child_process";
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const temporaryRoot = mkdtempSync(join(tmpdir(), "sommelier-consumer-"));
const tarballDirectory = join(temporaryRoot, "tarballs");
mkdirSync(tarballDirectory);

function run(command, args, cwd = root) {
  return execFileSync(command, args, { cwd, encoding: "utf8", stdio: "pipe" }).trim();
}

const publicPackages = readdirSync(join(root, "packages"), { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => ({
    directory: entry.name,
    manifest: JSON.parse(readFileSync(join(root, "packages", entry.name, "package.json"), "utf8")),
  }))
  .filter(({ manifest }) => manifest.private !== true)
  .sort((left, right) => left.manifest.name.localeCompare(right.manifest.name));

for (const { manifest } of publicPackages) {
  run("pnpm", ["--filter", manifest.name, "pack", "--pack-destination", tarballDirectory]);
}
const tarballs = readdirSync(tarballDirectory)
  .filter((file) => file.endsWith(".tgz"))
  .map((file) => join(tarballDirectory, file));
if (tarballs.length !== publicPackages.length) {
  throw new Error(`Expected ${publicPackages.length} tarballs, found ${tarballs.length}.`);
}
const tarballByPackage = new Map(
  publicPackages.map(({ manifest }) => {
    const filename = `${manifest.name.replace(/^@/, "").replace("/", "-")}-${manifest.version}.tgz`;
    const tarball = tarballs.find((candidate) => candidate.endsWith(filename));
    if (!tarball) throw new Error(`Missing tarball for ${manifest.name}.`);
    return [manifest.name, tarball];
  }),
);

const minimalDirectory = join(temporaryRoot, "minimal-cli-consumer");
mkdirSync(minimalDirectory);
writeFileSync(
  join(minimalDirectory, "package.json"),
  `${JSON.stringify({ name: "sommelier-minimal-smoke", private: true, type: "module" }, null, 2)}\n`,
);
const cliTarball = tarballs.find((file) => file.includes("lucamattiazzi-sommelier-config-"));
if (!cliTarball) throw new Error("Missing @lucamattiazzi/sommelier-config tarball.");
run(
  "npm",
  ["install", "--ignore-scripts", "--no-audit", "--no-fund", cliTarball],
  minimalDirectory,
);
const minimalTree = JSON.parse(run("npm", ["ls", "--all", "--json"], minimalDirectory));
const serializedTree = JSON.stringify(minimalTree);
for (const forbidden of [
  "@lucamattiazzi/sommelier-agent-http",
  "@lucamattiazzi/sommelier-evals",
  "@lucamattiazzi/sommelier-testing",
  "office-addin-debugging",
  "office-addin-dev-certs",
  "office-addin-manifest",
]) {
  if (serializedTree.includes(forbidden)) {
    throw new Error(`Minimal CLI unexpectedly installs ${forbidden}.`);
  }
}
function dependencyCount(node) {
  return Object.values(node.dependencies ?? {}).reduce(
    (total, dependency) => total + 1 + dependencyCount(dependency),
    0,
  );
}
const minimalDependencyCount = dependencyCount(minimalTree);
if (minimalDependencyCount > 75) {
  throw new Error(`Minimal CLI dependency tree regressed to ${minimalDependencyCount} packages.`);
}
run(process.execPath, [
  join(minimalDirectory, "node_modules/@lucamattiazzi/sommelier-config/dist/cli.js"),
  "--help",
]);

const singlePackageDirectory = join(temporaryRoot, "single-package-consumer");
const vendorDirectory = join(singlePackageDirectory, "vendor");
mkdirSync(vendorDirectory, { recursive: true });
for (const tarball of tarballs) copyFileSync(tarball, join(vendorDirectory, basename(tarball)));
const pairTarball = tarballByPackage.get("@lucamattiazzi/sommelier-client");
if (!pairTarball) throw new Error("Missing @lucamattiazzi/sommelier-client tarball.");
const overrides = Object.fromEntries(
  [...tarballByPackage].map(([name, tarball]) => [name, `file:./vendor/${basename(tarball)}`]),
);
writeFileSync(
  join(singlePackageDirectory, "package.json"),
  `${JSON.stringify(
    {
      name: "sommelier-single-tarball-smoke",
      private: true,
      type: "module",
      dependencies: {
        "@lucamattiazzi/sommelier-client": `file:./vendor/${basename(pairTarball)}`,
      },
    },
    null,
    2,
  )}\n`,
);
writeFileSync(
  join(singlePackageDirectory, "pnpm-workspace.yaml"),
  `packages:\n  - .\noverrides:\n${Object.entries(overrides)
    .map(([name, tarball]) => `  '${name}': ${tarball}`)
    .join("\n")}\n`,
);
run("pnpm", ["install", "--ignore-scripts"], singlePackageDirectory);
writeFileSync(
  join(singlePackageDirectory, "single.mjs"),
  `import { createPairClient } from "@lucamattiazzi/sommelier-client";\nif (typeof createPairClient !== "function") throw new Error("Missing single-package export");\n`,
);
run(process.execPath, ["single.mjs"], singlePackageDirectory);
const singleLockfile = readFileSync(join(singlePackageDirectory, "pnpm-lock.yaml"), "utf8");
for (const name of ["core", "excel", "addin-core", "protocol", "transport"]) {
  if (!singleLockfile.includes(`vendor/lucamattiazzi-sommelier-${name}-`)) {
    throw new Error(
      `Single-package install did not resolve @lucamattiazzi/sommelier-${name} from a local tarball.`,
    );
  }
}

const consumerDirectory = join(temporaryRoot, "all-packages-consumer");
mkdirSync(consumerDirectory);
writeFileSync(
  join(consumerDirectory, "package.json"),
  `${JSON.stringify({ name: "sommelier-external-smoke", private: true, type: "module" }, null, 2)}\n`,
);
run(
  "npm",
  ["install", "--ignore-scripts", "--no-audit", "--no-fund", ...tarballs],
  consumerDirectory,
);
writeFileSync(
  join(consumerDirectory, "esm.mjs"),
  `import { createAgentSession } from "@lucamattiazzi/sommelier-core";\nimport { HttpAgentAdapter } from "@lucamattiazzi/sommelier-agent-http";\nimport { createExcelTools } from "@lucamattiazzi/sommelier-excel";\nimport { runAdapterContract } from "@lucamattiazzi/sommelier-testing";\nimport { defineConfig } from "@lucamattiazzi/sommelier-config/config";\nimport { createPairClient } from "@lucamattiazzi/sommelier-client";\nimport { createAddinController } from "@lucamattiazzi/sommelier-addin-core";\nfor (const value of [createAgentSession, HttpAgentAdapter, createExcelTools, runAdapterContract, defineConfig, createPairClient, createAddinController]) {\n  if (typeof value !== "function") throw new Error("Missing ESM export");\n}\n`,
);
writeFileSync(
  join(consumerDirectory, "cjs.cjs"),
  `for (const name of ["core", "agent-http", "excel", "testing", "protocol", "transport", "addin-core", "client", ""]) {\n  const loaded = require("@lucamattiazzi/sommelier" + (name ? "-" + name : ""));\n  if (!loaded || typeof loaded !== "object") throw new Error("Missing CJS export: " + name);\n}\nconst config = require("@lucamattiazzi/sommelier-config/config");\nif (typeof config.defineConfig !== "function") throw new Error("Missing CJS config export");\n`,
);
writeFileSync(
  join(consumerDirectory, "types.ts"),
  `import type { AgentAdapter } from "@lucamattiazzi/sommelier-core";\nimport type { HttpAgentAdapterOptions } from "@lucamattiazzi/sommelier-agent-http";\nimport { createPairClient } from "@lucamattiazzi/sommelier-client";\nconst adapter: AgentAdapter | undefined = undefined;\nconst http: HttpAgentAdapterOptions = { endpoint: "https://example.com" };\nconst pairOptions: Parameters<typeof createPairClient>[0] | undefined = undefined;\nvoid [adapter, http, pairOptions];\n`,
);
writeFileSync(
  join(consumerDirectory, "tsconfig.json"),
  `${JSON.stringify(
    {
      compilerOptions: {
        strict: true,
        target: "ES2022",
        module: "NodeNext",
        moduleResolution: "NodeNext",
        skipLibCheck: true,
        noEmit: true,
      },
      include: ["types.ts"],
    },
    null,
    2,
  )}\n`,
);
run(process.execPath, ["esm.mjs"], consumerDirectory);
run(process.execPath, ["cjs.cjs"], consumerDirectory);
run(join(root, "node_modules/.bin/tsc"), ["--project", "tsconfig.json"], consumerDirectory);
for (const binary of ["sommelier-config", "sommelier-contract", "sommelier"]) {
  run(join(consumerDirectory, `node_modules/.bin/${binary}`), ["--help"], consumerDirectory);
}

const adapterDist = join(consumerDirectory, "node_modules/@lucamattiazzi/sommelier/dist");
for (const asset of [
  "skill/SKILL.md",
  "skill/scripts/session.mjs",
  "skill/scripts/lib/encrypted-socket.mjs",
]) {
  if (!readFileSync(join(adapterDist, asset), "utf8").length)
    throw new Error(`Missing adapter asset: ${asset}`);
}
const emptyProfiles = spawnSync(process.execPath, [join(adapterDist, "agent.js"), "list"], {
  cwd: consumerDirectory,
  encoding: "utf8",
  env: { ...process.env, SOMMELIER_HOME: join(temporaryRoot, "empty-profiles") },
});
if (emptyProfiles.status !== 0 || JSON.parse(emptyProfiles.stdout).terminals.length !== 0)
  throw new Error("Packed adapter cannot list profiles.");

// The existing Sommelier CLI reports usage with status 1 when no subcommand is supplied.
const pairUsage = spawnSync(join(consumerDirectory, "node_modules/.bin/sommelier-relay"), [], {
  cwd: consumerDirectory,
  encoding: "utf8",
});
if (pairUsage.status !== 1 || !pairUsage.stderr.includes("Usage: sommelier-relay relay")) {
  throw new Error("Packed Sommelier CLI did not expose its usage contract.");
}
writeFileSync(
  join(consumerDirectory, "pair.mjs"),
  `import assert from "node:assert/strict";
import { createPairClient, createPairAddinSession } from "@lucamattiazzi/sommelier-client";
import { createInMemoryTransportPair } from "@lucamattiazzi/sommelier-transport";
import { createAddinController } from "@lucamattiazzi/sommelier-addin-core";
import { InMemoryExcelAdapter } from "@lucamattiazzi/sommelier-excel";
const [addinTransport, agentTransport] = createInMemoryTransportPair();
const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1", values: [[10]] }] });
let approvals = 0;
const controller = createAddinController({
  adapter,
  workbook: { id: "synthetic", name: "Synthetic.xlsx" },
  requestApproval: () => { approvals += 1; return true; },
});
const addin = createPairAddinSession({ transport: addinTransport, controller });
const client = createPairClient({ transport: agentTransport });
await addin.start();
await client.connect();
try {
  const range = { sheetId: "Sheet1", address: "A1" };
  assert.deepEqual((await client.request("excel.range.read", { range })).values, [[10]]);
  assert.equal(approvals, 0);
  await client.request("excel.range.write", { range, values: [[25]] });
  assert.equal(approvals, 1);
  assert.deepEqual((await client.request("excel.range.read", { range })).values, [[25]]);
} finally {
  await client.close();
  await addin.stop();
}
`,
);
run(process.execPath, ["pair.mjs"], consumerDirectory);

console.log(
  `PASS external consumer: ${publicPackages.length} packed packages, one-package local override install, ESM, CJS, types, binaries, Pair read/write approval round trip; minimal CLI ${minimalDependencyCount} dependencies.`,
);
rmSync(temporaryRoot, { recursive: true, force: true });
