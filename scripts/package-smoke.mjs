import { createRequire } from "node:module";

const expected = [
  ["protocol", ".", "parseProtocolMessage"],
  ["transport", ".", "createInMemoryTransportPair"],
  ["core", ".", "createAgentSession"],
  ["excel", ".", "createExcelTools"],
  ["addin-core", ".", "createAddinController"],
  ["client", ".", "createPairClient"],
  ["bridge", ".", "createPairRelayServer"],
  ["agent-http", ".", "HttpAgentAdapter"],
  ["testing", ".", "VirtualWorkbookDriver"],
  ["config", "./config", "defineConfig"],
];
const require = createRequire(import.meta.url);

for (const [directory, subpath, exportName] of expected) {
  const packageFile = new URL(`../packages/${directory}/package.json`, import.meta.url);
  const packageJson = require(packageFile.pathname);
  const conditions = packageJson.exports[subpath];
  if (!conditions?.types || !conditions?.import || !conditions?.require) {
    throw new Error(
      `${packageJson.name}${subpath === "." ? "" : subpath} lacks explicit type/import/require exports.`,
    );
  }
  const esm = await import(
    new URL(`../packages/${directory}/${conditions.import}`, import.meta.url)
  );
  const cjs = require(
    new URL(`../packages/${directory}/${conditions.require}`, import.meta.url).pathname,
  );
  if (!(exportName in esm) || !(exportName in cjs)) {
    throw new Error(`${packageJson.name} is missing ${exportName} from an ESM or CommonJS entry.`);
  }
  const displayName = `${packageJson.name}${subpath === "." ? "" : subpath.slice(1)}`;
  console.log(`PASS ${displayName}: ESM/CommonJS/types expose ${exportName}`);
}
