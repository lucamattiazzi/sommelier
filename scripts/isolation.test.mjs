import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const json = (path) => JSON.parse(readFileSync(resolve(root, path), "utf8"));

test("Pair is an independent workspace with all local dependencies", () => {
  assert.equal(json("package.json").name, "sommelier");
  const manifests = ["apps", "packages"].flatMap((group) =>
    readdirSync(resolve(root, group)).map((name) => json(`${group}/${name}/package.json`)),
  );
  const names = new Set(manifests.map((manifest) => manifest.name));
  assert.ok(!names.has("@lucamattiazzi/sommelier-enterprise"));
  for (const manifest of manifests) {
    for (const [name, version] of Object.entries({
      ...manifest.dependencies,
      ...manifest.devDependencies,
    })) {
      if (version.startsWith("workspace:")) assert.ok(names.has(name), `${manifest.name}: ${name}`);
      assert.ok(!version.startsWith("link:") && !version.startsWith("file:"), name);
    }
  }
  for (const paths of Object.values(json("tsconfig.base.json").compilerOptions.paths)) {
    for (const path of paths) assert.ok(existsSync(resolve(root, path)), path);
  }
  for (const name of json(".changeset/config.json").fixed.flat()) assert.ok(names.has(name), name);
});

test("the add-in owns its icons and does not require the old examples", () => {
  for (const name of ["icon.png", "icon.svg"]) {
    assert.ok(existsSync(resolve(root, "apps/addin/public", name)), name);
  }
  for (const name of ["vite.config.ts", "sommelier.config.ts"]) {
    const source = readFileSync(resolve(root, "apps/addin", name), "utf8");
    assert.ok(!source.includes("../../examples/"), name);
  }
});
