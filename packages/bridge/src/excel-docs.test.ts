import { execFile } from "node:child_process";
import { cp, mkdir, mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { promisify } from "node:util";
import { expect, it, vi } from "vitest";
import { getExcelDoc, searchExcelDocs } from "./excel-docs.js";
import { callExcelTool, excelTools } from "./mcp.js";

it("retrieves exact functions, Italian aliases and topical guides with bounded results", () => {
  expect(searchExcelDocs({ query: "sum" }).results[0]?.id).toBe("sum");
  expect(searchExcelDocs({ query: "CERCA.X" }).results[0]?.id).toBe("xlookup");
  expect(searchExcelDocs({ query: "riferimenti assoluti" }).results[0]?.id).toBe("references");
  expect(searchExcelDocs({ query: "SPILL", limit: 1 }).results).toHaveLength(1);
  expect(searchExcelDocs({ query: "no-such-topic-9487" }).results).toEqual([]);
  expect(searchExcelDocs({ query: "formula", limit: 2 }).results.length).toBeLessThanOrEqual(2);
});

it("returns actionable documentation and rejects invalid requests rather than inventing pages", () => {
  const result = getExcelDoc({ id: "xlookup" });
  expect(result.document.syntax).toContain("XLOOKUP");
  expect(result.document.examples[0]?.formula).toMatch(/^=/);
  expect(result.document.compatibility).toContain("2019");
  expect(result.document.sources[0]).toMatch(/^https:\/\/support.microsoft.com\//);
  expect(result.catalogVersion).toBeTruthy();
  for (const input of [
    { query: " " },
    { query: "x".repeat(201) },
    { query: "sum", limit: 11 },
    { query: "sum", extra: true },
  ]) {
    expect(() => searchExcelDocs(input)).toThrow();
  }
  expect(() => getExcelDoc({ id: "../../private" })).toThrow();
  expect(() => getExcelDoc({ id: "invented" })).toThrow("Unknown Excel documentation ID");
});

it("serves the same documentation through MCP without contacting the workbook", async () => {
  const call = vi.fn();
  expect(
    excelTools.find((tool) => tool.name === "excel_docs_search")?.annotations?.readOnlyHint,
  ).toBe(true);
  expect(await callExcelTool(call, "excel_docs_search", { query: "CERCA.X" })).toEqual(
    searchExcelDocs({ query: "CERCA.X" }),
  );
  expect(await callExcelTool(call, "excel_docs_get", { id: "xlookup" })).toEqual(
    getExcelDoc({ id: "xlookup" }),
  );
  await expect(callExcelTool(call, "excel_docs_get", { id: "invented" })).rejects.toThrow();
  expect(call).not.toHaveBeenCalled();
});

it("provides documentation through the portable CLI without a session or project installation", async () => {
  const directory = await mkdtemp(join(tmpdir(), "sommelier-docs-"));
  const source = resolve(
    process.env.SOMMELIER_TEST_BRIDGE ?? "skills/sommelier/scripts/session.mjs",
  );
  const runtime = join(directory, "runtime");
  await mkdir(runtime);
  const script = join(runtime, "session.mjs");
  await cp(source, script);
  await cp(join(dirname(source), "lib"), join(runtime, "lib"), { recursive: true });
  const run = promisify(execFile);
  const env: NodeJS.ProcessEnv = { ...process.env, SOMMELIER_HOME: join(directory, "profiles") };
  delete env.SOMMELIER_URL;
  delete env.SOMMELIER_SESSION;
  try {
    const search = await run(
      process.execPath,
      [script, "docs-search", "--query", "CERCA.X", "--limit", "1"],
      { cwd: directory, env },
    );
    expect(JSON.parse(search.stdout)).toMatchObject({
      ok: true,
      result: searchExcelDocs({ query: "CERCA.X", limit: 1 }),
    });
    const get = await run(process.execPath, [script, "docs-get", "--id", "xlookup"], {
      cwd: directory,
      env,
    });
    expect(JSON.parse(get.stdout).result).toEqual(getExcelDoc({ id: "xlookup" }));
    await expect(
      run(process.execPath, [script, "docs-get", "--id", "invented"], { cwd: directory, env }),
    ).rejects.toMatchObject({ code: 1 });
    expect(await readdir(directory)).toEqual(["runtime"]);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
