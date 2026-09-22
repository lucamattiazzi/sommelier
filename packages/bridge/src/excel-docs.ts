import { z } from "zod";
import { excelDocCatalog } from "./excel-docs-catalog.js";

export const EXCEL_DOCS_VERSION = "2026-09-22";
export const docsSearchSchema = z.strictObject({
  query: z.string().trim().min(1).max(200),
  limit: z.number().int().min(1).max(10).default(5),
});
export const docsGetSchema = z.strictObject({
  id: z.string().regex(/^[a-z][a-z0-9-]{0,79}$/),
});

const normalize = (value: string) =>
  value
    .normalize("NFKD")
    .replace(/\p{M}/gu, "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();

/** A small offline index: queries never leave the local bridge. */
export function searchExcelDocs(input: unknown) {
  const { query, limit } = docsSearchSchema.parse(input);
  const normalized = normalize(query);
  const terms = [...new Set(normalized.split(" ").filter(Boolean))];
  const results = excelDocCatalog
    .map((document) => {
      const names = [document.id, document.title, ...document.aliases].map(normalize);
      const words = new Set(
        normalize([document.summary, ...document.notes, ...names].join(" ")).split(" "),
      );
      const score =
        names.includes(normalized) && normalized.length > 0
          ? 1000
          : terms.reduce(
              (total, term) =>
                total +
                (names.some((name) => name.split(" ").includes(term))
                  ? 10
                  : words.has(term)
                    ? 1
                    : 0),
              0,
            );
      return { document, score };
    })
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score || a.document.id.localeCompare(b.document.id))
    .slice(0, limit)
    .map(({ document: { id, title, summary } }) => ({ id, title, summary }));
  return { catalogVersion: EXCEL_DOCS_VERSION, results };
}

export function getExcelDoc(input: unknown) {
  const { id } = docsGetSchema.parse(input);
  const document = excelDocCatalog.find((entry) => entry.id === id);
  if (!document)
    throw new Error("Unknown Excel documentation ID. Use excel_docs_search or docs-search first.");
  return { catalogVersion: EXCEL_DOCS_VERSION, document };
}
