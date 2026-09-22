# Offline Excel documentation

Sommelier supplies a small, curated Excel reference alongside the existing workflow guide.
The agent searches short summaries and requests only the pages it needs. The catalog is bundled
with the bridge: queries do not reach the relay, Microsoft or a search provider. Source links
are attribution, not automatic fetches. The package runner can still contact the npm registry.

## Use from an existing agent

The generated TaskPane prompt includes the commands; a preinstalled skill is not required.
See the [published CLI commands](../packages/bridge/README.md#offline-excel-documentation).
From a source checkout, build the portable bridge once:

```sh
pnpm bridge:build
node skills/sommelier/scripts/session.mjs docs-search --query "riferimenti assoluti" --limit 3
node skills/sommelier/scripts/session.mjs docs-get --id references
```

These commands work before pairing and do not create a saved terminal profile. Results are JSON:
`{ok:true,result:{catalogVersion,results}}` for search and
`{ok:true,result:{catalogVersion,document}}` for get. Invalid input or unknown IDs exit nonzero.

Native adapters expose `excel_docs_search({query, limit?})` and `excel_docs_get({id})`, returning
that same `result` object. They execute locally rather than forwarding workbook RPC.
The existing `excel_guide` tool and `pair://docs/excel` resource explain when to consult them.

## Coverage and limits

The initial 19 pages cover SUM, SUMIF/SUMIFS, COUNTIF/COUNTIFS, IF/IFERROR, XLOOKUP,
INDEX/MATCH, FILTER, UNIQUE, SORT, LET, LAMBDA, DATE, references, tables, dynamic arrays and locale.
Documentation is in English, with selected Italian search aliases. Search uses exact-name priority
and keyword ranking; it does not call a language model or embedding service. Queries are bounded
at 200 characters and results at 10 (default 5). An empty result means the catalog has no match.

Each page includes syntax, an original synthetic example, practical notes, compatibility and
links to Microsoft Support or Learn. `catalogVersion` records the review date of the bundled
snapshot; the service does not claim those pages were fetched live. Host availability must still
be checked. Examples neither validate formulas nor grant permission to write them into a workbook.

The generated prompt remains short and points to this reference on demand. It preserves the
existing preview/approval/read-back workflow. FLAME, formula validation and dedicated formula-writing
RPCs are outside this change.

## Updating and verifying the catalog

Edit `packages/bridge/src/excel-docs-catalog.ts`, check the linked primary sources and update
`EXCEL_DOCS_VERSION` in `excel-docs.ts`. Add original explanations/examples instead of copying
articles. Keep IDs stable and aliases focused. Rebuild the portable assets and run:

```sh
pnpm bridge:build
pnpm exec vitest run packages/bridge/src/excel-docs.test.ts packages/bridge/src/harness.test.ts
```

The CLI test copies the bridge to an isolated directory and checks it without installed project
dependencies or a workbook session. For a built npm artifact, set `SOMMELIER_TEST_BRIDGE` to its
absolute `dist/skill/scripts/session.mjs` path and run `excel-docs.test.ts` again. The package
consumer smoke also checks the installed `sommelier-session` binary.
