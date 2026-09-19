# Contributing

Use Node.js 22+ and pnpm through Corepack. Install with `pnpm install --frozen-lockfile` and build
with `pnpm build` before starting the applications.

Write a focused failing test before changing non-trivial behavior, implement the minimum change,
and run checks scoped to the affected files or package. For example:

```sh
pnpm exec vitest run packages/client/src/lifecycle.test.ts
pnpm --filter @lucamattiazzi/sommelier-client typecheck
pnpm exec biome check packages/client/src/index.ts
```

Keep test workbooks synthetic. Never commit pairing URLs, credentials, private workbook contents,
certificates, or session state. Preserve bounded reads, validation, and host approval by default.
Document public API changes and add a Changeset for a release. See [AGENTS.md](AGENTS.md).
