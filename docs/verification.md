# Historical extraction verification

For current encrypted pairing results, see [the September 9 review](review-2026-09-09.md).

Verified locally with Node.js `24.19.0` and pnpm `11.21.0`:

- Fresh dependency installation and subsequent `pnpm install --frozen-lockfile` succeed.
- Two workspace isolation tests pass, checking local dependencies, TypeScript paths, release
  membership, and add-in-owned assets. Both failed before the extraction adjustments.
- 56 focused tests pass across 13 Pair, protocol, transport, Excel adapter, add-in, and server files.
  They cover persistent chat, relay sessions, bounded tools, approval, conflicts, undo, and lifecycle.
- Builds and type checks pass for the add-in, server, local relay CLI, and their dependency closure.
- Biome passes for the 23 edited code/configuration/script files checked during extraction.
- Development manifest generation and deterministic manifest checking pass without requiring
  a pre-existing pnpm CLI binary link.
- All ten library packages expose their expected ESM/CommonJS entry points. External consumers
  install packed tarballs, compile TypeScript declarations, invoke the CLIs, and execute a Pair
  read/write round trip with host approval. A separate single-package consumer installs Pair with
  its transitive dependencies resolved from local tarballs.
- The bundled server responds successfully with non-empty health, HTML, manifest, icon PNG/SVG,
  and agent bridge responses from an ephemeral loopback port.
- Local Markdown links resolve. No enterprise module, demo, private environment file, generated
  session script, certificate, database, or dependency/build directory enters the source inventory.
- The original Sommelier working tree remains clean.

No real Excel session, live OpenCode/model session, npm publication, or AppSource submission was
performed. The Docker daemon was unavailable at its configured OrbStack socket, so the container
build was not run. CI has been configured but has not run on a remote repository. These checks
establish standalone build/package behavior, not Marketplace acceptance or full Office compatibility.
