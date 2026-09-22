# @lucamattiazzi/sommelier

## 0.2.0-beta.2

### Patch Changes

- Add a versioned offline Excel documentation catalog, shared by `excel_docs_search` /
  `excel_docs_get` MCP tools and `docs-search` / `docs-get` portable bridge commands.

- Expose `sommelier-session` for pairing, workbook RPC and reconnection through npx,
  without a global installation, project setup or preinstalled skill.

## 0.2.0-beta.1

### Minor Changes

- Rename the product to Sommelier, with the `sommelier` harness connector, independent
  `@lucamattiazzi/sommelier-*` libraries, a Sommelier skill and updated add-in branding/assets.
  Preserve the workbook protocol, cryptographic domain separation, manifest identity and saved
  terminal compatibility. Existing `@ai-cdl/*` releases are not overwritten by this rename.

### Patch Changes

- @lucamattiazzi/sommelier-protocol@0.2.0-beta.1

## 0.2.0-beta.0

### Minor Changes

- 5c48cc5: Add native Codex, OpenCode and Claude Code channel adapters with remembered conversation bindings,
  shared Excel MCP tools/documentation, and local metadata-first trace subscriptions. Add bounded,
  approved chart creation/listing and expose operation undo availability to the TaskPane. Bundle
  the portable bridge in the CLI distribution. Claude Channels remains a research preview.

  Start an authenticated local OpenCode subprocess automatically, while keeping the hosted Pair
  service configurable for independent self-hosting. End users only launch the harness adapter.

### Patch Changes

- Keep the bridge reconnecting on Node 22 after refused connections that emit only an error event.
  Deduplicate error/close handling and discard events from failed sockets.

- 5c48cc5: Add authenticated end-to-end encrypted socket helpers and stable Pair identities, keeping encrypted
  framing separate from workbook RPC. Support correctly sized Office manifest icons and same-origin
  support pages. The Pair application now uses remembered terminal connections and disables its
  legacy plaintext hosted routes by default.

  Update ws to 8.21.0 and the js-yaml lockfile resolution to patched releases.

- Updated dependencies [5c48cc5]
  - @lucamattiazzi/sommelier-protocol@0.2.0-beta.0
