# @lucamattiazzi/sommelier-config

## 0.2.0-beta.1

### Minor Changes

- Rename the product to Sommelier, with the `sommelier` harness connector, independent
  `@lucamattiazzi/sommelier-*` libraries, a Sommelier skill and updated add-in branding/assets.
  Preserve the workbook protocol, cryptographic domain separation, manifest identity and saved
  terminal compatibility. Existing `@ai-cdl/*` releases are not overwritten by this rename.

## 0.2.0-beta.0

### Minor Changes

- 5c48cc5: Add authenticated end-to-end encrypted socket helpers and stable Pair identities, keeping encrypted
  framing separate from workbook RPC. Support correctly sized Office manifest icons and same-origin
  support pages. The Pair application now uses remembered terminal connections and disables its
  legacy plaintext hosted routes by default.

  Update ws to 8.21.0 and the js-yaml lockfile resolution to patched releases.

## 0.1.1

### Patch Changes

- 40387f7: Ship the coordinated package release, minimal CLI split, reusable OpenAI-compatible bridge,
  read-only doctor, external-consumer smoke coverage, and configurable contract timeout.
