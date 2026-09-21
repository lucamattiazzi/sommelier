# @lucamattiazzi/sommelier-config

Minimal `sommelier-config` binary and typed config API for manifest generation and deterministic deployment
packaging. Sommelier uses this package to generate its development and production Office manifests.

## Manifest assets

`commands.icons` optionally supplies separate local PNG paths for keys `16`, `32`, `64` and `80`.
The manifest uses 32/64 for insertion icons and 16/32/80 for ribbon commands. `commands.icon` remains
the backward-compatible fallback. `taskpane.supportPath` can point to a same-origin support page.

`taskpane.entryPath` optionally sets the TaskPane HTML path (for example `/taskpane.html`)
relative to the configured deployment base. Icons and support URLs keep their existing base,
so the root page can introduce the product while Excel opens the TaskPane directly.
