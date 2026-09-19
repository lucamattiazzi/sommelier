---
"@lucamattiazzi/sommelier-transport": minor
"@lucamattiazzi/sommelier-config": minor
"@lucamattiazzi/sommelier": patch
---

Add authenticated end-to-end encrypted socket helpers and stable Pair identities, keeping encrypted
framing separate from workbook RPC. Support correctly sized Office manifest icons and same-origin
support pages. The Pair application now uses remembered terminal connections and disables its
legacy plaintext hosted routes by default.

Update ws to 8.21.0 and the js-yaml lockfile resolution to patched releases.

Keep the persistent bridge reconnecting on Node 22 when a refused connection emits an error
without a close event. Handle each failed socket once and never replay pending workbook requests.
