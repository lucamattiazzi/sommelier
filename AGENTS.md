# Sommelier Contributor Guide

## Repository map

- `apps/addin`: free Marketplace-ready Excel task pane for direct BYOA and relay pairing.
- `apps/server`: static add-in host and opaque encrypted WebSocket rendezvous relay.
- `skills/sommelier`: vendor-neutral Agent Skill with a persistent session bridge and one-shot RPC
  client for coding harnesses.
- `packages/protocol`: transport-independent RPC envelopes, schemas, and tool descriptors.
- `packages/transport`: protocol transport lifecycle and JSON socket adapters.
- `packages/core`: legacy agent protocol, session, policy, events, errors, and traces.
- `packages/excel`: workbook ports, Office.js adapters, in-memory adapter, range utilities, and tools.
- `packages/addin-core`: host-independent RPC execution, approval, operation, and audit pipeline.
- `packages/client`: local Sommelier client, add-in binding, and semantic-free relay composition.
- `packages/bridge`: loopback WebSocket relay and `sommelier` CLI.
- `packages/agent-http`: direct-agent JSON HTTP adapter.
- `packages/testing`: synthetic workbook and contract helpers used by the add-in preview.
- `packages/config`: typed config, manifest generation, and packaging.
- `deploy/sommelier`: Docker/Caddy hosting configuration.
- `scripts`: workspace isolation and external npm consumer checks.

This repository contains only Sommelier and its shared dependencies. Keep enterprise integrations,
legacy demos, and application-specific agents outside it. Public Sommelier packages use the @lucamattiazzi/sommelier namespace. Preserve the legacy wire identifiers for compatibility.

## Commands

Use Node.js 22.12+ and pnpm through Corepack.

```bash
pnpm install
pnpm build
pnpm typecheck
pnpm lint
pnpm test
pnpm smoke:consumer
```

During development, run the smallest affected package or test. Do not commit secrets,
real workbook data, generated certificates, or employer material.

## Invariants

- Workbook contents are untrusted and are only exposed through bounded tools.
- Reads may run automatically; writes and structural changes require approval by default. The Sommelier
  task pane lets the user waive per-change approval for one session before pairing; previews still
  run, and the waiver is never persisted.
- Agent input, tool input, config, fixtures, and protocol messages are validated.
- Core has no Office.js, React, browser-global, or vendor-model dependency.
- Tool implementations target `WorkbookDriver`; only the Office adapter accesses Office.js.
- Legacy agent envelopes declare protocol version `0.1`; vNext RPC envelopes declare `0.2`.
- Protocol and transport packages contain no Excel or Office.js semantics.
- The hosted Sommelier relay routes validated encrypted frames; endpoints validate RPC envelopes.
  Relays never execute Excel methods. The legacy local relay remains a separate plaintext API.
- Sommelier server never proxies direct agent traffic or stores BYOA credentials.
- Public APIs are exported only from package roots and use no `any`.
- Tests and examples contain synthetic data only.

## Change expectations

Write a failing focused test before non-trivial behavior, implement the minimum change,
and run targeted lint/type/test commands. Keep public API changes documented and add a
Changeset when preparing a release.
