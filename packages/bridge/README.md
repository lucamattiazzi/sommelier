# @lucamattiazzi/sommelier

`sommelier` connects Codex, OpenCode or Claude Code to the Excel Sommelier TaskPane. Requires
Node.js 22.12+ and a signed-in harness. The included skill, bridge and crypto are self-contained.

## Pair your existing agent

Copy the setup prompt from the Excel TaskPane into your running agent. No skill or global
installation is needed: the prompt uses the standalone `sommelier-session` command through npx.
It includes the private connection URL, RPC commands and listening loop.

```sh
# Resume a previously saved connection (first pairing uses the TaskPane prompt):
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier-session start --name desk
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier-session status --name desk
```

npx downloads the package into its cache. Connection profiles are saved in `~/.sommelier`;
your project needs no `package.json` or `node_modules`. Keep the cache while the bridge is running.
The agent must keep listening; starting the bridge alone cannot wake a suspended agent.

## Optional native adapters

Run these directly through npx as well:

```sh
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier pair --name desk
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier codex --name desk
# Replace codex with opencode or claude (Claude requires local channel consent).
```

Sommelier asks for the private TaskPane URL once; subsequent starts use the saved profile and native
conversation, including the selected hosted or self-hosted Sommelier instance. End users do not start
a Sommelier relay. OpenCode's local harness API starts automatically on an authenticated loopback port
and stops with the adapter. Optional `--server` attaches an API you already manage instead.
Claude must stay open. The Codex adapter
starts App Server, not an arbitrary existing TUI. Native Codex command/file approvals are declined;
Excel writes/charts retain their TaskPane approval flow.

Public exports `bridgeCommand` and `subscribeBridge` allow local tool/eval integrations. Trace
payloads are excluded by default; `includeContent: true` opts in. Never commit full traces or saved
profiles. A timeout never causes mutation replay. Slow trace listeners cannot block the workbook.

`sommelier-relay relay --port 4310` remains the legacy **plaintext loopback relay**; the native adapters
require encrypted `/connect` profiles from the hosted Sommelier relay instead.

## Offline Excel documentation

The same curated catalog is available through MCP tools `excel_docs_search` / `excel_docs_get`
and standalone CLI commands. No workbook connection, skill, API key or global installation is needed:

```sh
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier-session docs-search --query "CERCA.X" --limit 3
npx --yes --ignore-scripts --package=@lucamattiazzi/sommelier@0.2.0-beta.2 -- sommelier-session docs-get --id xlookup
```

Search returns IDs and short summaries; get returns syntax, synthetic examples, pitfalls,
compatibility notes and Microsoft source links. Both include `catalogVersion`; CLI results use
`{ok:true,result:...}`. Queries accept 1–200 characters, limits 1–10 (default 5). Unknown IDs fail.
Function names and English/Italian aliases work; this is a small keyword index, not semantic search.
The catalog is bundled and versioned, not live documentation or a formula validator. The commands
perform no network requests; npx may access the package registry to download/cache the package.
