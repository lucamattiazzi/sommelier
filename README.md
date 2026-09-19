# AI-CDL Pair

## Beta harness adapters

Connect **Codex, OpenCode or Claude Code** to the TaskPane without repeating a setup prompt.
[Install and reconnect](docs/harness-adapters.md) · [Release readiness](docs/v1-readiness.md)

The native adapters expose shared MCP tools for bounded reads/writes, approved chart creation and
local Excel documentation. Capture local JSONL traces or subscribe with `subscribeBridge` for evals.
Workbook traffic is encrypted end-to-end; the hosted relay cannot read payloads.


Connect a coding harness to the workbook open in Excel. The user chats in the add-in; the harness
uses a persistent WebSocket session to inspect bounded ranges and propose workbook changes.
Pair includes the Excel add-in, relay server, local relay CLI, reusable TypeScript libraries, and
an Agent Skill for shell-capable harnesses such as OpenCode.

This is an independent Apache-2.0 workspace extracted from AI-CDL. It contains no enterprise
runtime, tenant authentication/authorization integration, enterprise prospect material, eval
runner, or currency/commodity demo. Existing `@ai-cdl/*` package names and public APIs are preserved.
Source: [lucamattiazzi/ai-cdl-pair](https://github.com/lucamattiazzi/ai-cdl-pair).
This beta has not been approved for the Microsoft Marketplace. See the release readiness notes
for the remaining real Excel and harness validation.

## Use a hosted instance

Install the add-in supplied by your service operator, or deploy your own instance with the
[Docker/self-hosting guide](deploy/pair/README.md). End users only run the local harness adapter;
the Pair server runs on the operator's infrastructure.

Install the published beta from npm:

```sh
npm install -g @ai-cdl/pair-cli@beta
ai-cdl-pair-agent pair --name desk
ai-cdl-pair-agent codex --name desk
# Or: opencode / claude
```

Paste the private connection URL from Excel once, at the local pairing prompt. The URL selects the
server and is remembered. There is no fixed hosted domain: self-hosting uses the same adapter and
E2EE protocol. For installation from source, see [adapter setup](docs/harness-adapters.md).

## Local development

Requires Node.js 22.12+ and Corepack for workspace development:

```sh
corepack enable
pnpm install --frozen-lockfile
pnpm build
pnpm manifest:generate
```

Run these in separate terminals from this directory:

```sh
pnpm pair:server:dev
```

```sh
pnpm pair:addin:dev
```

Sideload `apps/pair-addin/manifest.xml` in Excel. The add-in is served at
`https://localhost:3000`; the development relay listens on `127.0.0.1:3001`. The Office development
certificate may require trust on first use. A regular browser uses a synthetic in-memory workbook.

Install the portable integration for your harness once:

```sh
pnpm harness:install opencode
# Or: codex, claude, pi
```

In the pane, name your terminal and choose **Pair a terminal**. Run `pnpm agent pair --name desk`
and paste **Copy connection URL** at its local prompt. Start `pnpm agent codex --name desk` (or
`opencode` / `claude`, with the prerequisites in the adapter guide). The pane saves the connection
after mutual authentication. On later openings, start the same adapter and click **Reconnect**;
its native conversation is remembered too. Optional automatic connection applies on pane open.
The portable skill remains a fallback that must be invoked again if its agent stops waiting.

Workbook RPC, chat and tool results are end-to-end encrypted between pane and bridge. The server
routes opaque frames and sees connection metadata. The selected agent separately controls what it
sends to a model provider. Read the [transport decision and trust boundaries](docs/secure-pairing.md),
including the need to trust the client code and the lack of forward secrecy.

The [test guide](docs/pair-testing.md) covers actual Office testing. The portable skill is in
[`skills/ai-cdl-pair`](skills/ai-cdl-pair/SKILL.md). Local browser verification and its screenshots are
reproducible with `pnpm test:browser` after `pnpm pair:build` and `pnpm exec playwright install chromium`.

## Repository layout

| Directory | Purpose |
| --- | --- |
| `apps/pair-addin` | React task pane, owned icons, Office manifest |
| `apps/pair-server` | Static host and opaque encrypted WebSocket rendezvous |
| `packages/pair` | Typed Pair client and add-in session binding |
| `packages/pair-cli` | Loopback relay library and `ai-cdl-pair` executable |
| `packages/protocol`, `packages/transport` | Validated RPC 0.2 and transport lifecycle |
| `packages/addin-core`, `packages/excel` | Approvals, bounded Excel tools, Office and memory adapters |
| `packages/core`, `packages/agent-http` | Existing direct-agent mode using protocol 0.1 |
| `packages/testing` | Synthetic workbook and contract helpers used by the add-in preview |
| `packages/cli` | Manifest generation and packaging |
| `skills/ai-cdl-pair` | Persistent bridge, RPC client, harness instructions |
| `deploy/pair` | Docker/Caddy deployment configuration |

The shared packages above are transitive dependencies of Pair. They are included so this checkout
builds without a sibling repository or unpublished packages from a registry. Relay authentication
and workbook approval checks remain part of Pair.

## Library usage

The public entry point is `@ai-cdl/pair`; its dependencies are separate publishable packages.
In a consumer that has installed the packages, Node.js 22+ can connect with:

```ts
import { createPairClient } from "@ai-cdl/pair";
import { createJsonSocketTransport, createEncryptedSocket, relaySocketUrl } from "@ai-cdl/transport";

const pairUrl = process.env.AI_CDL_PAIR_URL;
if (!pairUrl) throw new Error("AI_CDL_PAIR_URL is required.");

const client = createPairClient({
  transport: createJsonSocketTransport(createEncryptedSocket(new WebSocket(relaySocketUrl(pairUrl)), pairUrl)),
});
// The pane must be open. Wait for authenticated peer presence before requests.
const authenticated = new Promise<void>((resolve) => {
  const unsubscribe = client.subscribe((event) => {
    if (event.event === "pair.peer.changed" && event.data.connected) { unsubscribe(); resolve(); }
  });
});
await client.connect();
await authenticated;
try {
  const sheets = await client.request("excel.sheet.list", {});
  console.log(sheets);
} finally {
  await client.close();
}
```

For an interactive harness, use the skill's persistent bridge instead of opening a connection per
request. See [RPC methods](skills/ai-cdl-pair/references/protocol.md). Direct-agent endpoints instead
use the [HTTP 0.1 contract](docs/agent-protocol.md).

## Verification and publication

```sh
pnpm test:isolation
pnpm pair:test
pnpm smoke:packages
pnpm smoke:consumer
```

The consumer smoke check packs the libraries and installs them in temporary consumers outside the
workspace, checking ESM, CommonJS, TypeScript declarations, and CLI entry points. It does not publish.

See [publication](docs/publishing.md) for npm artifacts and add-in deployment, and
[extraction provenance](docs/extraction.md) for the exact source revision.
