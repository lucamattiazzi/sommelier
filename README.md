# Sommelier

**Excels at pairing.** Your Excel workbook. Your AI agent. Together.

## Beta harness adapters

Connect **Codex, OpenCode or Claude Code** to the TaskPane without repeating a setup prompt.
[Install and reconnect](docs/harness-adapters.md) · [Release readiness](docs/v1-readiness.md)

The native adapters expose shared MCP tools for bounded reads/writes, approved chart creation and
local Excel documentation. Capture local JSONL traces or subscribe with `subscribeBridge` for evals.
Workbook traffic is encrypted end-to-end; the hosted relay cannot read payloads.


Connect a coding harness to the workbook open in Excel. The user chats in the add-in; the harness
uses a persistent WebSocket session to inspect bounded ranges and propose workbook changes.
Sommelier includes the Excel add-in, relay server, local relay CLI, reusable TypeScript libraries, and
an Agent Skill for shell-capable harnesses such as OpenCode.

Sommelier is an independent Apache-2.0 project, formerly AI-CDL Pair. Its public packages now use
`@lucamattiazzi/sommelier` and `@lucamattiazzi/sommelier-*`; the original `@ai-cdl/*` releases remain
available. See [migration](docs/migration-to-sommelier.md) and [source provenance](docs/extraction.md).
Source: [lucamattiazzi/sommelier](https://github.com/lucamattiazzi/sommelier).
This beta has not been approved for the Microsoft Marketplace. See the release readiness notes
for the remaining real Excel and harness validation.

## Use a hosted instance

Install the add-in supplied by your service operator, or deploy your own instance with the
[Docker/self-hosting guide](deploy/sommelier/README.md). End users only run the local harness adapter;
the Sommelier server runs on the operator's infrastructure.

Install the published beta from npm:

```sh
npm install -g @lucamattiazzi/sommelier@beta
sommelier pair --name desk
sommelier codex --name desk
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
pnpm server:dev
```

```sh
pnpm addin:dev
```

Sideload `apps/addin/manifest.xml` in Excel. The add-in is served at
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
[`skills/sommelier`](skills/sommelier/SKILL.md). Local browser verification and its screenshots are
reproducible with `pnpm test:browser` after `pnpm sommelier:build` and `pnpm exec playwright install chromium`.

## Repository layout

| Directory | Purpose |
| --- | --- |
| `apps/addin` | React task pane, owned icons, Office manifest |
| `apps/server` | Static host and opaque encrypted WebSocket rendezvous |
| `packages/client` | Typed Sommelier client and add-in session binding |
| `packages/bridge` | Loopback relay library and `sommelier` executable |
| `packages/protocol`, `packages/transport` | Validated RPC 0.2 and transport lifecycle |
| `packages/addin-core`, `packages/excel` | Approvals, bounded Excel tools, Office and memory adapters |
| `packages/core`, `packages/agent-http` | Existing direct-agent mode using protocol 0.1 |
| `packages/testing` | Synthetic workbook and contract helpers used by the add-in preview |
| `packages/config` | Manifest generation and packaging |
| `skills/sommelier` | Persistent bridge, RPC client, harness instructions |
| `deploy/sommelier` | Docker/Caddy deployment configuration |

The shared packages above are transitive dependencies of Sommelier. They are included so this checkout
builds without a sibling repository or unpublished packages from a registry. Relay authentication
and workbook approval checks remain part of Sommelier.

## Library usage

The public entry point is `@lucamattiazzi/sommelier-client`; its dependencies are separate publishable packages.
In a consumer that has installed the packages, Node.js 22+ can connect with:

```ts
import { createPairClient } from "@lucamattiazzi/sommelier-client";
import { createJsonSocketTransport, createEncryptedSocket, relaySocketUrl } from "@lucamattiazzi/sommelier-transport";

const pairUrl = process.env.SOMMELIER_URL;
if (!pairUrl) throw new Error("SOMMELIER_URL is required.");

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
request. See [RPC methods](skills/sommelier/references/protocol.md). Direct-agent endpoints instead
use the [HTTP 0.1 contract](docs/agent-protocol.md).

## Verification and publication

```sh
pnpm test:isolation
pnpm test:pairing
pnpm smoke:packages
pnpm smoke:consumer
```

The consumer smoke check packs the libraries and installs them in temporary consumers outside the
workspace, checking ESM, CommonJS, TypeScript declarations, and CLI entry points. It does not publish.

See [publication](docs/publishing.md) for npm artifacts and add-in deployment, and
[extraction provenance](docs/extraction.md) for the exact source revision.
