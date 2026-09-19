# Native harness adapters (v1)

Pair connects the open Excel workbook to **OpenCode, Codex or Claude Code**. The model and its
credentials stay with the selected harness. A skill alone cannot start an agent turn: these
adapters deliver incoming TaskPane messages to the harness's native input mechanism.

## Install and associate once

Requires Node.js 22.12+ and the chosen harness installed and signed in. Use the add-in supplied by
your Pair service operator, or your own self-hosted instance. **End users do not run a Pair server.**
The local adapter connects outbound to the instance chosen during pairing; no public inbound port
or fixed service domain is required on the user's machine.

Install the published beta adapter:

```sh
npm install -g @ai-cdl/pair-cli@beta
ai-cdl-pair-agent pair --name desk
```

In Excel choose **Pair a terminal → Copy connection URL**, and paste the URL when the terminal asks.
It contains the selected server address and a private key: use the local setup prompt, never a model
chat or a tracked file. Both are remembered locally. There is no central service domain compiled
into the adapter. Profiles can belong to different self-hosted instances.

Alternatively, build a trusted source checkout:

```sh
pnpm install --frozen-lockfile
pnpm bridge:build
pnpm --filter @ai-cdl/protocol --filter @ai-cdl/pair-cli build
pnpm agent pair --name desk
```

In the commands below, replace `ai-cdl-pair-agent` with `pnpm agent` when using that checkout.
Use the `beta` tag explicitly; stable versions of the shared packages remain on `latest`.

## Codex

```sh
ai-cdl-pair-agent codex --name desk
```

Starts a local `codex app-server`, creates a thread once, configures the Excel MCP tools for that
thread, and delivers TaskPane messages using `turn/start`. Repeating the same command resumes the
saved thread. Codex persists a conversation after its first turn: if you stop before sending any
message and resume fails, use `--new-session`. To explicitly associate an existing **stored** thread on first use:

```sh
ai-cdl-pair-agent codex --name desk --session THREAD_ID
```

This does not take over an arbitrary running Codex terminal or desktop session. Do not operate the
same thread concurrently in another client. Native command/file approval requests are declined by
this v1 adapter; Excel mutation approvals remain available in the TaskPane. No global Codex config,
model selection or provider credential is overwritten. Tool access is pre-authorized only for
Pair’s own MCP server; the TaskPane still enforces approval for every workbook mutation by default.

## OpenCode

Run one command from your project directory:

```sh
ai-cdl-pair-agent opencode --name desk
```

The adapter starts and stops its own local OpenCode API on an available loopback port, protected by
a fresh password. This is a local harness subprocess, separate from the hosted Pair relay. It adds
the Excel MCP tools and creates a session once, or resumes the remembered native session ID on
later starts. The temporary port and password are not saved. No provider configuration is changed.

To attach to an **already running** OpenCode API instead, use
`--server http://127.0.0.1:4096` and optionally `--session SESSION_ID`. That explicit API origin is
remembered; you manage that external process yourself. If password protected, set
`OPENCODE_SERVER_PASSWORD` (and optionally `OPENCODE_SERVER_USERNAME`) locally. `--new-session`
returns to automatic launch unless you also supply `--server`.

The adapter allows loopback HTTP only and refuses redirects: plaintext agent data never traverses
the Pair relay. It refuses to submit while that OpenCode session reports itself busy.

OpenCode MCP registration is shared within its server/project context. Use a dedicated project or
server for each workbook when isolating multiple agents; an authorized harness controls access to
its tools. Native permissions requiring interaction need an attached OpenCode UI; use an existing API with
`--server` for that workflow. The adapter does not approve native permissions automatically.

## Claude Code

```sh
ai-cdl-pair-agent claude --name desk
```

Launches Claude Code with an Excel MCP channel and an explicit session ID, remembers that ID, and
uses `--resume` on later launches. The channel delivers TaskPane requests into that running session;
Claude uses the shared Excel tools and `pair_reply` to answer in Excel.

**Channels is a research preview.** Custom channels require Claude's development-channel flag and
local consent; organization policy can block them. The launcher prints this requirement and lets
Claude present its own consent UI. This is a v1 preview integration, not an AppSource dependency.
Keep the Claude session open. If an initial launch failed before creating its conversation, use
`--new-session` to create a replacement rather than attempting to resume a nonexistent session.

## Reconnect and manage

```sh
ai-cdl-pair-agent list
ai-cdl-pair-agent status --name desk
ai-cdl-pair-agent codex --name desk          # or opencode / claude
ai-cdl-pair-agent stop --name desk
ai-cdl-pair-agent forget --name desk
```

Select the corresponding saved terminal in Excel and choose **Reconnect**, or enable automatic
connection on pane open. The local adapter must be running. The relay cannot start a stopped
computer or process. One adapter consumes chat per named profile; duplicate consumers are refused.
Use another name for another workbook. `--directory PATH` selects the project on first association;
changing a remembered directory or harness requires `--new-session`.

Stop the adapter before forgetting its profile. Forget in Excel too to remove both copies of the
association. Credentials live in owner-only `~/.ai-cdl-pair/<name>.json`; native session mappings live
in its `adapters` directory. `AI_CDL_PAIR_HOME` changes this location. Local files are not encrypted
at rest. Windows named-pipe ACLs still require platform validation.

Pending workbook requests fail on disconnect and are never replayed. A failed or timed-out native
turn stops the adapter to avoid overlapping turns. Reconnect, inspect the workbook and send a new
task. Chat is not a durable delivery queue: disconnects can discard pending messages. Range undo
history lasts only for the current TaskPane connection; chart undo is not implemented in v1.

## Tools and documentation

All adapters use the same MCP tools generated from the protocol registry: context, sheet/table
inspection, bounded range reads/writes, chart listing/creation, previews, commits, range undo and
user interaction. `excel_guide` and the MCP resource `pair://docs/excel` provide workflow guidance
without requiring a remote documentation service. Tool schemas are returned by MCP `tools/list`.
The [portable skill](../skills/ai-cdl-pair/SKILL.md) remains available as a manual fallback.

## Traces and eval integration

```sh
ai-cdl-pair-agent trace --name desk --out ./run.jsonl
# Explicit opt-in to record cells/prompts/results locally:
ai-cdl-pair-agent trace --name desk --out ./run-with-content.jsonl --include-content
```

The output is created exclusively with mode 0600; an existing file is not overwritten. Start
capture before the task. By default events contain timestamps, kinds, RPC identifiers/methods, outcomes, durations and chat message IDs,
not workbook content. Full capture is intentionally private local data and must not be committed.
These are workbook RPC/chat traces, not complete provider reasoning/token traces. Subscribe to the
chosen harness separately for provider-specific eval data.

The public `@ai-cdl/pair-cli` API supports arbitrary consumers without consuming chat:

```ts
import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { subscribeBridge } from "@ai-cdl/pair-cli";

const profile = JSON.parse(await readFile(join(homedir(), ".ai-cdl-pair/desk.json"), "utf8"));
const unsubscribe = subscribeBridge(profile.session, (event) => {
  if (event.kind === "rpc.response") console.log(event.id, event.method);
}, { onError: (error) => console.error(error.message) });
process.once("SIGINT", unsubscribe);
```

Listener exceptions are isolated. Slow consumers have bounded buffers and are disconnected rather
than blocking workbook execution. Traces are best-effort, not a transactional audit log. There is
no telemetry upload to Pair servers.

## Protocol references

Verified against installed Codex 0.153.4, OpenCode 1.18.29 and Claude Code 2.1.236 capabilities on
2026-09-11. See [verification](v1-readiness.md) for the distinction between tested contracts and
live model/Excel runs.

- [Codex App Server](https://learn.chatgpt.com/docs/app-server)
- [OpenCode server](https://opencode.ai/docs/server/)
- [Claude Channels contract](https://code.claude.com/docs/en/channels-reference)
- [Excel charts](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-charts)
