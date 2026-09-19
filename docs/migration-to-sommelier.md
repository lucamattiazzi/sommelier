# Migrating to Sommelier

Sommelier is the new name for AI-CDL Pair. The first release under this name is
`0.2.0-beta.1`. Its tagline is **Excels at pairing.** The repository is
[github.com/lucamattiazzi/sommelier](https://github.com/lucamattiazzi/sommelier).

The unscoped npm name `sommelier` belongs to an unrelated project. Install our package explicitly:

```sh
npm install -g @lucamattiazzi/sommelier@beta
sommelier list
sommelier codex --name desk
# Alternatives: sommelier opencode --name desk / sommelier claude --name desk
```

The executable is `sommelier`. Existing `@ai-cdl/*` packages remain available; installing
Sommelier does not overwrite them. Stop an old running bridge before starting its replacement.

## Library and command names

| Previous package | Sommelier package | Command, where applicable |
| --- | --- | --- |
| `@ai-cdl/pair-cli` | `@lucamattiazzi/sommelier` | `sommelier`, `sommelier-relay` |
| `@ai-cdl/pair` | `@lucamattiazzi/sommelier-client` | |
| `@ai-cdl/cli` | `@lucamattiazzi/sommelier-config` | `sommelier-config` |
| `@ai-cdl/core` | `@lucamattiazzi/sommelier-core` | |
| `@ai-cdl/excel` | `@lucamattiazzi/sommelier-excel` | |
| `@ai-cdl/protocol` | `@lucamattiazzi/sommelier-protocol` | |
| `@ai-cdl/transport` | `@lucamattiazzi/sommelier-transport` | |
| `@ai-cdl/addin-core` | `@lucamattiazzi/sommelier-addin-core` | |
| `@ai-cdl/agent-http` | `@lucamattiazzi/sommelier-agent-http` | |
| `@ai-cdl/testing` | `@lucamattiazzi/sommelier-testing` | `sommelier-contract` |

Update imports and dependency names together. Configuration moves to `sommelier.config.ts`;
`AiCdlConfig`, `aiCdlConfigSchema`, `AiCdlError`, `asAiCdlError` and related types now use
`Sommelier` / `sommelier`. Diagnostic codes use `SOMMELIER_*`. The pairing API names such as
`createPairClient` remain descriptive API names. Historical changelog entries predate these new
package names and refer to releases under the original namespace.

## Remembered connections

New installations store profiles in `~/.sommelier`. If only `~/.ai-cdl-pair` exists, both the
adapter and portable bridge reuse it. `SOMMELIER_HOME` overrides this choice; the previous
`AI_CDL_PAIR_HOME` is accepted as a fallback. If both directories exist, select the intended one
explicitly. `SOMMELIER_URL` and `SOMMELIER_SESSION` replace the old `AI_CDL_PAIR_*` environment
variables, which remain fallback aliases.

The add-in identity, browser storage keys, encrypted wire identifiers, native MCP registration
key, RPC methods and trace event names remain unchanged. This preserves associations when the
TaskPane stays on the same origin. A new hosting origin needs a new association because browser
storage is origin-specific. The direct-agent HTTP protocol retains `X-AI-CDL-Protocol-Version`
for compatibility with existing endpoints.

Install the portable skill from `skills/sommelier`; remove the old installed `ai-cdl-pair` skill
once the new one works to avoid duplicate skill discovery. Existing native session IDs are kept.

## Hosting and add-in assets

The directories are now `apps/addin`, `apps/server` and `deploy/sommelier`.
Use `SOMMELIER_DOMAIN` for Docker/Caddy and `SOMMELIER_PUBLIC_ORIGIN` for manifest generation.
The Node server also accepts previous `PAIR_*` environment settings as fallbacks. When updating
the Compose deployment, rename `PAIR_DOMAIN` in its `.env` file to `SOMMELIER_DOMAIN`.
The server address remains configurable; no operator domain is built in.

Generate the production manifest with the actual deployed HTTPS origin before submitting to
Microsoft. The Office add-in UUID is preserved; its display name, ribbon, help pages and icons
now use Sommelier. See [publishing](publishing.md) and [asset inventory](marketplace-assets.md).
