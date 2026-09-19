# Sommelier Add-in

Free Excel task pane for bring-your-own-agent usage. Development preview.

Start with the [manual test guide](../../docs/pair-testing.md).

## Local development

Run the relay/static server and Vite task pane in separate terminals:

```sh
pnpm server:dev
pnpm addin:dev
```

Sideload `apps/addin/manifest.xml` or use `https://localhost:3000/manifest.xml` where the Office
developer tooling accepts a manifest URL. Vite proxies `/api`, `/agent`, and `/connect` to the development server
on port 3001.

## Direct agents

The configured endpoint must implement the Sommelier agent protocol `0.1`. Remote endpoints must use
HTTPS and allow the task pane origin through CORS. If bearer authentication is used, allow the
`Authorization`, `Content-Type`, `Accept`, and `X-AI-CDL-Protocol-Version` request headers. Sommelier keeps
the token only in React memory; it is not written to browser or Office storage.

## Relay agents

The pane remembers named terminals after mutual end-to-end authentication. Install the portable
skill with `pnpm bridge:build` and `pnpm harness:install opencode` (or `codex`, `claude`, `pi`).
Copy its setup only for the first association. Later, start the named bridge from the skill and
click **Reconnect**, or opt into automatic connection when opening the pane.

Chat and workbook RPC 0.2 are encrypted between pane and bridge. The relay only validates public
transport frames. The agent must keep consuming the skill's chat loop; no background model process
is created by Sommelier. See [security and transport decisions](../../docs/secure-pairing.md).
