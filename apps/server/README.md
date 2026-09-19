# Sommelier Server

Production static host and WebSocket rendezvous relay for the free Sommelier add-in.

Required environment:

- `PAIR_PUBLIC_ORIGIN`: public HTTPS origin, for example `https://sommelier.example.com`.
- `PAIR_ADDIN_ORIGIN`: optional trusted task pane HTTPS origin when its assets are hosted separately.
  Only that browser origin is permitted for WebSocket pairing and config CORS.
- `PAIR_BRIDGE_SCRIPT`: optional path to the agent session bridge served at `/agent/session.mjs`.
  Defaults to `skills/sommelier/scripts/session.mjs` relative to the working directory.
- `PAIR_AGENT_ORIGIN`: optional agent-facing WS/WSS origin. Local development uses
  `ws://127.0.0.1:3001` to avoid the Vite development certificate; production defaults to the public
  WSS origin.
- `PORT`: internal HTTP port; defaults to `3000`.
- `HOST`: bind address; defaults to `127.0.0.1` outside the container.
- `PAIR_STATIC_DIR`: task pane build directory; defaults to `apps/addin/dist`.

The server exposes `GET /healthz`, `GET /api/config`, static help/assets, the portable bridge and
WebSocket `/connect`. The in-memory registry holds only opaque UIDs, routing capabilities and live
socket references. It is bounded to 1,000 entries; inactive entries expire after ten minutes.
Endpoint-owned credentials let sessions re-register after a relay restart. Ping/pong reaps dead peers.

Only encrypted transport frames cross `/connect`. The server does not hold end-to-end secrets,
execute workbook RPC, proxy direct BYOA endpoints, or store model keys. See
[the security architecture](../../docs/secure-pairing.md) for metadata visibility and client trust.
Run `pnpm bridge:build` before starting from source. The bridge needs its adjacent generated
`lib/encrypted-socket.mjs`; Docker copies the entire script directory.

Legacy plaintext routes are disabled by default. `allowUnencryptedLegacy` is an explicit library
migration/testing option, unavailable through the production entry point.
