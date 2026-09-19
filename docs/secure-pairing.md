# Remembered terminals and end-to-end transport

## Implemented behavior

A user associates an Excel task pane with a named local bridge once. The pane creates a random UID
and a separate 256-bit secret using Web Crypto. The one-time setup transfers them to the local
bridge. Both peers prove possession of that secret before the pane marks the terminal connected
or saves it as authorized. A UID alone is not authorization.

The pane stores an allowlisted record: UID, secret, local display name, agent-facing relay origin,
and an optional automatic reconnection preference. Up to 20 named terminals can be remembered.
There is no public endpoint that enumerates other users' terminals. Saved entries are shown as
**Saved terminal**, not as online until authenticated. Only one automatic connection is selected.

The bridge stores named profiles under `~/.sommelier` (`SOMMELIER_HOME` can override the path).
Directories use mode 0700 and atomic profile files mode 0600 on macOS/Linux. `start --name desk`
reuses a running bridge or restarts it with its saved credential. `list` never returns credentials.
`stop` retains the association; `forget` deletes it. The Unix control socket also uses owner-only
permissions. Windows named pipes and filesystem ACLs need separate platform validation.

One UID connects one pane and one bridge at a time. Multiple agent sessions should use different
names and associations. A second pane cannot take over the occupied role. A saved label identifies
a bridge profile; the system does not claim to attach to arbitrary pre-existing terminal processes.

Reopening a pane requires **Reconnect**, unless the user selected automatic connection. Network
loss closes the pane's active controller and clears pending approval. Reconnecting creates a new
controller, chat and undo history; completed workbook writes remain in the workbook. The agent must
inspect workbook state after an interrupted operation, never blindly retry a write. The bridge
reconnects its relay socket with capped backoff; it never retries application requests.

## Transport decision

| Option | Benefits | Cost / limitation | Decision |
| --- | --- | --- | --- |
| WebRTC data channel | Peer-to-peer when ICE succeeds; encrypted DTLS transport | Node needs a WebRTC implementation, signaling and TURN fallback; Office WebViews need a real compatibility matrix; authenticate exchanged fingerprints to prevent relay MITM | Evaluate as a later transport after platform tests |
| WSS plus end-to-end encryption | Native WebSocket + Web Crypto in browsers and Node 22+; one outbound connection; works with existing shell bridges | Relay bandwidth is used for every message; metadata remains visible | Implemented |
| Periodic HTTP polling | Easy request/response infrastructure | Repeated requests while idle, latency and more state for delivery/order | Not used |
| On-demand socket | No idle socket while pane is closed | Requires explicit reconnection; cannot wake a sleeping harness by itself | Pane default; optional connect-on-open |

The bridge keeps a socket while the skill is active. The relay uses ping/pong every 30 seconds to
reap dead sockets. Its registry is bounded (1000 entries by default), retains empty entries for ten
minutes and stores no workbook data. After a server restart the endpoints re-register on demand;
no server database or persistent store of terminal names is necessary. A public deployment should
apply network-level abuse controls appropriate to its traffic; capacity limits are not user quotas.

## Wire layers

```text
Task pane + workbook controller                 Local bridge + chosen harness
           RPC 0.2  <---- authenticated encrypted channel ---->  RPC 0.2
                    \---- opaque WSS frames via relay ----/
```

`/connect?uid=...&role=addin|agent&token=...` is the rendezvous endpoint. The relay capability is
HMAC-SHA256(secret, `ai-cdl-relay-v1:<uid>`). It is domain-separated from end-to-end keys. It allows
routing but cannot decrypt content or authenticate the peer. The bootstrap URL carries the secret
in its fragment; `relaySocketUrl` removes that fragment before WebSocket construction.

Encrypted frame version **1** and workbook RPC version **0.2** are separate. The relay validates
only `hello`, `auth`, and `data` frame shapes, bounded sizes and role occupancy. It neither parses
nor executes Excel RPC. Plaintext RPC is rejected on `/connect`.

For each pairing/reconnection both endpoints create fresh random 32-byte challenges. The transcript
binds the encryption protocol version, UID, add-in challenge and agent challenge in a fixed order.
Each peer proves the transcript and its role with HMAC-SHA256; verification uses Web Crypto.
HKDF-SHA256 derives separate AES-256-GCM keys for each direction from the shared secret and fresh
transcript. AES-GCM authenticates the transcript and sender role as additional data. Each direction
uses a monotonically increasing sequence number as its 96-bit IV. A fresh key is derived for every
connection; out-of-order or repeated sequence numbers are rejected. The tag is 128 bits.

Messages are processed serially to preserve ordering around asynchronous Web Crypto calls. There
is no offline workbook-message queue in the encrypted socket. Authentication times out after 15
seconds; bad proof, malformed frames and authentication failures close the connection without a
plaintext fallback. The server's presence events are only hints to start or cancel the handshake;
they cannot mark a terminal authenticated.

The portable skill's encryption module is generated from `packages/transport/src` by
`pnpm bridge:build`. It is included by the skill installer and Docker image, so browser and bridge
share one maintained implementation. It must be rebuilt after transport changes.

## Confidentiality boundary

The relay sees UID, a routing capability, roles, source IPs, connection timing and frame lengths.
It can interrupt, withhold or reorder traffic; this causes failure, not successful decryption.
Endpoint code decrypts cells, prompts, agent answers and tool results. The chosen harness controls
any subsequent transfer to its model provider. Sommelier does not call that provider on the user's behalf.

The task pane and bridge code must be trusted. End-to-end encryption cannot protect against an
operator that also changes the client JavaScript to steal secrets. For protection from a malicious
relay operator, host the pane independently under a trusted origin and install a reviewed bridge
from a trusted distribution. The supplied deployment serves both from the same application host,
so it protects against passive relay inspection, not a malicious replacement of client assets.
For that deployment, set `SOMMELIER_ADDIN_ORIGIN=https://trusted-pane.example` on the relay and build the
pane with `VITE_PAIR_RELAY_ORIGIN=https://relay.example`. The relay permits the configured pane
origin for WebSocket connections and config CORS. Only the trusted pane operator should control
client assets; installing a trusted bridge does not make a maliciously served pane safe.

The first setup includes a secret. The native `sommelier pair` command accepts it at a
local terminal prompt and stores it without submitting it to a model. The skill-only fallback
still exposes the setup to a model if the user pastes it into an agent chat; do not claim that
fallback hides the initial credential from the provider/history.

This PSK design does **not** provide forward secrecy: compromise of the saved secret can decrypt
captured sessions. Browser same-origin storage and local profile files are not hardware key stores.
Forget the profile at both ends to remove locally remembered trust; other running copies of an
already-paired credential must also be closed. Changing a terminal name does not rotate its secret.
The protocol has focused negative and integration tests, not an independent cryptographic audit.

## Compatibility

OpenCode, Codex and Claude Code have [native adapters](harness-adapters.md) and an Agent Skill fallback. The installer also
supports Pi, verified against its current official skill documentation and a local Pi executable.
The note “MyPy” remains unconfirmed; it is not treated as a supported harness name.

A skill can run the bridge and consume messages using shell tools. A bridge alone cannot inject
messages into or wake an arbitrary existing model session. The fallback depends on the harness
continuing the skill's `next` loop. Native adapters now deliver turns through Codex App Server,
OpenCode's loopback API or Claude's explicitly enabled MCP channel. They use the chosen harness's
existing model credentials. The workbook RPC and encrypted relay stay independent of the harness.

The old hosted `/api/pair/sessions` and `/pair` plaintext paths are disabled by default. The explicit
`allowUnencryptedLegacy` library option exists only for migration tests or trusted local development;
it is not exposed by the production entry point. The legacy loopback CLI and direct BYOA HTTP mode
remain separate APIs and are not described as end-to-end encrypted relay transport.

## Sources checked

- [OpenCode skill discovery](https://opencode.ai/docs/skills/)
- [Codex local skill discovery](https://learn.chatgpt.com/docs/build-skills)
- [Claude Code skills](https://code.claude.com/docs/en/skills)
- [Pi skills](https://github.com/earendil-works/pi/blob/main/packages/coding-agent/docs/skills.md)
- [WebRTC connectivity and ICE](https://developer.mozilla.org/en-US/docs/Web/API/WebRTC_API/Connectivity)
