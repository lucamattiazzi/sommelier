# @lucamattiazzi/sommelier-transport

Transport-neutral lifecycle and validated protocol message delivery for Sommelier.

## End-to-end encrypted Sommelier sockets

`createPairIdentity()` creates a random UID and a separate 256-bit secret. `pairConnectionUrl(origin,
identity, role)` derives a routing capability and keeps the encryption secret in the URL fragment.
Use `relaySocketUrl(url)` for the actual WebSocket, then wrap it with
`createEncryptedSocket(socket, url)` and `createJsonSocketTransport`.

The outer socket can be open before the peer is authenticated. Subscribe to `pair.peer.changed`:
this wrapper only emits `connected: true` after mutual authentication, unlike raw relay presence.
Application sends while no authenticated peer exists are discarded; there is no offline queue.
Callers must await authenticated presence before issuing requests. Async encryption is serialized.

`createSecureChannel` and `parseSecureFrame` expose the lower framing layer without workbook
semantics. Direct channel users must serialize `accept`/`send` calls; the socket wrapper does so.
See [the protocol and trust boundaries](../../docs/secure-pairing.md). A shared-secret protocol is
not forward-secret, and client code must come from a trusted distribution.
