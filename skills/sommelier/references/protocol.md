# Sommelier Session and RPC 0.2

## Persistent session

Start a remembered bridge with `node scripts/session.mjs start --name desk`. For the first start
only, pass the private setup URL in `SOMMELIER_URL`. Add `--name desk` to each subsequent command;
no credential needs to be copied again. `list` returns names and connection status, never secrets.

The detached process owns one persistent encrypted WebSocket. The URL is passed over stdin and
removed from the child environment. A named profile stores it in an owner-only local credential
file, intentionally, for later reconnection. Unnamed legacy starts keep credentials in memory only.
The control endpoint is a mode 0600 Unix socket in a mode 0700 directory on macOS/Linux; Windows
uses an unguessable named pipe and needs platform-specific ACL validation.

`start` returns the bridge's control path. It proves relay availability, not workbook availability:
check `status.taskPaneConnected`, which becomes true only after end-to-end authentication. Optional
`--session` or `SOMMELIER_SESSION` remain available for scripts using explicit control paths.

The local control protocol is one JSON request line and one JSON response line per connection.

## Commands

```text
session.mjs list                                   list remembered profiles
session.mjs forget --name <name>                   stop and delete a profile
session.mjs start --name <name>                     open the connection, print the session path
session.mjs status                                 report state, queued messages, pending calls
session.mjs next    --timeout <ms>                 wait for the next task-pane message
session.mjs request --method <m> --params <json>   one RPC over the shared connection
session.mjs reply   --content <text>               answer the user in the task pane
session.mjs stop                                   close the connection and clean up
```

## Chat loop

`next` returns the next `pair.chat.message` event whose `data.sender` is `user`:

```json
{ "ok": true, "type": "message", "message": { "messageId": "...", "sender": "user", "content": "...", "occurredAt": "..." } }
```

It returns `{"ok":true,"type":"idle"}` when nothing arrives before the timeout — poll again without
ending the harness turn — and `{"ok":true,"type":"closed"}` once the session ends. Messages that
arrive while no `next` is waiting are queued, bounded to 1000 entries.

`reply` sends the mirror event with `sender: "agent"`, a generated `messageId`, and the current
timestamp. Content is trimmed and bounded to 32000 characters.

`start` reports success only once the relay has accepted the connection, and otherwise prints the
close code and reason the relay gave. `status` reports `closedBecause` for a session that has ended.

`stop` closes the WebSocket and the local server and removes the socket and temporary directory.
Signals do the same. Encrypted connections retry transient relay loss with capped backoff; pending
RPC requests are rejected and never retried. Authentication failures stop reconnecting. A terminal
`stop` retains the credential; `forget` removes it.

## RPC

RPC requests are multiplexed over the same persistent WebSocket:

```sh
node scripts/session.mjs request --method <method> --params '<json-object>' --timeout 60000
```

The result is `{"ok":true,"result":...}`. A protocol `0.2` error envelope is printed as
`{"ok":false,"error":...}` and exits with status 2. Invalid arguments, connection failures, and
timeouts exit non-zero.

`scripts/rpc.mjs --method <method> --params '<json>'` remains available for a single one-shot call.
It opens and closes its own connection, so it cannot receive task-pane messages.

## Observe

```text
excel.context.get       {}
excel.workbook.describe {}
excel.sheet.list        {}
excel.sheet.describe    {"sheetId":"Sheet1"}
excel.range.read        {"range":{"sheetId":"Sheet1","address":"A1:D20"}}
excel.table.list        {"sheetId":"Sheet1"}
excel.table.read        {"tableId":"Sales","offset":0,"limit":100}
```

## Interact

```text
excel.range.select {"range":{"sheetId":"Sheet1","address":"A1:D20"}}
excel.user.ask     {"question":"Which scenario should I use?"}
excel.user.notify  {"message":"Analysis complete","level":"success"}
```

## Mutate

Preview a write:

```json
{
  "method": "excel.range.write",
  "params": {
    "range": { "sheetId": "Sheet1", "address": "B2:B3" },
    "values": [[10], [20]]
  }
}
```

Pass that object as params to `excel.operation.preview`. Its result contains `operationId`. After
user confirmation, call:

```text
excel.operation.commit {"operationId":"<id>"}
excel.operation.undo   {"operationId":"<id>"}
```

Clears use `excel.range.clear` with `range` and `applyTo` equal to `contents`, `formats`, or `all`.
Direct calls to write and clear exist, but agents should use preview and commit.

## Events

The session also carries events from the task pane:

```text
excel.selection.changed  the user changed the workbook selection
excel.operation.changed  a previewed operation was committed, rejected, or undone
excel.audit.recorded     an auditable add-in action completed
pair.peer.changed        the task pane connected or disconnected
pair.chat.message        chat between the task pane and the agent
```

The relay owns `pair.peer.changed` and rejects peers that try to forge it.

## Charts (v1)

List with `excel.chart.list {"sheetId":"Sheet1"}`. Preview `excel.chart.create` with:

```json
{"range":{"sheetId":"Sheet1","address":"A1:B12"},"name":"Monthly sales","title":"Monthly sales","chartType":"column"}
```

Commit the returned operation ID. Supported types: column, bar, line, pie, scatter. Chart names
must be unique per worksheet. The source is bounded to the read limit. Verify with chart.list;
chart undo is not available in v1 (delete it in Excel). The source cells are not changed.

Native adapters and traces are also documented by the `excel_guide` MCP tool and resource
`pair://docs/excel`. They do not require a Context7 account or remote documentation service.
