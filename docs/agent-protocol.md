# Agent protocol

## Version 0.1

Every request, event, result, and trace envelope declares `protocolVersion: "0.1"`. Unknown
versions are rejected. A future incompatible wire format will change the major component.

## HTTP contract

The adapter sends `POST` JSON with header `x-ai-cdl-protocol-version: 0.1`:

```json
{
  "protocolVersion": "0.1",
  "sessionId": "session-id",
  "turnId": "turn-id",
  "iteration": 1,
  "userMessage": { "protocolVersion": "0.1", "type": "user_message", "text": "Read A1:B2", "selectionIncluded": false },
  "tools": [],
  "toolResults": []
}
```

Return status 2xx, `content-type: application/json`, and:

```json
{
  "protocolVersion": "0.1",
  "events": [
    { "protocolVersion": "0.1", "type": "tool_call", "id": "call_123", "name": "excel.read_range", "arguments": { "worksheet": "Forecast", "range": "A1:F40" } }
  ]
}
```

On the next iteration, the original user message remains distinct and `toolResults` includes the
validated result. End with an `assistant_message` and `completion`. The MVP uses JSON responses;
SSE/NDJSON is not implemented. Honor aborted HTTP requests. Never put bearer tokens in response
bodies, logs, or trace fields.

## Error behavior

Non-2xx responses retain only status and a safe request ID. Invalid JSON, schema mismatches,
network failures, timeout, and cancellation use distinct stable error codes. Use
`sommelier-contract --endpoint <url> --timeout-ms 60000` before integration.
