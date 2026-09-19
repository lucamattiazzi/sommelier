# @lucamattiazzi/sommelier-protocol

Provider-independent TypeScript types and Zod runtime schemas for Sommelier protocol `0.2`.
The package describes Excel observation, action, and interaction methods without depending on
Office.js, a model provider, browser APIs, or `@lucamattiazzi/sommelier-core`.

```ts
import { parseProtocolMessage, PROTOCOL_VERSION } from "@lucamattiazzi/sommelier-protocol";

const message = parseProtocolMessage({
  protocolVersion: PROTOCOL_VERSION,
  type: "request",
  id: "request-1",
  method: "excel.sheet.list",
  params: {},
});
```
