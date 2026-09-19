import { z } from "zod";

/** Current Sommelier vNext wire protocol version. */
export const PROTOCOL_VERSION = "0.2" as const;

/** A value that can cross a JSON protocol boundary. */
export type JsonValue =
  | null
  | boolean
  | number
  | string
  | JsonValue[]
  | { [key: string]: JsonValue };

/** Portable JSON Schema used by protocol tool consumers. */
export type JsonSchema = Readonly<Record<string, JsonValue>>;

export const jsonValueSchema: z.ZodType<JsonValue> = z.lazy(() =>
  z.union([
    z.null(),
    z.boolean(),
    z.number().finite(),
    z.string(),
    z.array(jsonValueSchema),
    z.record(z.string(), jsonValueSchema),
  ]),
);

const identifierSchema = z.string().trim().min(1).max(200);
const addressSchema = z.string().trim().min(1).max(200);
const timestampSchema = z.iso.datetime({ offset: true });
const emptyParamsSchema = z.strictObject({});
const cellValueSchema = z.union([z.null(), z.boolean(), z.number().finite(), z.string()]);
const matrixSchema = z.array(z.array(cellValueSchema));

/** An explicit range retained by manual context mode. */
export const pinnedRangeSchema = z.strictObject({
  sheetId: identifierSchema,
  address: addressSchema,
});
export type PinnedRange = z.infer<typeof pinnedRangeSchema>;

export const contextModeSchema = z.enum(["follow-selection", "whole-workbook", "manual"]);
export type ContextMode = z.infer<typeof contextModeSchema>;

export const rangeReferenceSchema = pinnedRangeSchema;
export type RangeReference = PinnedRange;

const contextSchema = z.strictObject({
  mode: contextModeSchema,
  pinnedRanges: z.array(pinnedRangeSchema),
});

const workbookSchema = z.strictObject({
  id: identifierSchema,
  name: z.string().min(1),
  activeSheetId: identifierSchema.optional(),
});

const sheetSummarySchema = z.strictObject({
  id: identifierSchema,
  name: z.string().min(1),
  visibility: z.enum(["visible", "hidden", "very-hidden"]).default("visible"),
});

const sheetDescriptionSchema = sheetSummarySchema.extend({
  usedRange: pinnedRangeSchema.optional(),
  rowCount: z.number().int().nonnegative().optional(),
  columnCount: z.number().int().nonnegative().optional(),
});

const rangeDataSchema = z.strictObject({
  range: pinnedRangeSchema,
  values: matrixSchema,
  formulas: z.array(z.array(z.string().nullable())).optional(),
  truncated: z.boolean().default(false),
});

const mutationResultSchema = z.strictObject({
  operationId: identifierSchema.optional(),
  affectedCells: z.number().int().nonnegative(),
});

const tableSummarySchema = z.strictObject({
  id: identifierSchema,
  name: z.string().min(1),
  sheetId: identifierSchema,
  range: addressSchema,
});

const tableDataSchema = z.strictObject({
  table: tableSummarySchema,
  headers: z.array(z.string()),
  rows: matrixSchema,
  offset: z.number().int().nonnegative(),
  hasMore: z.boolean(),
});

const operationStatusSchema = z.enum(["previewed", "committing", "committed", "undone", "failed"]);
export type OperationStatus = z.infer<typeof operationStatusSchema>;

export const operationSchema = z.strictObject({
  operationId: identifierSchema,
  status: operationStatusSchema,
  summary: z.string().optional(),
  affectedRanges: z.array(pinnedRangeSchema).optional(),
});
export type Operation = z.infer<typeof operationSchema>;

const userChoiceSchema = z.strictObject({
  id: identifierSchema,
  label: z.string().min(1),
});

export type ToolCapability = "observe" | "act" | "interact";
export type ToolRisk = "low" | "medium" | "high";

export interface ToolDescriptor<
  ParamsSchema extends z.ZodType = z.ZodType,
  ResultSchema extends z.ZodType = z.ZodType,
> {
  readonly method: string;
  readonly description: string;
  readonly capability: ToolCapability;
  readonly risk: ToolRisk;
  readonly paramsSchema: ParamsSchema;
  readonly resultSchema: ResultSchema;
  readonly inputSchema: JsonSchema;
  readonly outputSchema: JsonSchema;
}

function descriptor<ParamsSchema extends z.ZodType, ResultSchema extends z.ZodType>(options: {
  method: string;
  description: string;
  capability: ToolCapability;
  risk: ToolRisk;
  paramsSchema: ParamsSchema;
  resultSchema: ResultSchema;
}): ToolDescriptor<ParamsSchema, ResultSchema> {
  return {
    ...options,
    inputSchema: z.toJSONSchema(options.paramsSchema, { target: "draft-7" }) as JsonSchema,
    outputSchema: z.toJSONSchema(options.resultSchema, { target: "draft-7" }) as JsonSchema,
  };
}

const rangeParamsSchema = z.strictObject({ range: pinnedRangeSchema });

/** The single registry for method validation, typing, discovery, capability, and risk. */
export const toolDescriptorRegistry = {
  "excel.context.get": descriptor({
    method: "excel.context.get",
    description: "Get the workbook context mode and pinned ranges.",
    capability: "observe",
    risk: "low",
    paramsSchema: emptyParamsSchema,
    resultSchema: contextSchema,
  }),
  "excel.workbook.describe": descriptor({
    method: "excel.workbook.describe",
    description: "Describe workbook identity and active sheet metadata.",
    capability: "observe",
    risk: "low",
    paramsSchema: emptyParamsSchema,
    resultSchema: workbookSchema,
  }),
  "excel.sheet.list": descriptor({
    method: "excel.sheet.list",
    description: "List worksheets without reading cell contents.",
    capability: "observe",
    risk: "low",
    paramsSchema: emptyParamsSchema,
    resultSchema: z.strictObject({ sheets: z.array(sheetSummarySchema) }),
  }),
  "excel.sheet.describe": descriptor({
    method: "excel.sheet.describe",
    description: "Describe one worksheet and its used range.",
    capability: "observe",
    risk: "low",
    paramsSchema: z.strictObject({ sheetId: identifierSchema }),
    resultSchema: sheetDescriptionSchema,
  }),
  "excel.range.read": descriptor({
    method: "excel.range.read",
    description: "Read values and optional formulas from an explicit range.",
    capability: "observe",
    risk: "low",
    paramsSchema: rangeParamsSchema,
    resultSchema: rangeDataSchema,
  }),
  "excel.range.write": descriptor({
    method: "excel.range.write",
    description: "Write a rectangular matrix to an explicit range.",
    capability: "act",
    risk: "high",
    paramsSchema: rangeParamsSchema.extend({ values: matrixSchema }),
    resultSchema: mutationResultSchema,
  }),
  "excel.range.clear": descriptor({
    method: "excel.range.clear",
    description: "Clear values, formats, or all content from an explicit range.",
    capability: "act",
    risk: "high",
    paramsSchema: rangeParamsSchema.extend({
      applyTo: z.enum(["contents", "formats", "all"]).default("contents"),
    }),
    resultSchema: mutationResultSchema,
  }),
  "excel.chart.list": descriptor({
    method: "excel.chart.list",
    description: "List chart IDs, names and titles on a worksheet.",
    capability: "observe",
    risk: "low",
    paramsSchema: z.strictObject({ sheetId: identifierSchema }),
    resultSchema: z.strictObject({
      charts: z.array(
        z.strictObject({ id: identifierSchema, name: identifierSchema, title: z.string() }),
      ),
    }),
  }),
  "excel.chart.create": descriptor({
    method: "excel.chart.create",
    description:
      "Create a named chart from a bounded source range. Preview then commit; requires Excel approval. Use chart.list to verify. Chart undo is not supported in v1.",
    capability: "act",
    risk: "high",
    paramsSchema: rangeParamsSchema.extend({
      name: identifierSchema,
      title: z.string().min(1).max(200),
      chartType: z.enum(["column", "bar", "line", "pie", "scatter"]),
    }),
    resultSchema: mutationResultSchema,
  }),
  "excel.range.select": descriptor({
    method: "excel.range.select",
    description: "Move the user's selection to an explicit range.",
    capability: "interact",
    risk: "low",
    paramsSchema: rangeParamsSchema,
    resultSchema: z.strictObject({ selected: z.literal(true) }),
  }),
  "excel.table.list": descriptor({
    method: "excel.table.list",
    description: "List workbook tables, optionally restricted to a sheet.",
    capability: "observe",
    risk: "low",
    paramsSchema: z.strictObject({ sheetId: identifierSchema.optional() }),
    resultSchema: z.strictObject({ tables: z.array(tableSummarySchema) }),
  }),
  "excel.table.read": descriptor({
    method: "excel.table.read",
    description: "Read a bounded page of rows from a table.",
    capability: "observe",
    risk: "low",
    paramsSchema: z.strictObject({
      tableId: identifierSchema,
      offset: z.number().int().nonnegative().default(0),
      limit: z.number().int().positive().max(1_000).default(100),
    }),
    resultSchema: tableDataSchema,
  }),
  "excel.operation.preview": descriptor({
    method: "excel.operation.preview",
    description: "Preview an action before requesting its commit.",
    capability: "act",
    risk: "medium",
    paramsSchema: z.strictObject({
      method: z.enum(["excel.range.write", "excel.range.clear", "excel.chart.create"]),
      params: jsonValueSchema,
    }),
    resultSchema: operationSchema.extend({ status: z.literal("previewed") }),
  }),
  "excel.operation.commit": descriptor({
    method: "excel.operation.commit",
    description: "Commit a previously previewed operation.",
    capability: "act",
    risk: "high",
    paramsSchema: z.strictObject({ operationId: identifierSchema }),
    resultSchema: operationSchema.extend({ status: z.literal("committed") }),
  }),
  "excel.operation.undo": descriptor({
    method: "excel.operation.undo",
    description: "Undo a committed operation when the host still supports it.",
    capability: "act",
    risk: "medium",
    paramsSchema: z.strictObject({ operationId: identifierSchema }),
    resultSchema: operationSchema.extend({ status: z.literal("undone") }),
  }),
  "excel.user.ask": descriptor({
    method: "excel.user.ask",
    description: "Ask the user for text or a choice.",
    capability: "interact",
    risk: "low",
    paramsSchema: z.strictObject({
      question: z.string().min(1),
      choices: z.array(userChoiceSchema).min(1).optional(),
    }),
    resultSchema: z.strictObject({ answer: z.string(), choiceId: identifierSchema.optional() }),
  }),
  "excel.user.notify": descriptor({
    method: "excel.user.notify",
    description: "Show an informational message to the user.",
    capability: "interact",
    risk: "low",
    paramsSchema: z.strictObject({
      message: z.string().min(1),
      level: z.enum(["info", "success", "warning", "error"]).default("info"),
    }),
    resultSchema: z.strictObject({ acknowledged: z.boolean() }),
  }),
} as const satisfies Record<string, ToolDescriptor>;

/** Alias emphasizing that method schemas are held by the descriptor registry. */
export const methodSchemas = toolDescriptorRegistry;
export type ProtocolMethod = keyof typeof toolDescriptorRegistry;

export const METHOD_NAMES = Object.freeze(
  Object.keys(toolDescriptorRegistry) as ProtocolMethod[],
) as readonly ProtocolMethod[];

export type ProtocolMethodMap = {
  [Method in ProtocolMethod]: {
    readonly params: z.infer<(typeof toolDescriptorRegistry)[Method]["paramsSchema"]>;
    readonly result: z.infer<(typeof toolDescriptorRegistry)[Method]["resultSchema"]>;
  };
};

export type ProtocolRequest<Method extends ProtocolMethod = ProtocolMethod> =
  Method extends ProtocolMethod
    ? {
        readonly protocolVersion: typeof PROTOCOL_VERSION;
        readonly type: "request";
        readonly id: string;
        readonly method: Method;
        readonly params: ProtocolMethodMap[Method]["params"];
      }
    : never;

export type ProtocolResponse<Method extends ProtocolMethod = ProtocolMethod> =
  Method extends ProtocolMethod
    ? {
        readonly protocolVersion: typeof PROTOCOL_VERSION;
        readonly type: "response";
        readonly id: string;
        readonly method: Method;
        readonly result: ProtocolMethodMap[Method]["result"];
      }
    : never;

const requestBaseSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("request"),
  id: identifierSchema,
  method: z.enum(METHOD_NAMES),
  params: z.unknown(),
});

export const protocolRequestSchema = requestBaseSchema.superRefine((request, context) => {
  const result = toolDescriptorRegistry[request.method].paramsSchema.safeParse(request.params);
  if (!result.success) {
    for (const issue of result.error.issues) {
      context.addIssue({ ...issue, path: ["params", ...issue.path] });
    }
  }
}) as z.ZodType<ProtocolRequest>;

const responseBaseSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("response"),
  id: identifierSchema,
  method: z.enum(METHOD_NAMES),
  result: z.unknown(),
});

export const protocolResponseSchema = responseBaseSchema.superRefine((response, context) => {
  const result = toolDescriptorRegistry[response.method].resultSchema.safeParse(response.result);
  if (!result.success) {
    for (const issue of result.error.issues) {
      context.addIssue({ ...issue, path: ["result", ...issue.path] });
    }
  }
}) as z.ZodType<ProtocolResponse>;

export const protocolErrorObjectSchema = z.strictObject({
  code: z.string().min(1),
  message: z.string().min(1),
  data: jsonValueSchema.optional(),
});
export type ProtocolErrorObject = z.infer<typeof protocolErrorObjectSchema>;

export const protocolErrorSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("error"),
  id: z.union([identifierSchema, z.null()]).optional(),
  error: protocolErrorObjectSchema,
});
export type ProtocolError = z.infer<typeof protocolErrorSchema>;

export const selectionChangedEventSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("event"),
  event: z.literal("excel.selection.changed"),
  data: z.strictObject({ range: pinnedRangeSchema.nullable() }),
});
export type SelectionChangedEvent = z.infer<typeof selectionChangedEventSchema>;

export const operationEventSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("event"),
  event: z.literal("excel.operation.changed"),
  data: z.strictObject({
    operationId: identifierSchema,
    status: operationStatusSchema,
    undoable: z.boolean().optional(),
    summary: z.string().optional(),
    occurredAt: timestampSchema,
    error: protocolErrorObjectSchema.optional(),
  }),
});
export type OperationEvent = z.infer<typeof operationEventSchema>;

export const auditEventSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("event"),
  event: z.literal("excel.audit.recorded"),
  data: z.strictObject({
    auditId: identifierSchema,
    action: z.string().min(1),
    outcome: z.enum(["success", "rejected", "failure"]),
    occurredAt: timestampSchema,
    operationId: identifierSchema.optional(),
    details: jsonValueSchema.optional(),
  }),
});
export type AuditEvent = z.infer<typeof auditEventSchema>;

export const pairPeerChangedEventSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("event"),
  event: z.literal("pair.peer.changed"),
  data: z.strictObject({
    role: z.enum(["addin", "agent"]),
    connected: z.boolean(),
  }),
});
export type PairPeerChangedEvent = z.infer<typeof pairPeerChangedEventSchema>;

export const pairChatMessageEventSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  type: z.literal("event"),
  event: z.literal("pair.chat.message"),
  data: z.strictObject({
    messageId: identifierSchema,
    sender: z.enum(["user", "agent"]),
    content: z.string().trim().min(1).max(32_000),
    occurredAt: timestampSchema,
  }),
});
export type PairChatMessageEvent = z.infer<typeof pairChatMessageEventSchema>;

export const protocolEventSchema = z.discriminatedUnion("event", [
  selectionChangedEventSchema,
  operationEventSchema,
  auditEventSchema,
  pairPeerChangedEventSchema,
  pairChatMessageEventSchema,
]);
export type ProtocolEvent = z.infer<typeof protocolEventSchema>;

export const protocolMessageSchema = z.union([
  protocolRequestSchema,
  protocolResponseSchema,
  protocolEventSchema,
  protocolErrorSchema,
]);
export type ProtocolMessage = ProtocolRequest | ProtocolResponse | ProtocolEvent | ProtocolError;

/** Parse and validate an untrusted vNext protocol message. */
export function parseProtocolMessage(input: unknown): ProtocolMessage {
  return protocolMessageSchema.parse(input);
}
