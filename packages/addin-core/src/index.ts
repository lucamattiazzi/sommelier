import {
  type CellValue,
  type ExcelAdapter,
  type ExcelAdapterUnsubscribe,
  type ExcelSelection,
  parseA1Range,
  type ReadRangeResult,
  type SheetDescription,
} from "@lucamattiazzi/sommelier-excel";
import {
  type AuditEvent,
  type ContextMode,
  type JsonValue,
  type Operation,
  type PinnedRange,
  PROTOCOL_VERSION,
  type ProtocolError,
  type ProtocolErrorObject,
  type ProtocolEvent,
  type ProtocolMethodMap,
  type ProtocolRequest,
  type ProtocolResponse,
  toolDescriptorRegistry,
} from "@lucamattiazzi/sommelier-protocol";

export interface WorkbookIdentity {
  readonly id: string;
  readonly name: string;
}

export interface ApprovalRequest {
  readonly method:
    | "excel.range.write"
    | "excel.range.clear"
    | "excel.chart.create"
    | "excel.operation.undo";
  readonly summary: string;
  readonly ranges: readonly PinnedRange[];
  readonly operationId?: string;
  readonly affectedCells?: number;
  readonly before?: readonly (readonly CellValue[])[];
  readonly after?: readonly (readonly CellValue[])[];
}

export interface AskUserRequest {
  readonly question: string;
  readonly choices?: readonly { readonly id: string; readonly label: string }[];
}

export interface AskUserResult {
  readonly answer: string;
  readonly choiceId?: string;
}

export interface NotifyUserRequest {
  readonly message: string;
  readonly level: "info" | "success" | "warning" | "error";
}

export interface UserInteraction {
  readonly ask?: (request: AskUserRequest) => AskUserResult | Promise<AskUserResult>;
  readonly notify?: (request: NotifyUserRequest) => boolean | Promise<boolean>;
}

export type ControllerIdKind = "operation" | "audit";

export interface AddinControllerOptions {
  readonly adapter: ExcelAdapter;
  readonly workbook: WorkbookIdentity;
  readonly maxCellsPerRead?: number;
  readonly maxCellsPerWrite?: number;
  readonly maxOperations?: number;
  readonly contextMode?: ContextMode;
  readonly pinnedRanges?: readonly PinnedRange[];
  readonly requestApproval?: (request: ApprovalRequest) => boolean | Promise<boolean>;
  readonly userInteraction?: UserInteraction;
  readonly auditSink?: (event: AuditEvent) => void | Promise<void>;
  readonly idFactory?: (kind: ControllerIdKind) => string;
  readonly timeFactory?: () => Date;
}

export type ProtocolEventListener = (event: ProtocolEvent) => void;
export type Unsubscribe = () => void;

export interface AddinController {
  start(): Promise<void>;
  stop(): Promise<void>;
  handle(request: ProtocolRequest): Promise<ProtocolResponse | ProtocolError>;
  subscribe(listener: ProtocolEventListener): Unsubscribe;
  setContext(mode: ContextMode, ranges?: readonly PinnedRange[]): void;
}

type MutationMethod = "excel.range.write" | "excel.range.clear" | "excel.chart.create";
type WriteParams = ProtocolMethodMap["excel.range.write"]["params"];
type ClearParams = ProtocolMethodMap["excel.range.clear"]["params"];
type ChartParams = ProtocolMethodMap["excel.chart.create"]["params"];
type MutationParams = WriteParams | ClearParams | ChartParams;

interface StoredOperation {
  readonly operationId: string;
  readonly method: MutationMethod;
  readonly params: MutationParams;
  readonly summary: string;
  readonly affectedRanges: readonly PinnedRange[];
  status: Operation["status"];
  before?: ReadRangeResult;
  after?: ReadRangeResult;
}

class ControllerFailure extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly outcome: "rejected" | "failure" = "failure",
  ) {
    super(message);
    this.name = "ControllerFailure";
  }
}

function defaultIdFactory(kind: ControllerIdKind): string {
  return `${kind}-${crypto.randomUUID()}`;
}

function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(String(value));
}

function visibility(description: SheetDescription): "visible" | "hidden" | "very-hidden" {
  return description.visibility === "veryHidden" ? "very-hidden" : description.visibility;
}

/** Create a host-independent controller over an ExcelAdapter. */
export function createAddinController(options: AddinControllerOptions): AddinController {
  const listeners = new Set<ProtocolEventListener>();
  const operations = new Map<string, StoredOperation>();
  let contextMode = options.contextMode ?? "follow-selection";
  let manualRanges = [...(options.pinnedRanges ?? [])];
  const idFactory = options.idFactory ?? defaultIdFactory;
  const timeFactory = options.timeFactory ?? (() => new Date());
  let currentSelection: PinnedRange | null = null;
  let adapterUnsubscribe: ExcelAdapterUnsubscribe | undefined;
  let started = false;
  let stopped = false;
  let generation = 0;
  let actionBusy = false;
  const maxRead = options.maxCellsPerRead ?? 10_000;
  const maxWrite = options.maxCellsPerWrite ?? 1_000;
  const maxOperations = options.maxOperations ?? 100;
  for (const limit of [maxRead, maxWrite, maxOperations]) {
    if (!Number.isSafeInteger(limit) || limit < 1)
      throw new Error("Limits must be positive safe integers.");
  }

  function checkRange(range: PinnedRange, limit: number): number {
    const count = parseA1Range(range.address).cellCount;
    if (count > limit)
      throw new ControllerFailure(
        "RANGE_LIMIT_EXCEEDED",
        `Range contains ${count} cells; limit is ${limit}. Request a smaller range.`,
      );
    return count;
  }

  async function unchanged(snapshot: ReadRangeResult): Promise<void> {
    const checkGeneration = generation;
    const current = await options.adapter.readRange({
      worksheet: snapshot.worksheet,
      range: snapshot.range,
    });
    if (stopped || generation !== checkGeneration)
      throw new ControllerFailure(
        "SESSION_STOPPED",
        "This Sommelier session has ended.",
        "rejected",
      );
    if (
      JSON.stringify(current.values) !== JSON.stringify(snapshot.values) ||
      JSON.stringify(current.formulas) !== JSON.stringify(snapshot.formulas)
    ) {
      throw new ControllerFailure(
        "WORKBOOK_CONFLICT",
        "The target changed. Read it again and create a new preview.",
        "rejected",
      );
    }
  }

  function approvalPreview(
    method: MutationMethod,
    params: MutationParams,
    before: ReadRangeResult,
  ) {
    if (method === "excel.chart.create") return { affectedCells: 0 };
    return {
      affectedCells: before.rowCount * before.columnCount,
      before: before.values.map((row, r) =>
        row.map((value, c) => before.formulas?.[r]?.[c] ?? value),
      ),
      after:
        method === "excel.range.write"
          ? (params as WriteParams).values
          : before.values.map((row) => row.map(() => null)),
    };
  }

  function emit(event: ProtocolEvent): void {
    for (const listener of listeners) {
      try {
        listener(event);
      } catch {
        // One host listener must not prevent protocol handling or other listeners.
      }
    }
  }

  function selectionRange(selection: ExcelSelection): PinnedRange {
    return { sheetId: selection.worksheet, address: selection.range };
  }

  function emitOperation(
    operation: StoredOperation,
    status: Operation["status"],
    error?: ProtocolErrorObject,
  ): void {
    operation.status = status;
    emit({
      protocolVersion: PROTOCOL_VERSION,
      type: "event",
      event: "excel.operation.changed",
      data: {
        operationId: operation.operationId,
        status,
        undoable: operation.method !== "excel.chart.create",
        summary: operation.summary,
        occurredAt: timeFactory().toISOString(),
        ...(error ? { error } : {}),
      },
    });
  }

  async function audit(
    action: string,
    outcome: "success" | "rejected" | "failure",
    operationId?: string,
    details?: JsonValue,
  ): Promise<void> {
    const event: AuditEvent = {
      protocolVersion: PROTOCOL_VERSION,
      type: "event",
      event: "excel.audit.recorded",
      data: {
        auditId: idFactory("audit"),
        action,
        outcome,
        occurredAt: timeFactory().toISOString(),
        ...(operationId ? { operationId } : {}),
        ...(details !== undefined ? { details } : {}),
      },
    };
    emit(event);
    try {
      await options.auditSink?.(event);
    } catch {
      // Audit transport failures do not change an already completed workbook action.
    }
  }

  function protocolError(id: string | undefined, failure: unknown): ProtocolError {
    const error = asError(failure);
    const code = failure instanceof ControllerFailure ? failure.code : "ADAPTER_ERROR";
    return {
      protocolVersion: PROTOCOL_VERSION,
      type: "error",
      ...(id ? { id } : {}),
      error: { code, message: error.message || "The request failed." },
    };
  }

  async function ensureValidMutation(
    method: MutationMethod,
    params: MutationParams,
  ): Promise<ReadRangeResult> {
    checkRange(
      params.range,
      method === "excel.chart.create" ? maxRead : Math.min(maxRead, maxWrite),
    );
    if (method === "excel.chart.create") {
      if (!options.adapter.createChart || !options.adapter.listCharts)
        throw new ControllerFailure("UNSUPPORTED_OPERATION", "This adapter cannot create charts.");
      const charts = await options.adapter.listCharts(params.range.sheetId);
      if (charts.some((chart) => chart.name === (params as ChartParams).name))
        throw new ControllerFailure(
          "CHART_EXISTS",
          "Chart name already exists. Choose another name.",
        );
    }
    if (method === "excel.range.clear" && (params as ClearParams).applyTo !== "contents") {
      throw new ControllerFailure(
        "UNSUPPORTED_OPERATION",
        "Sommelier currently supports clear(contents) only. No cells were changed.",
      );
    }
    const target = await options.adapter.readRange({
      worksheet: params.range.sheetId,
      range: params.range.address,
    });
    if (method === "excel.range.write") {
      const write = params as WriteParams;
      const rectangular = write.values.every((row) => row.length === target.columnCount);
      if (write.values.length !== target.rowCount || !rectangular) {
        throw new ControllerFailure(
          "INVALID_PARAMS",
          `values must have shape ${target.rowCount}x${target.columnCount}.`,
        );
      }
    }
    return target;
  }

  async function approve(request: ApprovalRequest): Promise<void> {
    if (stopped)
      throw new ControllerFailure(
        "SESSION_STOPPED",
        "This Sommelier session has ended.",
        "rejected",
      );
    const approvalGeneration = generation;
    if (!(await options.requestApproval?.(request))) {
      throw new ControllerFailure(
        "APPROVAL_REQUIRED",
        "The host must explicitly approve this workbook mutation.",
        "rejected",
      );
    }
    if (stopped || generation !== approvalGeneration)
      throw new ControllerFailure(
        "SESSION_STOPPED",
        "This session ended before approval completed.",
        "rejected",
      );
  }

  function mutationSummary(method: MutationMethod, params: MutationParams): string {
    if (method === "excel.chart.create")
      return `Create ${(params as ChartParams).chartType} chart "${(params as ChartParams).name}" from ${params.range.sheetId}!${params.range.address}`;
    const action = method === "excel.range.write" ? "Write" : "Clear";
    return `${action} ${params.range.sheetId}!${params.range.address}`;
  }

  async function executeMutation(method: MutationMethod, params: MutationParams): Promise<number> {
    if (method === "excel.chart.create") {
      if (!options.adapter.createChart)
        throw new ControllerFailure("UNSUPPORTED_OPERATION", "This adapter cannot create charts.");
      const chart = params as ChartParams;
      await options.adapter.createChart({
        worksheet: chart.range.sheetId,
        range: chart.range.address,
        name: chart.name,
        title: chart.title,
        chartType: chart.chartType,
      });
      return 0;
    }
    if (method === "excel.range.write") {
      return (
        await options.adapter.writeRange({
          worksheet: params.range.sheetId,
          range: params.range.address,
          values: (params as WriteParams).values,
        })
      ).cellCount;
    }
    return (
      await options.adapter.clearRange({
        worksheet: params.range.sheetId,
        range: params.range.address,
      })
    ).cellCount;
  }

  async function directMutation(method: MutationMethod, params: MutationParams) {
    const before = await ensureValidMutation(method, params);
    await approve({
      method,
      summary: mutationSummary(method, params),
      ranges: [params.range],
      ...approvalPreview(method, params, before),
    });
    await unchanged(before);
    return { affectedCells: await executeMutation(method, params) };
  }

  function publicOperation(operation: StoredOperation): Operation {
    return {
      operationId: operation.operationId,
      status: operation.status,
      summary: operation.summary,
      affectedRanges: [...operation.affectedRanges],
    };
  }

  async function dispatch(request: ProtocolRequest): Promise<unknown> {
    switch (request.method) {
      case "excel.context.get": {
        toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        if (contextMode === "manual") return { mode: contextMode, pinnedRanges: manualRanges };
        if (contextMode === "whole-workbook") return { mode: contextMode, pinnedRanges: [] };
        const context = await options.adapter.getContext();
        currentSelection = selectionRange(context.selection);
        return { mode: contextMode, pinnedRanges: [currentSelection] };
      }
      case "excel.workbook.describe": {
        toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const context = await options.adapter.getContext();
        return { ...options.workbook, activeSheetId: context.activeWorksheet };
      }
      case "excel.sheet.list": {
        toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const sheets = await options.adapter.listSheets();
        return {
          sheets: sheets.map((sheet) => ({
            id: sheet.name,
            name: sheet.name,
            visibility: visibility(sheet),
          })),
        };
      }
      case "excel.sheet.describe": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const sheet = await options.adapter.describeSheet({ worksheet: params.sheetId });
        return {
          id: sheet.name,
          name: sheet.name,
          visibility: visibility(sheet),
          ...(sheet.usedRange
            ? { usedRange: { sheetId: sheet.name, address: sheet.usedRange } }
            : {}),
        };
      }
      case "excel.range.read": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        checkRange(params.range, maxRead);
        const result = await options.adapter.readRange({
          worksheet: params.range.sheetId,
          range: params.range.address,
        });
        return {
          range: { sheetId: result.worksheet, address: result.range },
          values: result.values,
          ...(result.formulas ? { formulas: result.formulas } : {}),
          truncated: false,
        };
      }
      case "excel.chart.list": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        if (!options.adapter.listCharts)
          throw new ControllerFailure("UNSUPPORTED_OPERATION", "This adapter cannot list charts.");
        return { charts: await options.adapter.listCharts(params.sheetId) };
      }
      case "excel.chart.create":
      case "excel.range.write": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        return directMutation(request.method, params);
      }
      case "excel.range.clear": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        return directMutation(request.method, params);
      }
      case "excel.range.select": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        await options.adapter.selectRange({
          worksheet: params.range.sheetId,
          range: params.range.address,
        });
        return { selected: true };
      }
      case "excel.table.list": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const tables = await options.adapter.listTables(
          params.sheetId ? { worksheet: params.sheetId } : undefined,
        );
        return {
          tables: tables.map((table) => ({
            id: table.name,
            name: table.name,
            sheetId: table.worksheet,
            range: table.range,
          })),
        };
      }
      case "excel.table.read": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const tables = await options.adapter.listTables();
        const metadata = tables.find((table) => table.name === params.tableId);
        if (!metadata)
          throw new ControllerFailure("TABLE_NOT_FOUND", `Table not found: ${params.tableId}.`);
        // The current adapter loads a complete table before slicing. Bound that allocation too.
        checkRange({ sheetId: metadata.worksheet, address: metadata.range }, maxRead);
        const table = await options.adapter.readTable({ name: params.tableId });
        const rows = table.rows.slice(params.offset, params.offset + params.limit);
        return {
          table: {
            id: table.name,
            name: table.name,
            sheetId: table.worksheet,
            range: table.range,
          },
          headers: table.headers,
          rows,
          offset: params.offset,
          hasMore: params.offset + rows.length < table.rows.length,
        };
      }
      case "excel.operation.preview": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const mutationParams = toolDescriptorRegistry[params.method].paramsSchema.parse(
          params.params,
        );
        if (operations.size >= maxOperations)
          throw new ControllerFailure(
            "OPERATION_LIMIT_EXCEEDED",
            "Start a new Sommelier session to clear operation history.",
          );
        const before = await ensureValidMutation(params.method, mutationParams);
        const operation: StoredOperation = {
          operationId: idFactory("operation"),
          method: params.method,
          params: mutationParams,
          summary: mutationSummary(params.method, mutationParams),
          affectedRanges: [mutationParams.range],
          status: "previewed",
          before,
        };
        operations.set(operation.operationId, operation);
        emitOperation(operation, "previewed");
        return publicOperation(operation);
      }
      case "excel.operation.commit": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const operation = operations.get(params.operationId);
        if (!operation || operation.status !== "previewed") {
          throw new ControllerFailure(
            "OPERATION_NOT_PENDING",
            `Pending operation not found: ${params.operationId}.`,
          );
        }
        emitOperation(operation, "committing");
        try {
          if (!operation.before)
            throw new ControllerFailure("OPERATION_NOT_PENDING", "Missing preview snapshot.");
          await unchanged(operation.before);
          await approve({
            method: operation.method,
            summary: operation.summary,
            ranges: operation.affectedRanges,
            operationId: operation.operationId,
            ...approvalPreview(operation.method, operation.params, operation.before),
          });
          await unchanged(operation.before);
          await executeMutation(operation.method, operation.params);
          operation.after = await options.adapter.readRange({
            worksheet: operation.params.range.sheetId,
            range: operation.params.range.address,
          });
          emitOperation(operation, "committed");
          return publicOperation(operation);
        } catch (error) {
          emitOperation(operation, "failed", protocolErrorObject(error));
          throw error;
        }
      }
      case "excel.operation.undo": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        const operation = operations.get(params.operationId);
        if (!operation || operation.status !== "committed" || !operation.before) {
          throw new ControllerFailure(
            "OPERATION_NOT_UNDOABLE",
            `Committed operation not found: ${params.operationId}.`,
          );
        }
        if (operation.method === "excel.chart.create")
          throw new ControllerFailure(
            "OPERATION_NOT_UNDOABLE",
            "Remove this chart in Excel. Chart undo is not supported in v1.",
          );
        if (!operation.after)
          throw new ControllerFailure(
            "OPERATION_NOT_UNDOABLE",
            "No post-write snapshot is available.",
          );
        await unchanged(operation.after);
        if (
          operation.before.formulas?.some((row) => row.some(Boolean)) &&
          !options.adapter.restoreRange
        ) {
          throw new ControllerFailure(
            "OPERATION_NOT_UNDOABLE",
            "This adapter cannot restore formulas safely.",
          );
        }
        await approve({
          method: "excel.operation.undo",
          summary: `Undo ${operation.summary}`,
          ranges: operation.affectedRanges,
          operationId: operation.operationId,
        });
        await unchanged(operation.after);
        if (options.adapter.restoreRange) await options.adapter.restoreRange(operation.before);
        else
          await options.adapter.writeRange({
            worksheet: operation.before.worksheet,
            range: operation.before.range,
            values: operation.before.values,
          });
        emitOperation(operation, "undone");
        return publicOperation(operation);
      }
      case "excel.user.ask": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        if (!options.userInteraction?.ask) {
          throw new ControllerFailure("INTERACTION_UNAVAILABLE", "The host cannot ask the user.");
        }
        return options.userInteraction.ask({
          question: params.question,
          ...(params.choices ? { choices: params.choices } : {}),
        });
      }
      case "excel.user.notify": {
        const params = toolDescriptorRegistry[request.method].paramsSchema.parse(request.params);
        if (!options.userInteraction?.notify) {
          throw new ControllerFailure(
            "INTERACTION_UNAVAILABLE",
            "The host cannot notify the user.",
          );
        }
        return { acknowledged: await options.userInteraction.notify(params) };
      }
    }
  }

  function protocolErrorObject(failure: unknown): ProtocolErrorObject {
    return protocolError(undefined, failure).error;
  }

  return {
    async start() {
      if (started) return;
      stopped = false;
      const context = await options.adapter.getContext();
      currentSelection = selectionRange(context.selection);
      adapterUnsubscribe = await options.adapter.onSelectionChanged((selection) => {
        currentSelection = selectionRange(selection);
        emit({
          protocolVersion: PROTOCOL_VERSION,
          type: "event",
          event: "excel.selection.changed",
          data: { range: currentSelection },
        });
      });
      started = true;
    },

    async stop() {
      stopped = true;
      generation += 1;
      operations.clear();
      if (!started) return;
      const unsubscribe = adapterUnsubscribe;
      adapterUnsubscribe = undefined;
      started = false;
      await unsubscribe?.();
    },

    setContext(mode, ranges = []) {
      contextMode = mode;
      manualRanges = ranges.map((range) => ({ ...range }));
    },

    async handle(request) {
      if (stopped)
        return protocolError(
          request.id,
          new ControllerFailure("SESSION_STOPPED", "This Sommelier session has ended."),
        );
      const exclusive = toolDescriptorRegistry[request.method].capability !== "observe";
      if (exclusive && actionBusy)
        return protocolError(
          request.id,
          new ControllerFailure(
            "HOST_BUSY",
            "Finish the current action or approval before retrying.",
            "rejected",
          ),
        );
      if (exclusive) actionBusy = true;
      let operationId: string | undefined;
      if (
        request.method === "excel.operation.commit" ||
        request.method === "excel.operation.undo"
      ) {
        operationId = request.params.operationId;
      }
      try {
        const result = await dispatch(request);
        const validated = toolDescriptorRegistry[request.method].resultSchema.parse(result);
        const response = {
          protocolVersion: PROTOCOL_VERSION,
          type: "response",
          id: request.id,
          method: request.method,
          result: validated,
        } as ProtocolResponse;
        await audit(request.method, "success", operationId);
        return response;
      } catch (error) {
        const failure =
          error instanceof ControllerFailure
            ? error
            : error instanceof Error && error.name === "ZodError"
              ? new ControllerFailure("INVALID_PARAMS", error.message)
              : error;
        const outcome = failure instanceof ControllerFailure ? failure.outcome : "failure";
        await audit(request.method, outcome, operationId);
        return protocolError(request.id, failure);
      } finally {
        if (exclusive) actionBusy = false;
      }
    },

    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
