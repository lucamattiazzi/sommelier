import {
  type AddinController,
  type ApprovalRequest,
  type AskUserRequest,
  type AskUserResult,
  createAddinController,
} from "@lucamattiazzi/sommelier-addin-core";
import { HttpAgentAdapter } from "@lucamattiazzi/sommelier-agent-http";
import { createPairAddinSession, type PairAddinSession } from "@lucamattiazzi/sommelier-client";
import {
  type AgentSession,
  createAgentSession,
  type SessionEvent,
  type ToolPreview,
} from "@lucamattiazzi/sommelier-core";
import {
  type CellValue,
  createExcelTools,
  InMemoryExcelAdapter,
  OfficeJsExcelAdapter,
  OfficeJsWorkbookDriver,
  type WorkbookDriver,
} from "@lucamattiazzi/sommelier-excel";
import { VirtualWorkbookDriver } from "@lucamattiazzi/sommelier-testing";
import {
  createEncryptedSocket,
  createJsonSocketTransport,
  createPairIdentity,
  pairConnectionUrl,
  relaySocketUrl,
} from "@lucamattiazzi/sommelier-transport";
import { useCallback, useEffect, useId, useRef, useState } from "react";
import { shouldSendOnKey } from "./composer.js";
import { useModalFocus } from "./modal-focus.js";
import { type DirectAgentSettings, parseDirectAgentSettings } from "./settings.js";
import {
  parseSavedTerminals,
  rememberTerminal,
  type SavedTerminal,
  TERMINALS_KEY,
  terminalSetupPrompt,
} from "./terminals.js";

type PinnedRange = ApprovalRequest["ranges"][number];

type Screen = "home" | "direct" | "relay";

function workbookDriver(): WorkbookDriver {
  if (typeof Excel !== "undefined") return new OfficeJsWorkbookDriver();
  return new VirtualWorkbookDriver({
    selection: { worksheet: "Sheet1", range: "A1:B3" },
    worksheets: [
      {
        name: "Sheet1",
        values: [
          ["Item", "Amount"],
          ["Example", 10],
          ["Sample", 20],
        ],
        formulas: {},
        formats: {},
        tables: [],
      },
    ],
    namedItems: [],
  });
}

function createDirectSession(settings: DirectAgentSettings): AgentSession {
  return createAgentSession({
    agent: new HttpAgentAdapter({
      endpoint: settings.endpoint,
      ...(settings.token ? { getAccessToken: async () => settings.token } : {}),
    }),
    tools: createExcelTools({ driver: workbookDriver() }),
  });
}

interface DirectChatProps {
  readonly settings: DirectAgentSettings;
  readonly disconnect: () => void;
}

function DirectChat({ settings, disconnect }: DirectChatProps) {
  const [session] = useState(() => createDirectSession(settings));
  const [events, setEvents] = useState<SessionEvent[]>([]);
  const [message, setMessage] = useState("");
  const [approval, setApproval] = useState<{ requestId: string; preview: ToolPreview }>();
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    const unsubscribe = session.subscribe((event) => {
      setEvents((current) => [...current, event]);
      if (event.type === "approval.requested") {
        setApproval({ requestId: event.payload.requestId, preview: event.payload.preview });
      }
      if (event.type === "approval.resolved") setApproval(undefined);
      if (event.type === "turn.failed") setError(event.payload.error.message);
    });
    return () => {
      unsubscribe();
      session.cancel("Disconnected from Sommelier.");
    };
  }, [session]);

  const send = async (event: React.FormEvent): Promise<void> => {
    event.preventDefault();
    const text = message.trim();
    if (!text || busy) return;
    setMessage("");
    setError("");
    setBusy(true);
    try {
      await session.sendMessage({ text, context: { includeSelection: true } });
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      setApproval(undefined);
      setBusy(false);
    }
  };

  const messages = events.filter(
    (event) => event.type === "user.message" || event.type === "assistant.message",
  );

  return (
    <main>
      <ProductHeader status="Direct agent" disconnect={disconnect} />
      {typeof Excel === "undefined" && (
        <p className="demo-banner">Browser demo: only synthetic data is available.</p>
      )}
      {busy && (
        <button
          type="button"
          onClick={() => {
            session.cancel("Stopped by user.");
            setApproval(undefined);
          }}
        >
          Stop task
        </button>
      )}
      <section className="conversation" aria-live="polite">
        {messages.length === 0 && (
          <div className="empty">
            <span className="eyebrow">Connected</span>
            <h2>Your workbook, your agent</h2>
            <p>Ask the agent to inspect the current selection or propose a workbook change.</p>
          </div>
        )}
        {messages.map((item) => (
          <article
            key={item.eventId}
            className={`message ${item.type === "user.message" ? "user" : "agent"}`}
          >
            <span>{item.type === "user.message" ? "You" : "Agent"}</span>
            <p>{item.type === "user.message" ? item.payload.text : item.payload.content}</p>
          </article>
        ))}
      </section>
      {approval && (
        <ApprovalCard
          summary={approval.preview.summary}
          detail={`${approval.preview.affectedCells} cells · ${approval.preview.target ?? "workbook"}`}
          approve={() => void session.approve(approval.requestId)}
          reject={() => void session.reject(approval.requestId)}
        />
      )}
      {error && (
        <p className="error" role="alert">
          {error}
        </p>
      )}
      <form className="composer" onSubmit={send}>
        <textarea
          aria-label="Message your agent"
          placeholder="Ask about this workbook..."
          rows={3}
          value={message}
          onChange={(event) => setMessage(event.target.value)}
          onKeyDown={sendOnEnter}
        />
        <div className="composer-row">
          <span>Current selection is available on request</span>
          <button className="primary" type="submit" disabled={busy || !message.trim()}>
            {busy ? "Working" : "Send"}
          </button>
        </div>
      </form>
    </main>
  );
}

function sendOnEnter(event: React.KeyboardEvent<HTMLTextAreaElement>): void {
  const send = shouldSendOnKey({
    key: event.key,
    shiftKey: event.shiftKey,
    altKey: event.altKey,
    ctrlKey: event.ctrlKey,
    metaKey: event.metaKey,
    isComposing: event.nativeEvent.isComposing,
  });
  if (!send) return;
  event.preventDefault();
  event.currentTarget.form?.requestSubmit();
}

interface TerminalPairing {
  readonly terminal: SavedTerminal;
  readonly agentUrl: string;
}

interface RelayModeProps {
  readonly disconnect: () => void;
}

interface ChatMessage {
  readonly id: string;
  readonly sender: "user" | "agent";
  readonly content: string;
}

function RelayMode({ disconnect }: RelayModeProps) {
  const agentPromptId = useId();
  const [pairing, setPairing] = useState<TerminalPairing>();
  const [terminals, setTerminals] = useState<SavedTerminal[]>(() => {
    try {
      return parseSavedTerminals(localStorage.getItem(TERMINALS_KEY));
    } catch {
      return [];
    }
  });
  const terminalsRef = useRef(terminals);
  const [terminalName, setTerminalName] = useState("My terminal");
  const [storageNotice, setStorageNotice] = useState("");
  const nameId = useId();
  const persistTerminals = useCallback((records: SavedTerminal[]) => {
    terminalsRef.current = records;
    setTerminals(records);
    try {
      localStorage.setItem(TERMINALS_KEY, JSON.stringify(records));
    } catch {
      setStorageNotice(
        "Excel storage is unavailable. This connection will only be remembered while the pane is open.",
      );
    }
  }, []);
  const [status, setStatus] = useState("Not paired");
  const [error, setError] = useState("");
  const [approval, setApproval] = useState<ApprovalRequest>();
  const [question, setQuestion] = useState<AskUserRequest>();
  const [answer, setAnswer] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [operations, setOperations] = useState<
    {
      id: string;
      status: string;
      undoable?: boolean | undefined;
      summary?: string | undefined;
      error?: string;
    }[]
  >([]);
  const [draft, setDraft] = useState("");
  const [agentConnected, setAgentConnected] = useState(false);
  const [autoApprove, setAutoApprove] = useState(false);
  const [connecting, setConnecting] = useState(false);
  const [selection, setSelection] = useState<PinnedRange>();
  const [pinned, setPinned] = useState<PinnedRange>();
  const autoApproveId = useId();
  const autoApproveRef = useRef(false);
  const transcriptEnd = useRef<HTMLDivElement>(null);
  const lifecycle = useRef<PairAddinSession | undefined>(undefined);
  const controllerRef = useRef<AddinController | undefined>(undefined);
  const attempt = useRef(0);
  const abortFetch = useRef<AbortController | undefined>(undefined);
  const disposeSubscriptions = useRef<(() => void)[]>([]);
  const approvalResolution = useRef<((approved: boolean) => void) | undefined>(undefined);
  const answerResolution = useRef<((answer: AskUserResult) => void) | undefined>(undefined);

  const stopCurrent = useCallback(async (): Promise<void> => {
    abortFetch.current?.abort();
    setApproval(undefined);
    setQuestion(undefined);
    approvalResolution.current?.(false);
    approvalResolution.current = undefined;
    answerResolution.current?.({ answer: "Session ended without an answer." });
    answerResolution.current = undefined;
    for (const dispose of disposeSubscriptions.current) dispose();
    disposeSubscriptions.current = [];
    const session = lifecycle.current;
    lifecycle.current = undefined;
    controllerRef.current = undefined;
    await session?.stop();
  }, []);

  const createPairing = useCallback(
    async (existing?: SavedTerminal, name = "My terminal"): Promise<void> => {
      const currentAttempt = ++attempt.current;
      setConnecting(true);
      setError("");
      setPairing(undefined);
      setAgentConnected(false);
      setApproval(undefined);
      setQuestion(undefined);
      setOperations([]);
      setMessages([]);
      setPinned(undefined);
      setSelection(undefined);
      setAutoApprove(false);
      autoApproveRef.current = false;
      setStatus(existing ? "Reconnecting terminal…" : "Preparing a private connection…");
      try {
        await stopCurrent();
        if (currentAttempt !== attempt.current) return;
        const abort = new AbortController();
        abortFetch.current = abort;
        const relayOrigin =
          import.meta.env.VITE_SOMMELIER_RELAY_ORIGIN ||
          import.meta.env.VITE_PAIR_RELAY_ORIGIN ||
          window.location.origin;
        let terminal = existing;
        if (!terminal) {
          const response = await fetch(new URL("/api/config", relayOrigin), {
            signal: abort.signal,
          });
          if (!response.ok) throw new Error(`Sommelier server returned HTTP ${response.status}.`);
          const config: unknown = await response.json();
          if (
            !config ||
            typeof config !== "object" ||
            !("agentOrigin" in config) ||
            typeof config.agentOrigin !== "string"
          )
            throw new Error("The relay needs an update for encrypted pairing.");
          terminal = {
            ...createPairIdentity(),
            name: name.trim().slice(0, 64) || "My terminal",
            agentOrigin: config.agentOrigin,
            autoConnect: false,
          };
        }
        const remembered = terminal;
        const [addinUrl, agentUrl] = await Promise.all([
          pairConnectionUrl(relayOrigin, terminal, "addin"),
          pairConnectionUrl(terminal.agentOrigin, terminal, "agent"),
        ]);
        if (currentAttempt !== attempt.current) return;
        const created = { terminal, agentUrl };
        const socket = createEncryptedSocket(new WebSocket(relaySocketUrl(addinUrl)), addinUrl);
        const transport = createJsonSocketTransport(socket);
        const adapter =
          typeof Excel !== "undefined"
            ? new OfficeJsExcelAdapter()
            : new InMemoryExcelAdapter({
                sheets: [
                  {
                    name: "Sheet1",
                    values: [
                      ["Item", "Amount"],
                      ["Example", 10],
                      ["Sample", 20],
                    ],
                  },
                ],
                selection: "A1:B3",
              });
        const controller = createAddinController({
          adapter,
          workbook: { id: "active-workbook", name: "Active workbook" },
          requestApproval: (request) => {
            if (autoApproveRef.current) return true;
            return new Promise<boolean>((resolve) => {
              setApproval(request);
              approvalResolution.current = resolve;
            });
          },
          userInteraction: {
            ask: (request) =>
              new Promise<AskUserResult>((resolve) => {
                setQuestion(request);
                setAnswer("");
                answerResolution.current = resolve;
              }),
            notify: ({ message, level }) => {
              setMessages((current) => [
                ...current,
                { id: crypto.randomUUID(), sender: "agent", content: `${level}: ${message}` },
              ]);
              return true;
            },
          },
        });
        controllerRef.current = controller;
        disposeSubscriptions.current.push(
          controller.subscribe((event) => {
            if (event.event === "excel.selection.changed")
              setSelection(event.data.range ?? undefined);
            if (event.event === "excel.operation.changed") {
              const entry = {
                id: event.data.operationId,
                status: event.data.status,
                undoable: event.data.undoable,
                summary: event.data.summary,
                ...(event.data.error ? { error: event.data.error.message } : {}),
              };
              setOperations((current) =>
                [...current.filter((item) => item.id !== entry.id), entry].slice(-10),
              );
            }
          }),
        );
        const session = createPairAddinSession({ transport, controller });
        let peerPresent = false;
        disposeSubscriptions.current.push(
          transport.subscribe((message) => {
            if (
              message.type !== "event" ||
              message.event !== "pair.peer.changed" ||
              message.data.role !== "agent"
            )
              return;
            peerPresent = message.data.connected;
            setAgentConnected(peerPresent);
            if (peerPresent) {
              const latest =
                terminalsRef.current.find((record) => record.id === remembered.id) ?? remembered;
              persistTerminals(rememberTerminal(terminalsRef.current, latest));
            }
            setStatus(
              peerPresent
                ? "Connected · end-to-end encrypted"
                : "Waiting for the authorized terminal",
            );
            if (!peerPresent) {
              autoApproveRef.current = false;
              setAutoApprove(false);
              approvalResolution.current?.(false);
              approvalResolution.current = undefined;
              setApproval(undefined);
              answerResolution.current?.({
                answer: "Agent disconnected before the user answered.",
              });
              answerResolution.current = undefined;
              setQuestion(undefined);
            }
          }),
        );
        disposeSubscriptions.current.push(
          transport.subscribeState((state) => {
            if (state !== "disconnected") return;
            setAgentConnected(false);
            setStatus(
              "Connection closed. Reconnect your saved terminal; if another pane is using it, disconnect that pane first.",
            );
            void stopCurrent();
          }),
        );
        session.subscribeChat((message) =>
          setMessages((current) => [
            ...current,
            { id: message.data.messageId, sender: "agent", content: message.data.content },
          ]),
        );
        lifecycle.current = session;
        await session.start();
        if (currentAttempt !== attempt.current) {
          await session.stop();
          return;
        }
        const context = await adapter.getContext();
        setSelection({ sheetId: context.selection.worksheet, address: context.selection.range });
        setPairing(created);
        if (!peerPresent) setStatus("Waiting for the authorized terminal");
      } catch (failure) {
        if (currentAttempt !== attempt.current) return;
        await stopCurrent();
        setStatus("Connection failed");
        setError(failure instanceof Error ? failure.message : String(failure));
      } finally {
        if (currentAttempt === attempt.current) setConnecting(false);
      }
    },
    [stopCurrent, persistTerminals],
  );

  useEffect(() => {
    const automatic = terminalsRef.current.find((record) => record.autoConnect);
    if (automatic) void createPairing(automatic);
    return () => {
      attempt.current += 1;
      void stopCurrent();
    };
  }, [createPairing, stopCurrent]);

  useEffect(() => {
    if (messages.length === 0) return;
    transcriptEnd.current?.scrollIntoView({
      behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth",
      block: "end",
    });
  }, [messages.length]);

  const resolveApproval = (approved: boolean): void => {
    approvalResolution.current?.(approved);
    approvalResolution.current = undefined;
    setApproval(undefined);
  };
  const resolveAnswer = (value: AskUserResult): void => {
    answerResolution.current?.(value);
    answerResolution.current = undefined;
    setQuestion(undefined);
  };
  const questionDialog = useModalFocus(Boolean(question), () =>
    resolveAnswer({ answer: "The user declined to answer." }),
  );
  const sendMessage = (event: React.FormEvent): void => {
    event.preventDefault();
    const content = draft.trim();
    if (!lifecycle.current || !content || !agentConnected) return;
    try {
      const sent = lifecycle.current.sendUserMessage(content);
      setMessages((current) => [
        ...current,
        { id: sent.data.messageId, sender: "user", content: sent.data.content },
      ]);
      setDraft("");
      setError("");
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    }
  };
  const copyAgentPrompt = async (): Promise<void> => {
    if (!pairing) return;
    try {
      await navigator.clipboard.writeText(terminalSetupPrompt(pairing.agentUrl));
      setStatus("Setup copied. Paste it once in the terminal you want to remember.");
    } catch {
      setError("Select and copy the prompt manually; clipboard access is unavailable.");
    }
  };
  const togglePinned = (): void => {
    if (!controllerRef.current) return;
    if (pinned) {
      controllerRef.current.setContext("follow-selection");
      setPinned(undefined);
    } else if (selection) {
      controllerRef.current.setContext("manual", [selection]);
      setPinned(selection);
    }
  };
  const undo = async (operationId: string): Promise<void> => {
    const response = await controllerRef.current?.handle({
      protocolVersion: "0.2",
      type: "request",
      id: crypto.randomUUID(),
      method: "excel.operation.undo",
      params: { operationId },
    });
    if (response?.type === "error") setError(response.error.message);
  };
  const contextRange = pinned ?? selection;

  return (
    <main>
      <ProductHeader status="External agent" disconnect={disconnect} />
      {typeof Excel === "undefined" && (
        <p className="demo-banner">
          Browser demo: operations affect synthetic data only. Open this add-in in Excel to use your
          workbook.
        </p>
      )}
      <section className="relay-panel">
        <span className="eyebrow">Your workbook. Your agent.</span>
        <h2>{agentConnected ? pairing?.terminal.name : "Connect your terminal"}</h2>
        <p>
          Excels at pairing. Use your existing agent in OpenCode, Codex or Claude Code. Sommelier
          remembers the terminals you authorize.
        </p>
        <div className="privacy-note">
          End-to-end encrypted · Sommelier relays cannot read workbook data or conversations.
        </div>
        {terminals.length > 0 && (
          <details className="connection-settings" open={!agentConnected}>
            <summary>{agentConnected ? "Manage connection" : "Saved terminals"}</summary>
            <section className="terminal-list" aria-label="Authorized terminals">
              {terminals.map((terminal) => {
                const connected = pairing?.terminal.id === terminal.id && agentConnected;
                return (
                  <div className="terminal-card" key={terminal.id}>
                    <div className="terminal-title">
                      <strong>{terminal.name}</strong>
                      <small>{connected ? "Connected" : "Saved terminal"}</small>
                    </div>
                    <div className="button-row">
                      <button
                        className={connected ? "secondary" : "primary"}
                        type="button"
                        disabled={connecting || !!approval || !!question}
                        onClick={() => {
                          if (connected) {
                            attempt.current += 1;
                            void stopCurrent();
                            setAgentConnected(false);
                            setPairing(undefined);
                            setStatus("Disconnected");
                          } else void createPairing(terminal);
                        }}
                      >
                        {connected ? "Disconnect" : "Reconnect"}
                      </button>
                      <button
                        className="text-button"
                        type="button"
                        disabled={connecting || !!approval || !!question}
                        onClick={() => {
                          if (pairing?.terminal.id === terminal.id) {
                            attempt.current += 1;
                            void stopCurrent();
                            setPairing(undefined);
                            setAgentConnected(false);
                          }
                          persistTerminals(
                            terminalsRef.current.filter((record) => record.id !== terminal.id),
                          );
                          setStatus("Terminal forgotten on this device");
                        }}
                      >
                        Forget
                      </button>
                    </div>
                    <label className="toggle">
                      <input
                        type="checkbox"
                        checked={terminal.autoConnect}
                        onChange={(event) =>
                          persistTerminals(
                            rememberTerminal(terminalsRef.current, {
                              ...terminal,
                              autoConnect: event.target.checked,
                            }),
                          )
                        }
                      />
                      <span>Connect automatically when this pane opens</span>
                    </label>
                  </div>
                );
              })}
              <p className="fine-print">
                Start your saved Sommelier adapter, then reconnect here. One workbook per terminal
                connection.
              </p>
            </section>
          </details>
        )}
        {!agentConnected &&
          (!pairing || terminals.some((record) => record.id === pairing.terminal.id)) && (
            <div className="new-terminal">
              <label htmlFor={nameId}>Name this terminal</label>
              <input
                id={nameId}
                value={terminalName}
                maxLength={64}
                onChange={(event) => setTerminalName(event.target.value)}
                placeholder="e.g. Codex · Finance"
              />
              <button
                className={terminals.length ? "secondary" : "primary"}
                type="button"
                disabled={connecting || !terminalName.trim()}
                onClick={() => void createPairing(undefined, terminalName)}
              >
                {terminals.length ? "Pair another terminal" : "Pair a terminal"}
              </button>
            </div>
          )}
        {pairing &&
          !agentConnected &&
          !terminals.some((record) => record.id === pairing.terminal.id) && (
            <div className="pairing-ticket">
              <strong>One-time setup</strong>
              <p>
                Run sommelier pair --name desk in your terminal and paste the private URL below.
                Then start your harness adapter. You only need this URL once.
              </p>
              <button
                className="primary"
                type="button"
                onClick={() => {
                  void navigator.clipboard.writeText(pairing.agentUrl).then(
                    () =>
                      setStatus("Private URL copied. Paste it into the Sommelier setup command."),
                    () => setError("Clipboard unavailable. Copy the URL below."),
                  );
                }}
              >
                Copy connection URL
              </button>
              <details>
                <summary>Show private connection URL</summary>
                <textarea
                  aria-label="Private connection URL"
                  readOnly
                  rows={3}
                  value={pairing.agentUrl}
                />
              </details>
              <details>
                <summary>Alternative: connect using only a skill</summary>
                <button type="button" className="secondary" onClick={() => void copyAgentPrompt()}>
                  Copy skill setup
                </button>
                <label htmlFor={agentPromptId}>Private setup for your terminal</label>
                <textarea
                  id={agentPromptId}
                  className="agent-prompt"
                  readOnly
                  rows={6}
                  value={terminalSetupPrompt(pairing.agentUrl)}
                />
              </details>
              <a href="/setup.html" target="_blank" rel="noreferrer">
                Install adapters · reconnect without prompts
              </a>
              <button
                className="text-button"
                type="button"
                onClick={() => {
                  attempt.current += 1;
                  void stopCurrent();
                  setPairing(undefined);
                  setAgentConnected(false);
                  setStatus("Setup cancelled");
                }}
              >
                Cancel setup
              </button>
            </div>
          )}
        {connecting && <output>Connecting to the relay…</output>}
        <output className="status-line" aria-live="polite">
          {status}
        </output>
        {storageNotice && <output>{storageNotice}</output>}
        <label className="toggle" htmlFor={autoApproveId}>
          <input
            id={autoApproveId}
            type="checkbox"
            checked={autoApprove}
            disabled={!pairing || !!approval}
            onChange={(event) => {
              autoApproveRef.current = event.target.checked;
              setAutoApprove(event.target.checked);
            }}
          />
          <span>Approve changes automatically for this connection only. Previews still run.</span>
        </label>
      </section>
      {contextRange && (
        <section className="context-panel">
          <strong>{pinned ? "Pinned context" : "Following selection"}</strong>
          <code>
            {contextRange.sheetId}!{contextRange.address}
          </code>
          <button type="button" onClick={togglePinned}>
            {pinned ? "Follow selection" : "Pin selection"}
          </button>
          <small>Context is a hint for the agent; it does not restrict access to this range.</small>
        </section>
      )}
      {pairing && (
        <>
          <section className="conversation" aria-live="polite">
            {messages.length === 0 && (
              <div className="empty">
                <h2>What would you like to do?</h2>
                <p>Try: “Read my selection and explain what it contains.”</p>
              </div>
            )}
            {messages.map((item) => (
              <article key={item.id} className={`message ${item.sender}`}>
                <span>{item.sender === "user" ? "You" : "Agent"}</span>
                <p>{item.content}</p>
              </article>
            ))}
            <div ref={transcriptEnd} />
          </section>
          <form className="composer" onSubmit={sendMessage}>
            <textarea
              aria-label="Message your paired agent"
              placeholder={
                agentConnected ? "Ask about this workbook..." : "Waiting for your agent to connect"
              }
              rows={3}
              maxLength={32000}
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={sendOnEnter}
            />
            <div className="composer-row">
              <span>{agentConnected ? "Agent connected" : "Agent not connected"}</span>
              <button className="primary" type="submit" disabled={!agentConnected || !draft.trim()}>
                Send
              </button>
            </div>
          </form>
        </>
      )}
      {operations.length > 0 && (
        <details className="operation-history">
          <summary>Recent operations ({operations.length})</summary>
          {operations.map((operation) => (
            <div key={operation.id}>
              <span>{operation.status}</span>
              <small>{operation.summary ?? operation.id}</small>
              {operation.error && <p>{operation.error}</p>}
              {operation.status === "committed" && operation.undoable !== false && (
                <button
                  type="button"
                  disabled={!!approval || !!question || !controllerRef.current}
                  onClick={() => void undo(operation.id)}
                >
                  Undo
                </button>
              )}
            </div>
          ))}
        </details>
      )}
      <button className="text-button" type="button" onClick={disconnect}>
        Disconnect / use a direct agent
      </button>
      {approval && (
        <ApprovalCard
          summary={approval.summary}
          detail={approval.ranges.map((range) => `${range.sheetId}!${range.address}`).join(", ")}
          before={approval.before}
          after={approval.after}
          approve={() => resolveApproval(true)}
          reject={() => resolveApproval(false)}
        />
      )}
      {question && (
        <div className="modal-backdrop">
          <section
            ref={questionDialog}
            className="approval"
            role="dialog"
            aria-modal="true"
            aria-label="Agent question"
          >
            <h2>{question.question}</h2>
            {question.choices?.map((choice) => (
              <button
                type="button"
                key={choice.id}
                onClick={() => resolveAnswer({ answer: choice.label, choiceId: choice.id })}
              >
                {choice.label}
              </button>
            ))}
            <textarea
              aria-label="Answer the agent"
              value={answer}
              onChange={(event) => setAnswer(event.target.value)}
            />
            <button
              type="button"
              disabled={!answer.trim()}
              onClick={() => resolveAnswer({ answer: answer.trim() })}
            >
              Send answer
            </button>
            <button
              type="button"
              onClick={() => resolveAnswer({ answer: "The user declined to answer." })}
            >
              Decline
            </button>
          </section>
        </div>
      )}
      {error && (
        <p className="error" role="alert">
          {error}
        </p>
      )}
      <footer>
        <a href="/setup.html" target="_blank" rel="noreferrer">
          Setup
        </a>{" "}
        ·{" "}
        <a href="/privacy.html" target="_blank" rel="noreferrer">
          Data &amp; privacy
        </a>{" "}
        ·{" "}
        <a href="/support.html" target="_blank" rel="noreferrer">
          Support
        </a>
      </footer>
    </main>
  );
}

function ProductHeader({
  status,
  disconnect,
}: {
  readonly status: string;
  readonly disconnect: () => void;
}) {
  return (
    <header className="product-header">
      <button className="wordmark" type="button" onClick={disconnect}>
        <b>Sommelier</b>
      </button>
      <span className="connection-pill">{status}</span>
    </header>
  );
}

function ApprovalCard({
  summary,
  detail,
  approve,
  reject,
  before,
  after,
}: {
  readonly summary: string;
  readonly detail: string;
  readonly before?: readonly (readonly CellValue[])[] | undefined;
  readonly after?: readonly (readonly CellValue[])[] | undefined;
  readonly approve: () => void;
  readonly reject: () => void;
}) {
  const dialog = useModalFocus(true, reject);
  return (
    <div className="modal-backdrop">
      <section
        ref={dialog}
        className="approval"
        role="alertdialog"
        aria-modal="true"
        aria-label="Workbook mutation approval"
      >
        <span className="eyebrow">Approval needed</span>
        <h2>{summary}</h2>
        <p>{detail}</p>
        {before && after && (
          <div className="change-preview">
            <table>
              <thead>
                <tr>
                  <th>Row</th>
                  <th>Before</th>
                  <th>After</th>
                </tr>
              </thead>
              <tbody>
                {before.slice(0, 8).map((row, r) => (
                  // biome-ignore lint/suspicious/noArrayIndexKey: Snapshot rows are fixed workbook coordinates, never reordered.
                  <tr key={`preview-${r}`}>
                    <td>{r + 1}</td>
                    <td>{JSON.stringify(row.slice(0, 5))}</td>
                    <td>{JSON.stringify(after[r]?.slice(0, 5))}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <small>
              Preview shows up to 8 rows and 5 columns. The approval covers the complete target
              range.
            </small>
          </div>
        )}
        <div className="button-row">
          <button type="button" onClick={reject}>
            Reject
          </button>
          <button className="approve" type="button" onClick={approve}>
            Approve once
          </button>
        </div>
      </section>
    </div>
  );
}

export function App() {
  const endpointId = useId();
  const tokenId = useId();
  const [screen, setScreen] = useState<Screen>("relay");
  const [endpoint, setEndpoint] = useState(() => {
    try {
      return localStorage.getItem("ai-cdl-pair-endpoint") ?? "";
    } catch {
      return "";
    }
  });
  const [token, setToken] = useState("");
  const [settings, setSettings] = useState<DirectAgentSettings>();
  const [error, setError] = useState("");

  if (screen === "direct" && settings) {
    return (
      <DirectChat
        settings={settings}
        disconnect={() => {
          setSettings(undefined);
          setToken("");
          setScreen("relay");
        }}
      />
    );
  }
  if (screen === "relay") return <RelayMode disconnect={() => setScreen("home")} />;

  const connectDirect = (event: React.FormEvent): void => {
    event.preventDefault();
    try {
      const parsed = parseDirectAgentSettings(endpoint, token);
      try {
        if (!new URL(parsed.endpoint).search)
          localStorage.setItem("ai-cdl-pair-endpoint", parsed.endpoint);
      } catch {
        /* Storage can be disabled in an Office webview. */
      }
      setSettings(parsed);
      setScreen("direct");
      setError("");
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    }
  };

  return (
    <main className="onboarding">
      <header className="hero">
        <span className="edition">FREE EXCEL ADD-IN</span>
        <h1>
          Bring your agent.
          <br />
          Keep control of Excel.
        </h1>
        <p>
          Sommelier connects an agent you trust to the workbook in front of you. Reads are bounded.
          Changes wait for you.
        </p>
      </header>
      <form className="connect-card" onSubmit={connectDirect}>
        <span className="step">01</span>
        <div>
          <h2>Connect directly</h2>
          <p>Best for an HTTPS agent endpoint that supports Sommelier protocol 0.1 and CORS.</p>
        </div>
        <label htmlFor={endpointId}>Agent endpoint</label>
        <input
          id={endpointId}
          type="url"
          placeholder="https://agent.example.com/turn"
          value={endpoint}
          onChange={(event) => setEndpoint(event.target.value)}
          required
        />
        <label htmlFor={tokenId}>
          Bearer token <span>optional, never saved</span>
        </label>
        <input
          id={tokenId}
          type="password"
          autoComplete="off"
          value={token}
          onChange={(event) => setToken(event.target.value)}
        />
        <button className="primary wide" type="submit">
          Connect agent
        </button>
      </form>
      <button className="relay-choice" type="button" onClick={() => setScreen("relay")}>
        <span className="step">02</span>
        <span>
          <b>Pair an external or local agent</b>
          <small>Reconnect an authorized terminal with end-to-end encryption.</small>
        </span>
        <strong>→</strong>
      </button>
      {error && (
        <p className="error" role="alert">
          {error}
        </p>
      )}
      <footer>
        Open protocol · Approval before writes · No agent credentials on Sommelier servers
      </footer>
    </main>
  );
}
