import { spawn } from "node:child_process";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { afterEach, expect, it, vi } from "vitest";
import { startOpenCodeProcess } from "./opencode-process.js";

vi.mock("node:child_process", () => ({ spawn: vi.fn() }));
afterEach(() => vi.restoreAllMocks());

function child() {
  const process = Object.assign(new EventEmitter(), {
    stdout: new PassThrough(),
    stderr: new PassThrough(),
    kill: vi.fn(() => {
      queueMicrotask(() => process.emit("exit", null, "SIGTERM"));
      return true;
    }),
  });
  vi.mocked(spawn).mockReturnValue(process as unknown as ReturnType<typeof spawn>);
  return process;
}

it("starts an authenticated loopback API, reads split startup output and stops its child", async () => {
  const process = child();
  const pending = startOpenCodeProcess("/synthetic/project");
  process.stdout.write("opencode server listening on http://127.0.");
  process.stdout.write("0.1:45678\n");
  const server = await pending;
  expect(server.origin).toBe("http://127.0.0.1:45678");
  expect(server.password.length).toBeGreaterThanOrEqual(32);
  expect(spawn).toHaveBeenCalledWith(
    "opencode",
    ["serve", "--hostname", "127.0.0.1", "--port", "0"],
    expect.objectContaining({
      cwd: "/synthetic/project",
      env: expect.objectContaining({
        OPENCODE_SERVER_PASSWORD: server.password,
        OPENCODE_SERVER_USERNAME: server.username,
      }),
    }),
  );
  await server.close();
  await server.close();
  expect(process.kill).toHaveBeenCalledTimes(1);
});

it("reports an early exit without printing process output", async () => {
  const process = child();
  const pending = startOpenCodeProcess("/synthetic/project");
  process.stderr.write("private provider details");
  process.emit("exit", 1);
  await expect(pending).rejects.toThrow("OpenCode exited before its local API was ready");
});

it("kills the child when startup is cancelled", async () => {
  const process = child();
  const abort = new AbortController();
  const pending = startOpenCodeProcess("/synthetic/project", abort.signal);
  abort.abort();
  await expect(pending).rejects.toThrow("OpenCode startup cancelled");
  expect(process.kill).toHaveBeenCalledTimes(1);
});
