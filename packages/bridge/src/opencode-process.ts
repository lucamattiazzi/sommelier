import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";

/** Owns only the local harness API. The Sommelier relay remains hosted independently. */
export async function startOpenCodeProcess(directory: string, signal?: AbortSignal) {
  signal?.throwIfAborted();
  const password = randomBytes(32).toString("base64url");
  const username = "opencode";
  const child = spawn("opencode", ["serve", "--hostname", "127.0.0.1", "--port", "0"], {
    cwd: directory,
    stdio: ["ignore", "pipe", "pipe"],
    env: {
      ...process.env,
      OPENCODE_SERVER_PASSWORD: password,
      OPENCODE_SERVER_USERNAME: username,
    },
  });
  let exited = false;
  const exit = new Promise<void>((resolve) => {
    const done = () => {
      exited = true;
      resolve();
    };
    child.once("exit", done);
    child.once("error", done);
  });
  let closing: Promise<void> | undefined;
  const close = () => {
    closing ??= (async () => {
      if (exited) return;
      child.kill();
      const timer = setTimeout(() => child.kill("SIGKILL"), 3000);
      timer.unref();
      await exit;
      clearTimeout(timer);
    })();
    return closing;
  };
  try {
    const origin = await new Promise<string>((resolve, reject) => {
      const timer = setTimeout(() => finish(new Error("OpenCode startup timed out.")), 20000);
      const abort = () => finish(new Error("OpenCode startup cancelled."));
      const error = () =>
        finish(new Error("Could not start OpenCode. Check that it is installed."));
      const earlyExit = () => finish(new Error("OpenCode exited before its local API was ready."));
      const readers = [child.stdout, child.stderr].map((stream) => {
        let buffer = "";
        const read = (chunk: Buffer) => {
          buffer = (buffer + chunk.toString()).slice(-8192);
          const match = buffer.match(
            /opencode server listening on (http:\/\/127\.0\.0\.1:\d+)\r?\n/,
          );
          if (match?.[1]) finish(undefined, match[1]);
        };
        stream.on("data", read);
        return () => stream.off("data", read);
      });
      const cleanup = () => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", abort);
        child.off("error", error);
        child.off("exit", earlyExit);
        for (const off of readers) off();
      };
      child.once("error", error);
      child.once("exit", earlyExit);
      signal?.addEventListener("abort", abort, { once: true });
      function finish(error?: Error, origin?: string) {
        cleanup();
        if (error) reject(error);
        else if (origin) resolve(origin);
      }
    });
    return { origin, password, username, close };
  } catch (error) {
    await close();
    throw error;
  }
}
