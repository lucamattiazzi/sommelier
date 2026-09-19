import { mkdtemp, rm } from "node:fs/promises";
import { createServer } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { expect, it } from "vitest";
import { bridgeCommand, subscribeBridge } from "./bridge.js";

it("subscribes independently, unsubscribes and rejects failed local RPC", async () => {
  const dir = await mkdtemp(join(tmpdir(), "pair-ipc-"));
  const path = join(dir, "socket");
  const server = createServer((socket) =>
    socket.once("data", (data) => {
      const request = JSON.parse(data.toString());
      if (request.command === "watch")
        socket.write(
          `${JSON.stringify({ kind: "rpc.request", method: "excel.range.read", occurredAt: "2026-09-11T00:00:00Z" })}\n`,
        );
      else
        socket.end(
          `${JSON.stringify({ ok: false, error: { message: "Workbook disconnected" } })}\n`,
        );
    }),
  );
  await new Promise<void>((resolve) => server.listen(path, resolve));
  try {
    let off = () => {};
    const event = await new Promise<unknown>((resolve) => {
      off = subscribeBridge(path, resolve);
    });
    expect(event).toMatchObject({ kind: "rpc.request" });
    off();
    await expect(bridgeCommand(path, { command: "status" })).rejects.toThrow(
      "Workbook disconnected",
    );
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(dir, { recursive: true });
  }
});
