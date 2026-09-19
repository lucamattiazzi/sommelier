#!/usr/bin/env node

import { createPairRelayServer } from "./index.js";

function option(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index < 0 ? undefined : process.argv[index + 1];
}

async function main(): Promise<void> {
  if (process.argv[2] !== "relay") {
    console.error("Usage: sommelier-relay relay [--host 127.0.0.1] [--port 4310]");
    process.exitCode = 1;
    return;
  }
  const host = option("--host") ?? "127.0.0.1";
  const portValue = option("--port") ?? "4310";
  const port = Number(portValue);
  if (!Number.isInteger(port) || port < 0 || port > 65_535) {
    console.error(`Invalid port: ${portValue}`);
    process.exitCode = 1;
    return;
  }
  const relay = await createPairRelayServer({ host, port });
  console.log(`Sommelier relay listening on ws://${relay.host}:${relay.port}/pair`);
  const shutdown = async (): Promise<void> => {
    await relay.close();
    process.exit(0);
  };
  process.once("SIGINT", () => void shutdown());
  process.once("SIGTERM", () => void shutdown());
}

void main();
