import { resolve } from "node:path";
import { createPairServer } from "./index.js";

async function main(): Promise<void> {
  const publicOrigin = process.env.SOMMELIER_PUBLIC_ORIGIN ?? process.env.PAIR_PUBLIC_ORIGIN;
  if (!publicOrigin) throw new Error("SOMMELIER_PUBLIC_ORIGIN is required.");

  const port = Number(process.env.PORT ?? "3000");
  if (!Number.isInteger(port) || port < 1 || port > 65_535) throw new Error("PORT is invalid.");

  const addinOrigin = process.env.SOMMELIER_ADDIN_ORIGIN ?? process.env.PAIR_ADDIN_ORIGIN;
  const agentOrigin = process.env.SOMMELIER_AGENT_ORIGIN ?? process.env.PAIR_AGENT_ORIGIN;

  const server = await createPairServer({
    host: process.env.HOST ?? "127.0.0.1",
    port,
    publicOrigin,
    ...(addinOrigin ? { addinOrigin } : {}),
    ...(agentOrigin ? { agentOrigin } : {}),
    staticDirectory: resolve(
      process.env.SOMMELIER_STATIC_DIR ?? process.env.PAIR_STATIC_DIR ?? "apps/addin/dist",
    ),
    bridgeScript: resolve(
      process.env.SOMMELIER_BRIDGE_SCRIPT ??
        process.env.PAIR_BRIDGE_SCRIPT ??
        "skills/sommelier/scripts/session.mjs",
    ),
  });

  console.log(`Sommelier server listening on ${server.host}:${server.port}`);

  const shutdown = async (): Promise<void> => {
    await server.close();
    process.exit(0);
  };
  process.once("SIGINT", () => void shutdown());
  process.once("SIGTERM", () => void shutdown());
}

void main();
