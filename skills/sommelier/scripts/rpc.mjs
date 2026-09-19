#!/usr/bin/env node

import { randomUUID } from "node:crypto";

function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index < 0 ? fallback : process.argv[index + 1];
}

const url = process.env.SOMMELIER_URL ?? process.env.AI_CDL_PAIR_URL;
const method = argument("--method");
const paramsText = argument("--params", "{}");
const timeoutText = argument("--timeout", "120000");

if (!url) throw new Error("SOMMELIER_URL is required.");
if (!method) throw new Error("--method is required.");

let params;
try {
  params = JSON.parse(paramsText);
} catch {
  throw new Error("--params must be valid JSON.");
}
if (!params || Array.isArray(params) || typeof params !== "object") {
  throw new Error("--params must be a JSON object.");
}

const timeoutMs = Number(timeoutText);
if (!Number.isInteger(timeoutMs) || timeoutMs <= 0) throw new Error("--timeout must be positive.");

const id = randomUUID();
const socket = new WebSocket(url);

const response = await new Promise((resolve, reject) => {
  let settled = false;
  const rejectOnce = (error) => {
    if (settled) return;
    settled = true;
    clearTimeout(timer);
    reject(error);
  };
  const timer = setTimeout(() => {
    socket.close();
    rejectOnce(new Error(`Sommelier request timed out after ${timeoutMs}ms.`));
  }, timeoutMs);

  socket.addEventListener("open", () => {
    socket.send(JSON.stringify({ protocolVersion: "0.2", type: "request", id, method, params }));
  });
  socket.addEventListener("message", (event) => {
    try {
      const message = JSON.parse(String(event.data));
      if (message.type === "event") {
        process.stderr.write(`${JSON.stringify(message)}\n`);
        return;
      }
      if (message.id !== id) return;
      settled = true;
      clearTimeout(timer);
      resolve(message);
    } catch (error) {
      rejectOnce(error);
    }
  });
  socket.addEventListener("error", () => {
    rejectOnce(new Error("Could not connect to the Sommelier URL."));
  });
  socket.addEventListener("close", () => {
    rejectOnce(new Error("Sommelier connection closed before a response was received."));
  });
});

socket.close();
process.stdout.write(`${JSON.stringify(response)}\n`);
if (response.type === "error") process.exitCode = 2;
