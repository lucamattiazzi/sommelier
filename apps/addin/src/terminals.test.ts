import { createPairIdentity } from "@lucamattiazzi/sommelier-transport";
import { expect, it } from "vitest";
import { parseSavedTerminals, rememberTerminal } from "./terminals.js";

it("restores named identities without persisting workbook data or write approval", () => {
  const record = {
    ...createPairIdentity(),
    name: "Codex · Desk",
    agentOrigin: "wss://pair.example.test",
    autoConnect: false,
  };
  const saved = rememberTerminal([], record);
  expect(parseSavedTerminals(JSON.stringify(saved))).toEqual([record]);
  expect(
    parseSavedTerminals(JSON.stringify([{ ...record, autoApprove: true, cells: ["private"] }])),
  ).toEqual([record]);
});
it("rejects corrupted credentials and insecure remote URLs", () => {
  expect(parseSavedTerminals("not JSON")).toEqual([]);
  expect(
    parseSavedTerminals(
      JSON.stringify([
        {
          ...createPairIdentity(),
          name: "Desk",
          agentOrigin: "ws://remote.test",
          autoConnect: true,
        },
      ]),
    ),
  ).toEqual([]);
});
it("replaces an existing identity and permits only one automatic connection", () => {
  const first = {
    ...createPairIdentity(),
    name: "one",
    agentOrigin: "wss://pair.example.test",
    autoConnect: true,
  };
  const second = {
    ...createPairIdentity(),
    name: "two",
    agentOrigin: "wss://pair.example.test",
    autoConnect: true,
  };
  const saved = rememberTerminal([first], second);
  expect(saved.map((s) => s.autoConnect)).toEqual([false, true]);
  expect(rememberTerminal(saved, { ...second, name: "renamed" })).toHaveLength(2);
});
