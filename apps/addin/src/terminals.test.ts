import { createPairIdentity } from "@lucamattiazzi/sommelier-transport";
import { expect, it } from "vitest";
import { parseSavedTerminals, rememberTerminal, terminalSetupPrompt } from "./terminals.js";

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

it("provides a self-contained prompt for an existing agent without a preinstalled skill", () => {
  const url = "wss://relay.example.test/connect?id=synthetic#private-synthetic-key";
  const prompt = terminalSetupPrompt(url);
  expect(prompt).toContain(url);
  expect(prompt).toContain("@lucamattiazzi/sommelier@0.2.0-beta.1");
  expect(prompt).toContain("dist/skill");
  expect(prompt).not.toContain("SKILL.md");
  expect(prompt).not.toContain("references/protocol.md");
  expect(prompt).not.toContain("installed sommelier skill");
  expect(prompt).toContain("excel.range.read");
  expect(prompt).toContain("excel.chart.create");
  expect(prompt).toContain("taskPaneConnected");
  expect(prompt).toContain("idle");
  expect(prompt).toContain("excel.operation.preview");
  expect(prompt).toContain("excel.operation.commit");
  expect(prompt).toContain("Do not launch another agent");
});
