import { describe, expect, it } from "vitest";
import { parseDirectAgentSettings } from "./settings.js";

describe("direct agent settings", () => {
  it("accepts HTTPS and keeps bearer tokens in the returned in-memory settings", () => {
    expect(parseDirectAgentSettings("https://agent.example.com/turn", " secret ")).toEqual({
      endpoint: "https://agent.example.com/turn",
      token: "secret",
    });
  });

  it("rejects insecure remote endpoints but permits localhost development", () => {
    expect(() => parseDirectAgentSettings("http://agent.example.com", "")).toThrow("HTTPS");
    expect(parseDirectAgentSettings("http://127.0.0.1:8787/agent", "")).toEqual({
      endpoint: "http://127.0.0.1:8787/agent",
    });
  });

  it("rejects credentials embedded in an endpoint URL", () => {
    expect(() => parseDirectAgentSettings("https://user:secret@agent.example.com", "")).toThrow(
      /credentials/i,
    );
  });
});
