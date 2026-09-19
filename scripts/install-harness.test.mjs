import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

test("installs self-contained skills in each harness's discoverable directory", () => {
  const home = mkdtempSync(join(tmpdir(), "pair-install-"));
  try {
    for (const [harness, directory] of Object.entries({
      codex: ".agents/skills",
      opencode: ".config/opencode/skills",
      claude: ".claude/skills",
      pi: ".pi/agent/skills",
    })) {
      execFileSync(process.execPath, ["scripts/install-harness.mjs", harness, "--home", home]);
      const installed = join(home, directory, "sommelier");
      assert.match(readFileSync(join(installed, "SKILL.md"), "utf8"), /name: sommelier/);
      const result = JSON.parse(
        execFileSync(process.execPath, [join(installed, "scripts/session.mjs"), "list"], {
          env: { ...process.env, SOMMELIER_HOME: join(home, "profiles") },
          encoding: "utf8",
        }),
      );
      assert.deepEqual(result.terminals, []);
      assert.ok(
        readFileSync(join(installed, "scripts/lib/encrypted-socket.mjs"), "utf8").includes(
          "AES-GCM",
        ),
      );
    }
    assert.throws(() =>
      execFileSync(process.execPath, ["scripts/install-harness.mjs", "mypy", "--home", home], {
        stdio: "pipe",
      }),
    );
  } finally {
    rmSync(home, { recursive: true, force: true });
  }
});
