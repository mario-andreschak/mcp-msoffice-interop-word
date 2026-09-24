import { spawnSync } from "node:child_process";
import assert from "node:assert/strict";
const user = spawnSync(
  "docker",
  [
    "run",
    "--rm",
    "--network=none",
    "--entrypoint",
    "id",
    "mcp-word-check",
    "-u",
  ],
  { encoding: "utf8", timeout: 30000 },
);
assert.equal(user.status, 0);
assert.notEqual(user.stdout.trim(), "0");
const probe = spawnSync(process.execPath, ["scripts/verify-protocol.mjs"], {
  env: { ...process.env, MCP_WORD_TEST_DOCKER: "mcp-word-check" },
  stdio: "inherit",
  timeout: 60000,
});
assert.equal(probe.status, 0);
console.log(
  "Non-root network-isolated image passes actual protocol/status/error gates; Word COM is unavailable on Linux.",
);
