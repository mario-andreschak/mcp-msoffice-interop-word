import { spawn } from "node:child_process";
import { once } from "node:events";
import { createInterface } from "node:readline";
import { resolve, join } from "node:path";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import assert from "node:assert/strict";
const entry = resolve(process.argv[2] ?? "dist/index.js");
for (const modern of [true, false]) {
  const workspace = await mkdtemp(join(tmpdir(), "word-protocol-"));
  const child = spawn(
    process.env.MCP_WORD_TEST_DOCKER ? "docker" : process.execPath,
    process.env.MCP_WORD_TEST_DOCKER
      ? [
          "run",
          "--rm",
          "--network=none",
          "-i",
          process.env.MCP_WORD_TEST_DOCKER,
        ]
      : [entry],
    {
      env: {
        ...process.env,
        MCP_TRANSPORT: "stdio",
        WORD_ALLOWED_ROOTS: JSON.stringify([workspace]),
      },
      stdio: ["pipe", "pipe", "pipe"],
    },
  );
  let stderr = "",
    next = 1;
  const pending = new Map();
  child.stderr.on("data", (c) => (stderr += c));
  const exited = once(child, "exit");
  const lines = createInterface({ input: child.stdout });
  lines.on("line", (line) => {
    try {
      const message = JSON.parse(line);
      const found = pending.get(message.id);
      if (found) {
        pending.delete(message.id);
        found.resolve(message);
      }
    } catch (error) {
      for (const p of pending.values())
        p.reject(new Error("Non-JSON stdout: " + line, { cause: error }));
    }
  });
  child.on("error", (error) => {
    for (const p of pending.values()) p.reject(error);
  });
  child.on("exit", () => {
    for (const p of pending.values())
      p.reject(new Error("Unexpected exit: " + stderr));
  });
  function rpc(method, params = {}) {
    const id = next++;
    if (modern)
      params = {
        ...params,
        _meta: {
          "io.modelcontextprotocol/protocolVersion": "2026-07-28",
          "io.modelcontextprotocol/clientCapabilities": {},
        },
      };
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error("Timeout: " + method + "; " + stderr));
      }, 15000);
      pending.set(id, {
        resolve: (m) => {
          clearTimeout(timer);
          resolve(m);
        },
        reject: (e) => {
          clearTimeout(timer);
          reject(e);
        },
      });
      child.stdin.write(
        JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n",
      );
    });
  }
  try {
    const hello = modern
      ? await rpc("server/discover")
      : await rpc("initialize", {
          protocolVersion: "2025-11-25",
          capabilities: {},
          clientInfo: { name: "packed-word-gate", version: "1" },
        });
    assert.equal(hello.error, undefined);
    if (modern) assert.equal(hello.result.resultType, "complete");
    else
      assert.equal(hello.result.serverInfo.name, "mcp-msoffice-interop-word");
    if (!modern)
      child.stdin.write(
        JSON.stringify({
          jsonrpc: "2.0",
          method: "notifications/initialized",
        }) + "\n",
      );
    const list = await rpc("tools/list");
    assert.equal(list.result.tools.length, 38);
    assert.ok(list.result.tools.some((t) => t.name === "word_toggleBold"));
    assert.ok(
      list.result.tools.some((t) => t.name === "word_applyTableAutoFormat"),
    );
    if (modern) {
      assert.equal(list.result.resultType, "complete");
      assert.equal(typeof list.result.ttlMs, "number");
      assert.ok(list.result.cacheScope);
    }
    const status = await rpc("tools/call", {
      name: "word_runtimeStatus",
      arguments: {},
    });
    assert.equal(status.error, undefined);
    assert.equal(status.result.structuredContent.platform, process.platform);
    if (process.platform === "win32")
      assert.equal(
        status.result.structuredContent.nativeAddonAvailable,
        true,
        "actual installed winax addon loads on Windows",
      );
    const invalid = await rpc("tools/call", {
      name: "word_setParagraphAlignment",
      arguments: { alignment: 999 },
    });
    assert.ok(invalid.error || invalid.result?.isError);
    const close = await rpc("tools/call", {
      name: "word_closeActiveDocument",
      arguments: { saveChanges: -2 },
    });
    assert.ok(close.error || close.result?.isError);
    const escaped = await rpc("tools/call", {
      name: "word_openDocument",
      arguments: { filePath: join(workspace, "../outside.docx") },
    });
    assert.ok(escaped.error || escaped.result?.isError);
    if (process.platform !== "win32") {
      const unsupported = await rpc("tools/call", {
        name: "word_createDocument",
        arguments: {},
      });
      assert.equal(unsupported.result.isError, true);
      assert.match(JSON.stringify(unsupported.result.content), /Windows/);
    }
    const unknown = await rpc("tools/call", {
      name: "does-not-exist",
      arguments: {},
    });
    assert.ok(unknown.error || unknown.result?.isError);
    child.stdin.end();
    const closed = await Promise.race([
      exited,
      new Promise((_, reject) => {
        const t = setTimeout(
          () => reject(new Error("EOF did not shut down")),
          5000,
        );
        t.unref();
      }),
    ]);
    assert.equal(closed[0], 0, stderr);
    console.log(
      (modern ? "2026-07-28" : "2025-11-25") +
        ": actual Word status/catalog/schema/path/native-platform errors and clean EOF passed",
    );
  } finally {
    lines.close();
    if (child.exitCode === null) child.kill();
    await rm(workspace, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 200,
    });
  }
}
