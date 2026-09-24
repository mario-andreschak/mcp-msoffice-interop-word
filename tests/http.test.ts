import test from "node:test";
import assert from "node:assert/strict";
import { request as httpRequest } from "node:http";
import { startHttp } from "../src/http.js";
const token = "word-regression-operator-token-000000000000000";
function raw(port: number, path: string, headers: Record<string, string> = {}) {
  return new Promise<number>((resolve, reject) => {
    const req = httpRequest(
      { hostname: "127.0.0.1", port, path, headers },
      (res) => {
        res.resume();
        resolve(res.statusCode!);
      },
    );
    req.on("error", reject);
    req.end();
  });
}
test("real HTTP serves modern and legacy requests while enforcing operator, Origin and Host boundaries", async (t) => {
  const old = { ...process.env };
  process.env.PORT = "0";
  process.env.HOST = "127.0.0.1";
  process.env.MCP_AUTH_TOKEN = token;
  delete process.env.MCP_PUBLIC_ORIGIN;
  const app = await startHttp();
  t.after(async () => {
    await app.close();
    for (const key of ["PORT", "HOST", "MCP_AUTH_TOKEN", "MCP_PUBLIC_ORIGIN"]) {
      if (old[key] === undefined) delete process.env[key];
      else process.env[key] = old[key];
    }
  });
  const base = `http://127.0.0.1:${app.port}`;
  const authorization = { Authorization: `Bearer ${token}` };
  assert.equal(await raw(app.port, "/mcp", { Host: "evil.test" }), 421);
  assert.equal(
    await raw(app.port, "/mcp", {
      Origin: "https://evil.test",
      ...authorization,
    }),
    403,
  );
  assert.equal(
    await raw(app.port, "/mcp", { Origin: "null", ...authorization }),
    403,
  );
  assert.equal(await raw(app.port, "/mcp?token=" + token), 401);
  assert.equal(await raw(app.port, "/sse"), 401);
  async function rpc(
    method: string,
    params: any = {},
    modern = true,
    extra: Record<string, string> = {},
  ) {
    if (modern)
      params = {
        ...params,
        _meta: {
          "io.modelcontextprotocol/protocolVersion": "2026-07-28",
          "io.modelcontextprotocol/clientCapabilities": {},
        },
      };
    const res = await fetch(base + "/mcp", {
      method: "POST",
      headers: {
        ...authorization,
        "Content-Type": "application/json",
        Accept: "application/json, text/event-stream",
        ...(modern
          ? {
              "MCP-Protocol-Version": "2026-07-28",
              "Mcp-Method": method,
              ...(params.name ? { "Mcp-Name": params.name } : {}),
            }
          : {}),
        ...extra,
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method, params }),
    });
    const text = await res.text();
    const body = JSON.parse(
      res.headers.get("content-type")?.includes("text/event-stream")
        ? text
            .split("\n")
            .find((line) => line.startsWith("data: "))!
            .slice(6)
        : text,
    );
    return { res, body };
  }
  const discovery = await rpc("server/discover");
  assert.equal(discovery.res.status, 200);
  assert.equal(discovery.body.result.resultType, "complete");
  const list = await rpc("tools/list");
  assert.equal(list.body.result.tools.length, 38);
  const status = await rpc("tools/call", {
    name: "word_runtimeStatus",
    arguments: {},
  });
  assert.equal(status.body.result.structuredContent.platform, process.platform);
  const legacy = await rpc(
    "initialize",
    {
      protocolVersion: "2025-11-25",
      capabilities: {},
      clientInfo: { name: "legacy-http", version: "1" },
    },
    false,
  );
  assert.equal(legacy.body.result.serverInfo.name, "mcp-msoffice-interop-word");
  const legacyList = await rpc("tools/list", {}, false);
  assert.equal(legacyList.body.result.tools.length, 38);
  const mismatch = await rpc("tools/list", {}, true, {
    "Mcp-Method": "tools/call",
  });
  assert.ok(mismatch.res.status >= 400 || mismatch.body.error);
  const wrongContent = await fetch(base + "/mcp", {
    method: "POST",
    headers: authorization,
    body: "x",
  });
  assert.equal(wrongContent.status, 415);
  const oversized = await fetch(base + "/mcp", {
    method: "POST",
    headers: { ...authorization, "Content-Type": "application/json" },
    body: JSON.stringify({ large: "x".repeat(2 * 1024 * 1024) }),
  });
  assert.equal(oversized.status, 413);
  await oversized.text();
  const malformed = await fetch(base + "/mcp", {
    method: "POST",
    headers: { ...authorization, "Content-Type": "application/json" },
    body: "{",
  });
  assert.equal(malformed.status, 400);
  assert.equal(await malformed.text(), "Invalid request");
  // Legacy SSE transports have separate initialization state and keep the first client alive.
  async function sse() {
    const abort = new AbortController();
    const response = await fetch(base + "/sse", {
      headers: authorization,
      signal: abort.signal,
    });
    assert.equal(response.status, 200);
    const reader = response.body!.getReader();
    let buffer = "",
      endpoint = "";
    const replies = new Map<number, any>();
    const reading = (async () => {
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += new TextDecoder().decode(value);
          let cut;
          while ((cut = buffer.indexOf("\n\n")) >= 0) {
            const event = buffer.slice(0, cut);
            buffer = buffer.slice(cut + 2);
            const data = event
              .split("\n")
              .find((l) => l.startsWith("data: "))
              ?.slice(6);
            if (!data) continue;
            if (event.includes("event: endpoint")) endpoint = data;
            else {
              const message = JSON.parse(data);
              replies.set(message.id, message);
            }
          }
        }
      } catch {}
    })();
    async function until(check: () => boolean) {
      for (let i = 0; i < 200 && !check(); i++)
        await new Promise((resolve) => setTimeout(resolve, 10));
      assert.ok(check(), "SSE result arrives");
    }
    await until(() => !!endpoint);
    return {
      endpoint,
      async call(id: number, method: string, params: any = {}) {
        const result = await fetch(base + endpoint, {
          method: "POST",
          headers: { ...authorization, "Content-Type": "application/json" },
          body: JSON.stringify({ jsonrpc: "2.0", id, method, params }),
        });
        assert.equal(result.status, 202);
        await result.text();
        await until(() => replies.has(id));
        return replies.get(id);
      },
      async close() {
        abort.abort();
        await reader.cancel().catch(() => {});
        await reading;
      },
    };
  }
  const first = await sse(),
    second = await sse();
  try {
    assert.notEqual(first.endpoint, second.endpoint);
    for (const c of [first, second])
      assert.equal(
        (
          await c.call(1, "initialize", {
            protocolVersion: "2025-11-25",
            capabilities: {},
            clientInfo: { name: "sse", version: "1" },
          })
        ).result.serverInfo.name,
        "mcp-msoffice-interop-word",
      );
    assert.equal((await first.call(2, "tools/list")).result.tools.length, 38);
    assert.equal((await second.call(2, "tools/list")).result.tools.length, 38);
  } finally {
    await first.close();
    await second.close();
  }
  const requests = await Promise.all(
    Array.from({ length: 6 }, () =>
      rpc("tools/call", { name: "word_runtimeStatus", arguments: {} }),
    ),
  );
  assert.ok(
    requests.every((r) => r.body.result.structuredContent.singleOperator),
  );
});
