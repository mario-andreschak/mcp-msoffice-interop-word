import http, { type IncomingMessage, type ServerResponse } from "node:http";
import { timingSafeEqual } from "node:crypto";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { toNodeHandler } from "@modelcontextprotocol/node";
import { SSEServerTransport } from "@modelcontextprotocol/server-legacy/sse";
import { createServer } from "./server.js";
function equal(a: string, b: string) {
  const x = Buffer.from(a),
    y = Buffer.from(b);
  return x.length === y.length && timingSafeEqual(x, y);
}
function origin(value: string): string {
  const url = new URL(value);
  if (
    !["http:", "https:"].includes(url.protocol) ||
    url.username ||
    url.password ||
    url.pathname !== "/" ||
    url.search ||
    url.hash
  )
    throw new Error("Invalid allowed origin");
  return url.origin;
}
class BodyError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}
function body(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks: Buffer[] = [];
    const cleanup = () => {
      clearTimeout(timer);
      req.removeListener("data", onData);
      req.removeListener("end", onEnd);
      req.removeListener("error", onError);
      req.removeListener("aborted", onAborted);
    };
    const fail = (error: Error) => {
      cleanup();
      req.resume();
      reject(error);
    };
    const onData = (chunk: Buffer) => {
      size += chunk.length;
      if (size > 2 * 1024 * 1024) {
        fail(new BodyError(413, "Request exceeds 2MiB"));
        return;
      }
      chunks.push(chunk);
    };
    const onEnd = () => {
      cleanup();
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch {
        reject(new BodyError(400, "Invalid request"));
      }
    };
    const onError = () => fail(new BodyError(400, "Invalid request"));
    const onAborted = () => fail(new BodyError(400, "Invalid request"));
    const timer = setTimeout(
      () => fail(new BodyError(408, "Request body timed out")),
      15000,
    );
    req.on("data", onData);
    req.once("end", onEnd);
    req.once("error", onError);
    req.once("aborted", onAborted);
  });
}
export async function startHttp() {
  const host = process.env.HOST ?? "127.0.0.1",
    port = Number(process.env.PORT ?? 3001),
    authToken = process.env.MCP_AUTH_TOKEN ?? "";
  if (!Number.isInteger(port) || port < 0 || port > 65535)
    throw new Error("Invalid PORT");
  if (!/^[\x21-\x7e]{32,256}$/.test(authToken))
    throw new Error("HTTP requires a 32-256 character MCP_AUTH_TOKEN.");
  const publicOrigin = process.env.MCP_PUBLIC_ORIGIN;
  if (
    !["127.0.0.1", "localhost", "::1"].includes(host) &&
    (!publicOrigin || !origin(publicOrigin).startsWith("https://"))
  )
    throw new Error(
      "Remote HTTP requires an HTTPS MCP_PUBLIC_ORIGIN and a TLS reverse proxy.",
    );
  const config = {
    host,
    port,
    authToken,
    allowedOrigins: publicOrigin ? [origin(publicOrigin)] : [],
    allowedHosts: publicOrigin ? [new URL(origin(publicOrigin)).host] : [],
  };
  const factory = createServer;
  const handler = createMcpHandler(factory, { legacy: "stateless" });
  const nodeHandler = toNodeHandler(handler);
  const sessions = new Map<
    string,
    {
      transport: SSEServerTransport;
      server: ReturnType<typeof factory>;
      lastUsed: number;
    }
  >();
  const origins = new Set(config.allowedOrigins.map(origin));
  const hosts = new Set(config.allowedHosts);
  for (const host of hosts)
    if (!/^(\[[a-fA-F0-9:]+\]|[a-zA-Z0-9.-]+)(:[0-9]{1,5})?$/.test(host))
      throw new Error("Invalid allowed host");
  let listenPort = config.port;
  const server = http.createServer(async (req, res) => {
    const reject = (status: number, message: string) => {
      res.writeHead(status, {
        "content-type": "text/plain",
        "cache-control": "no-store",
      });
      res.end(message);
    };
    try {
      const bindHost = config.host.includes(":")
        ? "[" + config.host + "]"
        : config.host;
      const localHosts = ["127.0.0.1", "localhost", "[::1]", bindHost];
      const trustedHosts = new Set([
        ...hosts,
        ...localHosts.map((h) => h + ":" + listenPort),
        ...(listenPort === 80 ? localHosts : []),
      ]);
      if (!trustedHosts.has(req.headers.host ?? ""))
        return reject(421, "Invalid Host");
      const browserOrigin = req.headers.origin;
      const trustedOrigins = new Set([
        ...origins,
        ...localHosts.map((h) => origin("http://" + h + ":" + listenPort)),
      ]);
      if (browserOrigin && !trustedOrigins.has(browserOrigin))
        return reject(403, "Invalid Origin");
      if (browserOrigin) {
        res.setHeader("Access-Control-Allow-Origin", browserOrigin);
        res.setHeader("Vary", "Origin");
      }
      res.setHeader("Cache-Control", "no-store");
      res.setHeader("X-Content-Type-Options", "nosniff");
      const url = new URL(req.url ?? "/", "http://localhost");
      if (req.method === "OPTIONS") {
        res.writeHead(204, {
          "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
          "Access-Control-Allow-Headers":
            "Content-Type, Authorization, MCP-Protocol-Version, Mcp-Method, Mcp-Name, Mcp-Session-Id",
        });
        return res.end();
      }
      if (!equal(req.headers.authorization ?? "", "Bearer " + config.authToken))
        return reject(401, "Bearer authentication required");
      if (url.pathname === "/mcp") {
        if (
          req.method === "POST" &&
          !(req.headers["content-type"] ?? "")
            .toLowerCase()
            .startsWith("application/json")
        )
          return reject(415, "Expected application/json");
        const parsed = req.method === "POST" ? await body(req) : undefined;
        return await nodeHandler(req, res, parsed);
      }
      if (url.pathname === "/sse" && req.method === "GET") {
        if (sessions.size >= 32) return reject(503, "Session limit reached");
        const transport = new SSEServerTransport("/messages", res);
        const mcp = factory();
        sessions.set(transport.sessionId, {
          transport,
          server: mcp,
          lastUsed: Date.now(),
        });
        res.on("close", () => {
          sessions.delete(transport.sessionId);
          void mcp.close();
        });
        return await mcp.connect(transport);
      }
      if (url.pathname === "/messages" && req.method === "POST") {
        const session = sessions.get(url.searchParams.get("sessionId") ?? "");
        if (!session) return reject(404, "Unknown session");
        session.lastUsed = Date.now();
        // The SDK receives the parsed object and never rereads the consumed body.
        return await session.transport.handlePostMessage(
          req,
          res,
          await body(req),
        );
      }
      return reject(404, "Use /mcp or /sse");
    } catch (error) {
      if (!res.headersSent) {
        res.setHeader("Connection", "close");
        reject(
          error instanceof BodyError ? error.status : 400,
          error instanceof BodyError ? error.message : "Invalid request",
        );
      } else res.end();
    }
  });
  server.requestTimeout = 30000;
  server.headersTimeout = 15000;
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(config.port, config.host, () => resolve());
  });
  const address = server.address();
  if (!address || typeof address === "string")
    throw new Error("HTTP address unavailable");
  listenPort = address.port;
  const timer = setInterval(() => {
    for (const [id, s] of sessions)
      if (Date.now() - s.lastUsed > 600000) {
        sessions.delete(id);
        void s.server.close();
      }
  }, 30000);
  timer.unref();
  return {
    server,
    port: listenPort,
    async close() {
      clearInterval(timer);
      await Promise.all([...sessions.values()].map((s) => s.server.close()));
      sessions.clear();
      await handler.close();
      server.closeAllConnections();
      await new Promise<void>((resolve) => server.close(() => resolve()));
    },
  };
}
