#!/usr/bin/env node
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { createServer } from "./server.js";
import { controller } from "./word/controller.js";
import { startHttp } from "./http.js";
let closeHttp: (() => Promise<void>) | undefined,
  closing = false;
async function close() {
  if (closing) return;
  closing = true;
  await controller.close();
  await closeHttp?.();
}
process.once("SIGINT", () => {
  void close().then(() => process.exit(0));
});
process.once("SIGTERM", () => {
  void close().then(() => process.exit(0));
});
async function main() {
  const mode = process.env.MCP_TRANSPORT ?? "stdio";
  if (mode === "stdio") {
    await serveStdio(createServer);
    process.stdin.once("end", () => {
      void close();
    });
  } else if (mode === "http" || mode === "sse") {
    const http = await startHttp();
    closeHttp = () => http.close();
    console.error(
      `Word MCP listening on ${http.port}; /mcp modern/legacy and /sse legacy endpoints.`,
    );
  } else throw new Error("MCP_TRANSPORT must be stdio, http or sse.");
}
main().catch(() => {
  console.error(
    "Unable to start Word MCP. Check transport, port, bearer token and allowed HTTP deployment configuration.",
  );
  void close().finally(() => {
    process.exitCode = 1;
  });
});
