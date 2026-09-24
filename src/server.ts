import { createRequire } from "node:module";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/server";
import { ToolRegistrar } from "./register.js";
import { registerCursorSelectionTools } from "./tools/cursor-selection-tools.js";
import { registerDocumentTools } from "./tools/document-tools.js";
import { registerHeaderFooterTools } from "./tools/header-footer-tools.js";
import { registerImageTools } from "./tools/image-tools.js";
import { registerPageSetupTools } from "./tools/page-setup-tools.js";
import { registerParagraphTools } from "./tools/paragraph-tools.js";
import { registerTableTools } from "./tools/table-tools.js";
import { registerTextTools } from "./tools/text-tools.js";
export function createServer() {
  const server = new McpServer(
    { name: "mcp-msoffice-interop-word", version: "1.0.0" },
    {
      instructions:
        "Microsoft Word automation for one trusted operator. Requires Windows and installed Word. File access is confined to WORD_ALLOWED_ROOTS. Native operations are serialized with deadlines; a timed-out worker requires restart, and unsaved Word windows are retained for recovery.",
    },
  );
  const registrar = new ToolRegistrar(server);
  registerCursorSelectionTools(registrar);
  registerDocumentTools(registrar);
  registerHeaderFooterTools(registrar);
  registerImageTools(registrar);
  registerPageSetupTools(registrar);
  registerParagraphTools(registrar);
  registerTableTools(registrar);
  registerTextTools(registrar);
  server.registerTool(
    "word_runtimeStatus",
    {
      description:
        "Report the runtime platform and native addon availability without opening Word. Word installation and interactive acceptance are separate checks.",
      inputSchema: z.object({}),
      annotations: {
        readOnlyHint: true,
        idempotentHint: true,
        destructiveHint: false,
        openWorldHint: false,
      },
    },
    async () => {
      let nativeAddonAvailable = false;
      if (process.platform === "win32")
        try {
          createRequire(import.meta.url)("../vendor/winax/index.js");
          nativeAddonAvailable = true;
        } catch {}
      const status = {
        platform: process.platform,
        nodeVersion: process.versions.node,
        nativeAddonAvailable,
        wordInstallation: "not checked",
        requires:
          "Windows with licensed Microsoft Word in an interactive user session",
        singleOperator: true,
      };
      return {
        content: [{ type: "text", text: JSON.stringify(status) }],
        structuredContent: status,
      };
    },
  );
  return server;
}
