import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import express, { Request, Response } from "express";
import http from 'http';

// Basic server info
const serverInfo = {
  name: "mcp-msoffice-interop-word",
  version: "1.0.0",
};

// Create the MCP server instance
const mcpServer = new McpServer(serverInfo, {
  instructions: "MCP Server for interacting with Microsoft Word.",
});

// --- Register Resources and Tools ---
import { registerDocumentTools } from "./tools/document-tools.js";
import { registerTextTools } from "./tools/text-tools.js";
import { registerParagraphTools } from "./tools/paragraph-tools.js";
import { registerTableTools } from "./tools/table-tools.js";
import { registerImageTools } from "./tools/image-tools.js";
import { registerHeaderFooterTools } from "./tools/header-footer-tools.js";
import { registerPageSetupTools } from "./tools/page-setup-tools.js";
// Import other tool/resource registration functions here

registerDocumentTools(mcpServer);
registerTextTools(mcpServer);
registerParagraphTools(mcpServer);
registerTableTools(mcpServer);
registerImageTools(mcpServer);
registerHeaderFooterTools(mcpServer);
registerPageSetupTools(mcpServer);
// Call other registration functions here
// mcpServer.resource(...)

// --- Transport Setup ---

async function startServer() {
  const transportMode = process.env.MCP_TRANSPORT || 'stdio'; // Default to stdio

  if (transportMode === 'stdio') {
    console.log("Starting MCP server with stdio transport...");
    const transport = new StdioServerTransport();
    await mcpServer.connect(transport);
    console.log("MCP server connected via stdio.");
  } else if (transportMode === 'sse') {
    const port = process.env.PORT || 3001;
    const app = express();
    const httpServer = http.createServer(app);

    // Store transports by session ID for multiple connections
    const transports: { [sessionId: string]: SSEServerTransport } = {};

    app.get("/sse", async (_req: Request, res: Response) => {
      console.log("SSE connection requested");
      const transport = new SSEServerTransport('/messages', res);
      transports[transport.sessionId] = transport;
      console.log(`SSE transport created with session ID: ${transport.sessionId}`);

      res.on("close", () => {
        console.log(`SSE connection closed for session ID: ${transport.sessionId}`);
        delete transports[transport.sessionId];
        // Optionally, notify the server instance if needed
        // mcpServer.disconnectClient(transport.sessionId); 
      });

      try {
        await mcpServer.connect(transport);
        console.log(`MCP server connected via SSE for session ID: ${transport.sessionId}`);
      } catch (error) {
        console.error(`Error connecting MCP server via SSE for session ID: ${transport.sessionId}`, error);
        // Ensure response is ended if connection fails
        if (!res.writableEnded) {
          res.status(500).end('Failed to connect MCP server');
        }
      }
    });

    // Middleware to parse JSON bodies for POST requests
    app.use('/messages', express.json({ limit: '4mb' })); // Adjust limit as needed

    app.post("/messages", async (req: Request, res: Response) => {
      const sessionId = req.query.sessionId as string;
      console.log(`Received POST message for session ID: ${sessionId}`);
      const transport = transports[sessionId];
      if (transport) {
        try {
          // Pass the already parsed body if using express.json()
          await transport.handlePostMessage(req, res, req.body);
          console.log(`Handled POST message for session ID: ${sessionId}`);
        } catch (error) {
          console.error(`Error handling POST message for session ID: ${sessionId}`, error);
          if (!res.headersSent) {
             res.status(400).send('Error processing message');
          }
        }
      } else {
        console.warn(`No active SSE transport found for session ID: ${sessionId}`);
        res.status(400).send('No transport found for sessionId');
      }
    });

    httpServer.listen(port, () => {
      console.log(`MCP server listening with SSE transport on port ${port}`);
      console.log(`SSE endpoint: /sse`);
      console.log(`Message endpoint: /messages`);
    });

  } else {
    console.error(`Unsupported MCP_TRANSPORT: ${transportMode}. Use 'stdio' or 'sse'.`);
    process.exit(1);
  }
}

startServer().catch(error => {
  console.error("Failed to start MCP server:", error);
  process.exit(1);
});
