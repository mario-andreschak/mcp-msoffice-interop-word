import { McpServer, type CallToolResult } from "@modelcontextprotocol/server";
import { z } from "zod";
import { callSignal } from "./word/controller.js";
export class ToolRegistrar {
  constructor(readonly server: McpServer) {}
  tool<S extends z.ZodRawShape>(
    name: string,
    description: string,
    shape: S,
    handler: (args: z.output<z.ZodObject<S>>) => Promise<CallToolResult>,
  ) {
    this.server.registerTool(
      name,
      {
        description,
        inputSchema: z.object(shape),
        annotations: {
          readOnlyHint: name.startsWith("word_get"),
          destructiveHint: !name.startsWith("word_get"),
          idempotentHint: name.startsWith("word_get"),
          openWorldHint: true,
        },
      },
      (args, context) =>
        callSignal.run(context?.mcpReq?.signal, () =>
          handler(args as z.output<z.ZodObject<S>>),
        ),
    );
  }
}
