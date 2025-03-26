import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Set Header/Footer Text ---
const setHeaderFooterTextSchema = z.object({
  text: z.string().describe("The text content to set in the header or footer."),
  isHeader: z.boolean().describe("True to modify the header, False to modify the footer."),
  sectionIndex: z.number().int().min(1).optional().default(1).describe("The 1-based index of the document section (default is 1)."),
  headerFooterType: z.number().int().min(1).max(3).optional().default(1).describe("Type of header/footer (1=Primary, 2=First Page, 3=Even Pages). Default is 1 (Primary). Corresponds to WdHeaderFooterIndex enum."),
});

async function setHeaderFooterTextTool(args: z.infer<typeof setHeaderFooterTextSchema>): Promise<CallToolResult> {
  try {
    await wordService.setHeaderFooterText(args.sectionIndex, args.headerFooterType, args.isHeader, args.text);
    const typeMap: { [key: number]: string } = { 1: "Primary", 2: "First Page", 3: "Even Pages" };
    const location = args.isHeader ? "header" : "footer";
    const typeName = typeMap[args.headerFooterType];
    return {
      content: [{ type: "text", text: `Successfully set text for ${typeName} ${location} in section ${args.sectionIndex}.` }],
    };
  } catch (error: any) {
    console.error("Error in setHeaderFooterTextTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set header/footer text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerHeaderFooterTools(server: McpServer) {
  server.tool(
    "word_setHeaderFooterText",
    "Sets the text content for a specific header or footer in a given section.",
    setHeaderFooterTextSchema.shape,
    setHeaderFooterTextTool
  );
  // Add more tools for header/footer operations if needed (e.g., add page numbers, fields)
}
