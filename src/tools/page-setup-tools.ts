import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Set Page Margins ---
const setPageMarginsSchema = z.object({
  topPoints: z.number().min(0).describe("Top margin in points."),
  bottomPoints: z.number().min(0).describe("Bottom margin in points."),
  leftPoints: z.number().min(0).describe("Left margin in points."),
  rightPoints: z.number().min(0).describe("Right margin in points."),
});

async function setPageMarginsTool(args: z.infer<typeof setPageMarginsSchema>): Promise<CallToolResult> {
  try {
    await wordService.setPageMargins(args.topPoints, args.bottomPoints, args.leftPoints, args.rightPoints);
    return {
      content: [{ type: "text", text: `Successfully set page margins (Top: ${args.topPoints}, Bottom: ${args.bottomPoints}, Left: ${args.leftPoints}, Right: ${args.rightPoints} points).` }],
    };
  } catch (error: any) {
    console.error("Error in setPageMarginsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set page margins: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Page Orientation ---
const setPageOrientationSchema = z.object({
  orientation: z.number().int().min(0).max(1).describe("Page orientation (0=Portrait, 1=Landscape). Corresponds to WdOrientation enum."),
});

async function setPageOrientationTool(args: z.infer<typeof setPageOrientationSchema>): Promise<CallToolResult> {
  try {
    await wordService.setPageOrientation(args.orientation);
    const orientationName = args.orientation === 0 ? "Portrait" : "Landscape";
    return {
      content: [{ type: "text", text: `Successfully set page orientation to ${orientationName}.` }],
    };
  } catch (error: any) {
    console.error("Error in setPageOrientationTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set page orientation: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paper Size ---
const setPaperSizeSchema = z.object({
  paperSize: z.number().int().describe("Paper size value corresponding to WdPaperSize enum (e.g., 1=Letter, 8=A4)."),
});

async function setPaperSizeTool(args: z.infer<typeof setPaperSizeSchema>): Promise<CallToolResult> {
  try {
    await wordService.setPaperSize(args.paperSize);
    // We could add a map for common paper size names if needed
    return {
      content: [{ type: "text", text: `Successfully set paper size (Enum value: ${args.paperSize}).` }],
    };
  } catch (error: any) {
    console.error("Error in setPaperSizeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set paper size: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerPageSetupTools(server: McpServer) {
  server.tool(
    "word_setPageMargins",
    "Sets the top, bottom, left, and right margins for the active document.",
    setPageMarginsSchema.shape,
    setPageMarginsTool
  );
  server.tool(
    "word_setPageOrientation",
    "Sets the page orientation (Portrait or Landscape) for the active document.",
    setPageOrientationSchema.shape,
    setPageOrientationTool
  );
  server.tool(
    "word_setPaperSize",
    "Sets the paper size (e.g., Letter, A4) for the active document using WdPaperSize enum values.",
    setPaperSizeSchema.shape,
    setPaperSizeTool
  );
}
