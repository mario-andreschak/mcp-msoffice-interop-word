import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Move Cursor to Start ---
const moveCursorToStartSchema = z.object({});

async function moveCursorToStartTool(): Promise<CallToolResult> {
  try {
    await wordService.moveCursorToStart();
    return {
      content: [{ type: "text", text: "Successfully moved cursor to the start of the document." }],
    };
  } catch (error: any) {
    console.error("Error in moveCursorToStartTool:", error);
    return {
      content: [{ type: "text", text: `Failed to move cursor to start: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Move Cursor to End ---
const moveCursorToEndSchema = z.object({});

async function moveCursorToEndTool(): Promise<CallToolResult> {
  try {
    await wordService.moveCursorToEnd();
    return {
      content: [{ type: "text", text: "Successfully moved cursor to the end of the document." }],
    };
  } catch (error: any) {
    console.error("Error in moveCursorToEndTool:", error);
    return {
      content: [{ type: "text", text: `Failed to move cursor to end: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Move Cursor ---
const moveCursorSchema = z.object({
  unit: z.number().int().min(1).max(12).default(1).describe("Unit to move by (1=Character, 2=Word, 3=Sentence, 4=Paragraph, 5=Line, 6=Story, etc.)"),
  count: z.number().int().describe("Number of units to move. Positive moves forward, negative moves backward."),
  extend: z.boolean().optional().default(false).describe("Whether to extend the selection (true) or move the insertion point (false)."),
});

async function moveCursorTool(args: z.infer<typeof moveCursorSchema>): Promise<CallToolResult> {
  try {
    await wordService.moveCursor(args.unit, args.count, args.extend);
    
    const unitMap: { [key: number]: string } = { 
      1: "character(s)", 
      2: "word(s)", 
      3: "sentence(s)", 
      4: "paragraph(s)",
      5: "line(s)",
      6: "story",
      7: "screen",
      8: "section",
      9: "column",
      10: "row",
      11: "window",
      12: "cell"
    };
    
    const unitName = unitMap[args.unit] ?? `unit(s)`;
    const direction = args.count >= 0 ? "forward" : "backward";
    const action = args.extend ? "extended selection" : "moved cursor";
    
    return {
      content: [{ type: "text", text: `Successfully ${action} ${Math.abs(args.count)} ${unitName} ${direction}.` }],
    };
  } catch (error: any) {
    console.error("Error in moveCursorTool:", error);
    return {
      content: [{ type: "text", text: `Failed to move cursor: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Select All ---
const selectAllSchema = z.object({});

async function selectAllTool(): Promise<CallToolResult> {
  try {
    await wordService.selectAll();
    return {
      content: [{ type: "text", text: "Successfully selected the entire document." }],
    };
  } catch (error: any) {
    console.error("Error in selectAllTool:", error);
    return {
      content: [{ type: "text", text: `Failed to select all: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Select Paragraph ---
const selectParagraphSchema = z.object({
  paragraphIndex: z.number().int().min(1).describe("1-based index of the paragraph to select."),
});

async function selectParagraphTool(args: z.infer<typeof selectParagraphSchema>): Promise<CallToolResult> {
  try {
    await wordService.selectParagraph(args.paragraphIndex);
    return {
      content: [{ type: "text", text: `Successfully selected paragraph ${args.paragraphIndex}.` }],
    };
  } catch (error: any) {
    console.error("Error in selectParagraphTool:", error);
    return {
      content: [{ type: "text", text: `Failed to select paragraph: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Collapse Selection ---
const collapseSelectionSchema = z.object({
  toStart: z.boolean().optional().default(true).describe("If true, collapse to start; if false, collapse to end."),
});

async function collapseSelectionTool(args: z.infer<typeof collapseSelectionSchema>): Promise<CallToolResult> {
  try {
    await wordService.collapseSelection(args.toStart);
    const position = args.toStart ? "start" : "end";
    return {
      content: [{ type: "text", text: `Successfully collapsed selection to its ${position}.` }],
    };
  } catch (error: any) {
    console.error("Error in collapseSelectionTool:", error);
    return {
      content: [{ type: "text", text: `Failed to collapse selection: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Get Selection Text ---
const getSelectionTextSchema = z.object({});

async function getSelectionTextTool(): Promise<CallToolResult> {
  try {
    const text = await wordService.getSelectionText();
    return {
      content: [
        { type: "text", text: "Current selection text:" },
        { type: "text", text: text || "(empty selection)" }
      ],
    };
  } catch (error: any) {
    console.error("Error in getSelectionTextTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get selection text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Get Selection Info ---
const getSelectionInfoSchema = z.object({});

async function getSelectionInfoTool(): Promise<CallToolResult> {
  try {
    const info = await wordService.getSelectionInfo();
    
    // Map selection type to a human-readable string
    const typeMap: { [key: number]: string } = {
      0: "None",
      1: "Normal",
      2: "Column",
      3: "Row",
      4: "Block",
      5: "InlineShape",
      6: "Shape",
      7: "Frame"
    };
    
    const typeStr = typeMap[info.type] || `Unknown (${info.type})`;
    
    return {
      content: [
        { type: "text", text: "Selection Information:" },
        { type: "text", text: `- Text: ${info.text || "(empty)"}` },
        { type: "text", text: `- Start Position: ${info.start}` },
        { type: "text", text: `- End Position: ${info.end}` },
        { type: "text", text: `- Is Active: ${info.isActive}` },
        { type: "text", text: `- Selection Type: ${typeStr}` }
      ],
    };
  } catch (error: any) {
    console.error("Error in getSelectionInfoTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get selection info: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerCursorSelectionTools(server: McpServer) {
  server.tool(
    "word_moveCursorToStart",
    "Moves the cursor to the start of the document.",
    moveCursorToStartSchema.shape,
    moveCursorToStartTool
  );
  
  server.tool(
    "word_moveCursorToEnd",
    "Moves the cursor to the end of the document.",
    moveCursorToEndSchema.shape,
    moveCursorToEndTool
  );
  
  server.tool(
    "word_moveCursor",
    "Moves the cursor by the specified unit and count.",
    moveCursorSchema.shape,
    moveCursorTool
  );
  
  server.tool(
    "word_selectAll",
    "Selects the entire document.",
    selectAllSchema.shape,
    selectAllTool
  );
  
  server.tool(
    "word_selectParagraph",
    "Selects a specific paragraph by index.",
    selectParagraphSchema.shape,
    selectParagraphTool
  );
  
  server.tool(
    "word_collapseSelection",
    "Collapses the current selection to its start or end point.",
    collapseSelectionSchema.shape,
    collapseSelectionTool
  );
  
  server.tool(
    "word_getSelectionText",
    "Gets the text of the current selection.",
    getSelectionTextSchema.shape,
    getSelectionTextTool
  );
  
  server.tool(
    "word_getSelectionInfo",
    "Gets detailed information about the current selection.",
    getSelectionInfoSchema.shape,
    getSelectionInfoTool
  );
}
