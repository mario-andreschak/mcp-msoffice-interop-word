import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Insert Text ---
const insertTextSchema = z.object({
  text: z.string().describe("The text to insert at the current cursor position or over the selection."),
});

async function insertTextTool(args: z.infer<typeof insertTextSchema>): Promise<CallToolResult> {
  try {
    await wordService.insertText(args.text);
    return {
      content: [{ type: "text", text: "Successfully inserted text." }],
    };
  } catch (error: any) {
    console.error("Error in insertTextTool:", error);
    return {
      content: [{ type: "text", text: `Failed to insert text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Delete Text ---
const deleteTextSchema = z.object({
  count: z.number().int().optional().default(1).describe("Number of units to delete. Positive deletes forward/after selection, negative deletes backward/before selection. Default is 1."),
  unit: z.number().int().optional().default(1).describe("Unit to delete (1=Character, 2=Word, 3=Sentence, 4=Paragraph). Default is 1 (Character)."),
});

async function deleteTextTool(args: z.infer<typeof deleteTextSchema>): Promise<CallToolResult> {
  try {
    await wordService.deleteText(args.count, args.unit);
    const unitMap: { [key: number]: string } = { 1: "character(s)", 2: "word(s)", 3: "sentence(s)", 4: "paragraph(s)" };
    const unitName = unitMap[args.unit ?? 1] ?? `unit ${args.unit}`;
    const direction = (args.count ?? 1) >= 0 ? "forward" : "backward";
    return {
      content: [{ type: "text", text: `Successfully deleted ${Math.abs(args.count ?? 1)} ${unitName} ${direction}.` }],
    };
  } catch (error: any) {
    console.error("Error in deleteTextTool:", error);
    return {
      content: [{ type: "text", text: `Failed to delete text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Find and Replace Text ---
const findReplaceSchema = z.object({
  findText: z.string().describe("The text to search for."),
  replaceText: z.string().describe("The text to replace occurrences with."),
  matchCase: z.boolean().optional().default(false).describe("Perform a case-sensitive search."),
  matchWholeWord: z.boolean().optional().default(false).describe("Only find whole word matches."),
  replaceAll: z.boolean().optional().default(true).describe("Replace all occurrences (true) or only the first one (false)."),
});

async function findReplaceTool(args: z.infer<typeof findReplaceSchema>): Promise<CallToolResult> {
  try {
    const found = await wordService.findAndReplace(
      args.findText,
      args.replaceText,
      args.matchCase,
      args.matchWholeWord,
      args.replaceAll
    );
    const message = found
      ? `Successfully found and replaced text "${args.findText}".`
      : `Text "${args.findText}" not found.`;
    return {
      content: [{ type: "text", text: message }],
      isError: !found && args.replaceAll, // Consider it an error only if replaceAll was true and nothing was found
    };
  } catch (error: any) {
    console.error("Error in findReplaceTool:", error);
    return {
      content: [{ type: "text", text: `Failed to find and replace text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Toggle Bold ---
const toggleBoldSchema = z.object({});

async function toggleBoldTool(): Promise<CallToolResult> {
  try {
    await wordService.toggleBold();
    return {
      content: [{ type: "text", text: "Toggled bold formatting for the selection." }],
    };
  } catch (error: any) {
    console.error("Error in toggleBoldTool:", error);
    return {
      content: [{ type: "text", text: `Failed to toggle bold: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Toggle Italic ---
const toggleItalicSchema = z.object({});

async function toggleItalicTool(): Promise<CallToolResult> {
  try {
    await wordService.toggleItalic();
    return {
      content: [{ type: "text", text: "Toggled italic formatting for the selection." }],
    };
  } catch (error: any) {
    console.error("Error in toggleItalicTool:", error);
    return {
      content: [{ type: "text", text: `Failed to toggle italic: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Toggle Underline ---
const toggleUnderlineSchema = z.object({
    underlineStyle: z.number().int().optional().default(1).describe("Optional: Underline style (WdUnderline enum value, e.g., 1=Single, 4=Double). Default is 1."),
});

async function toggleUnderlineTool(args: z.infer<typeof toggleUnderlineSchema>): Promise<CallToolResult> {
  try {
    await wordService.toggleUnderline(args.underlineStyle);
    return {
      content: [{ type: "text", text: "Toggled underline formatting for the selection." }],
    };
  } catch (error: any) {
    console.error("Error in toggleUnderlineTool:", error);
    return {
      content: [{ type: "text", text: `Failed to toggle underline: ${error.message}` }],
      isError: true,
    };
  }
}


// --- Register Tools ---
export function registerTextTools(server: McpServer) {
  server.tool(
    "word_insertText",
    "Inserts the given text at the current selection in the active Word document.",
    insertTextSchema.shape,
    insertTextTool
  );
  server.tool(
    "word_deleteText",
    "Deletes text relative to the current selection in the active Word document.",
    deleteTextSchema.shape,
    deleteTextTool
  );
  server.tool(
    "word_findAndReplace",
    "Finds and replaces text within the active Word document.",
    findReplaceSchema.shape,
    findReplaceTool
  );
  server.tool(
    "word_toggleBold",
    "Toggles bold formatting for the current selection.",
    toggleBoldSchema.shape,
    toggleBoldTool
  );
  server.tool(
    "word_toggleItalic",
    "Toggles italic formatting for the current selection.",
    toggleItalicSchema.shape,
    toggleItalicTool
  );
  server.tool(
    "word_toggleUnderline",
    "Toggles underline formatting for the current selection.",
    toggleUnderlineSchema.shape,
    toggleUnderlineTool
  );
}
