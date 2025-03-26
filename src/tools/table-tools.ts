import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Add Table ---
const addTableSchema = z.object({
  numRows: z.number().int().min(1).describe("Number of rows for the new table."),
  numCols: z.number().int().min(1).describe("Number of columns for the new table."),
  // Optional: Add defaultTableBehavior and autoFitBehavior if needed, using numeric values for enums
});

async function addTableTool(args: z.infer<typeof addTableSchema>): Promise<CallToolResult> {
  try {
    await wordService.addTable(args.numRows, args.numCols);
    return {
      content: [{ type: "text", text: `Successfully added a ${args.numRows}x${args.numCols} table.` }],
    };
  } catch (error: any) {
    console.error("Error in addTableTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add table: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Table Cell Text ---
const setTableCellTextSchema = z.object({
  tableIndex: z.number().int().min(1).describe("The 1-based index of the table in the document."),
  rowIndex: z.number().int().min(1).describe("The 1-based index of the row within the table."),
  colIndex: z.number().int().min(1).describe("The 1-based index of the column within the table."),
  text: z.string().describe("The text to set in the specified cell."),
});

async function setTableCellTextTool(args: z.infer<typeof setTableCellTextSchema>): Promise<CallToolResult> {
  try {
    await wordService.setTableCellText(args.tableIndex, args.rowIndex, args.colIndex, args.text);
    return {
      content: [{ type: "text", text: `Successfully set text in table ${args.tableIndex}, cell (${args.rowIndex}, ${args.colIndex}).` }],
    };
  } catch (error: any) {
    console.error("Error in setTableCellTextTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set cell text: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Insert Table Row ---
const insertTableRowSchema = z.object({
    tableIndex: z.number().int().min(1).describe("The 1-based index of the table."),
    beforeRowIndex: z.number().int().min(1).optional().describe("Optional: 1-based index of the row to insert before. If omitted, adds row to the end."),
});

async function insertTableRowTool(args: z.infer<typeof insertTableRowSchema>): Promise<CallToolResult> {
    try {
        await wordService.insertTableRow(args.tableIndex, args.beforeRowIndex);
        const position = args.beforeRowIndex ? `before row ${args.beforeRowIndex}` : "at the end";
        return {
            content: [{ type: "text", text: `Successfully inserted row into table ${args.tableIndex} ${position}.` }],
        };
    } catch (error: any) {
        console.error("Error in insertTableRowTool:", error);
        return {
            content: [{ type: "text", text: `Failed to insert table row: ${error.message}` }],
            isError: true,
        };
    }
}

// --- Tool: Insert Table Column ---
const insertTableColumnSchema = z.object({
    tableIndex: z.number().int().min(1).describe("The 1-based index of the table."),
    beforeColIndex: z.number().int().min(1).optional().describe("Optional: 1-based index of the column to insert before. If omitted, adds column to the right end."),
});

async function insertTableColumnTool(args: z.infer<typeof insertTableColumnSchema>): Promise<CallToolResult> {
    try {
        await wordService.insertTableColumn(args.tableIndex, args.beforeColIndex);
         const position = args.beforeColIndex ? `before column ${args.beforeColIndex}` : "at the right end";
        return {
            content: [{ type: "text", text: `Successfully inserted column into table ${args.tableIndex} ${position}.` }],
        };
    } catch (error: any) {
        console.error("Error in insertTableColumnTool:", error);
        return {
            content: [{ type: "text", text: `Failed to insert table column: ${error.message}` }],
            isError: true,
        };
    }
}

// --- Tool: Apply Table AutoFormat ---
const applyTableAutoFormatSchema = z.object({
    tableIndex: z.number().int().min(1).describe("The 1-based index of the table."),
    formatName: z.union([z.string(), z.number()]).describe("Name of the table style (e.g., 'Table Grid') or a numeric WdTableFormat enum value."),
    // applyFormatting: z.number().int().optional().describe("Optional: Bitmask flags (WdTableFormatApply) indicating which parts of the format to apply.")
});

async function applyTableAutoFormatTool(args: z.infer<typeof applyTableAutoFormatSchema>): Promise<CallToolResult> {
    try {
        // We'll omit applyFormatting for simplicity, letting Word use defaults
        await wordService.applyTableAutoFormat(args.tableIndex, args.formatName);
        return {
            content: [{ type: "text", text: `Successfully applied format '${args.formatName}' to table ${args.tableIndex}.` }],
        };
    } catch (error: any) {
        console.error("Error in applyTableAutoFormatTool:", error);
        return {
            content: [{ type: "text", text: `Failed to apply table format: ${error.message}` }],
            isError: true,
        };
    }
}


// --- Register Tools ---
export function registerTableTools(server: McpServer) {
  server.tool(
    "word_addTable",
    "Adds a new table at the current selection point.",
    addTableSchema.shape,
    addTableTool
  );
  server.tool(
    "word_setTableCellText",
    "Sets the text content of a specific cell in a table.",
    setTableCellTextSchema.shape,
    setTableCellTextTool
  );
  server.tool(
    "word_insertTableRow",
    "Inserts a new row into a specified table.",
    insertTableRowSchema.shape,
    insertTableRowTool
  );
  server.tool(
    "word_insertTableColumn",
    "Inserts a new column into a specified table.",
    insertTableColumnSchema.shape,
    insertTableColumnTool
  );
  server.tool(
    "word_applyTableAutoFormat",
    "Applies a predefined style or autoformat to a table.",
    applyTableAutoFormatSchema.shape,
    applyTableAutoFormatTool
  );
}
