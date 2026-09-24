import { z } from "zod";
import { ToolRegistrar as McpServer } from "../register.js";
import { CallToolResult } from "@modelcontextprotocol/server";
import { wordService } from "../word/controller.js";
import path from "path"; // Import path for potential path manipulation

// --- Tool: Create New Document ---
const createDocumentSchema = z.object({}); // No arguments needed

async function createDocumentTool(): Promise<CallToolResult> {
  try {
    await wordService.createDocument();
    return {
      content: [
        { type: "text", text: "Successfully created a new Word document." },
      ],
    };
  } catch (error: any) {
    console.error("Error in createDocumentTool:", error);
    return {
      content: [
        { type: "text", text: `Failed to create document: ${error.message}` },
      ],
      isError: true,
    };
  }
}

// --- Tool: Open Document ---
const openDocumentSchema = z.object({
  filePath: z
    .string()
    .max(32768)
    .refine(path.isAbsolute, "Use an absolute local path")
    .describe("The absolute path to the Word document to open."),
});

async function openDocumentTool(
  args: z.infer<typeof openDocumentSchema>,
): Promise<CallToolResult> {
  try {
    // Consider validating or resolving the path if necessary
    const absolutePath = path.resolve(args.filePath); // Example: ensure absolute path
    await wordService.openDocument(absolutePath);
    return {
      content: [
        { type: "text", text: `Successfully opened document: ${absolutePath}` },
      ],
    };
  } catch (error: any) {
    console.error("Error in openDocumentTool:", error);
    return {
      content: [
        {
          type: "text",
          text: `Failed to open document '${args.filePath}': ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

// --- Tool: Save Active Document ---
const saveDocumentSchema = z.object({}); // No arguments needed

async function saveDocumentTool(): Promise<CallToolResult> {
  try {
    await wordService.saveActiveDocument();
    return {
      content: [
        { type: "text", text: "Successfully saved the active document." },
      ],
    };
  } catch (error: any) {
    console.error("Error in saveDocumentTool:", error);
    // Provide more context if possible (e.g., if no active doc)
    return {
      content: [
        {
          type: "text",
          text: `Failed to save active document: ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

// --- Tool: Save Active Document As ---
const saveDocumentAsSchema = z.object({
  filePath: z
    .string()
    .max(32768)
    .refine(path.isAbsolute, "Use an absolute local path")
    .describe("The absolute path to save the document to."),
  fileFormat: z
    .optional(z.number().int().min(0).max(24))
    .describe(
      "Optional: The numeric value corresponding to Word's WdSaveFormat enum (e.g., 16 for docx, 17 for pdf).",
    ),
});

async function saveDocumentAsTool(
  args: z.infer<typeof saveDocumentAsSchema>,
): Promise<CallToolResult> {
  try {
    const absolutePath = path.resolve(args.filePath);
    await wordService.saveActiveDocumentAs(absolutePath, args.fileFormat);
    return {
      content: [
        {
          type: "text",
          text: `Successfully saved document as: ${absolutePath}`,
        },
      ],
    };
  } catch (error: any) {
    console.error("Error in saveDocumentAsTool:", error);
    return {
      content: [
        {
          type: "text",
          text: `Failed to save document as '${args.filePath}': ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

// --- Tool: Close Active Document ---
const closeDocumentSchema = z.object({
  saveChanges: z
    .union([z.literal(0), z.literal(-1)])
    .optional()
    .describe(
      "Save changes before close: 0=No, -1=Yes. Prompts are unsupported.",
    ),
});

async function closeDocumentTool(
  args: z.infer<typeof closeDocumentSchema>,
): Promise<CallToolResult> {
  try {
    await wordService.closeActiveDocument(args.saveChanges);
    return {
      content: [
        { type: "text", text: "Successfully closed the active document." },
      ],
    };
  } catch (error: any) {
    console.error("Error in closeDocumentTool:", error);
    // Check if the error is because no document was active
    if (error.message.includes("No active document")) {
      return {
        content: [{ type: "text", text: "No active document to close." }],
        isError: false, // Not necessarily an error in this context
      };
    }
    return {
      content: [
        {
          type: "text",
          text: `Failed to close active document: ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerDocumentTools(server: McpServer) {
  server.tool(
    "word_createDocument",
    "Creates a new, blank Word document.",
    createDocumentSchema.shape, // Pass the shape for McpServer
    createDocumentTool,
  );
  server.tool(
    "word_openDocument",
    "Opens an existing Word document from the specified file path.",
    openDocumentSchema.shape,
    openDocumentTool,
  );
  server.tool(
    "word_saveActiveDocument",
    "Saves the currently active Word document.",
    saveDocumentSchema.shape,
    saveDocumentTool,
  );
  server.tool(
    "word_saveActiveDocumentAs",
    "Saves the currently active Word document to a new file path and/or format.",
    saveDocumentAsSchema.shape,
    saveDocumentAsTool,
  );
  server.tool(
    "word_closeActiveDocument",
    "Closes the currently active Word document, optionally saving changes.",
    closeDocumentSchema.shape,
    closeDocumentTool,
  );
}
