import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";
import path from 'path';

// --- Tool: Insert Picture ---
const insertPictureSchema = z.object({
  filePath: z.string().describe("The absolute path to the image file to insert."),
  linkToFile: z.boolean().optional().default(false).describe("Link to the file instead of embedding it."),
  saveWithDocument: z.boolean().optional().default(true).describe("Save the linked image with the document."),
});

async function insertPictureTool(args: z.infer<typeof insertPictureSchema>): Promise<CallToolResult> {
  try {
    const absolutePath = path.resolve(args.filePath); // Ensure absolute path
    await wordService.insertPicture(absolutePath, args.linkToFile, args.saveWithDocument);
    return {
      content: [{ type: "text", text: `Successfully inserted picture from: ${absolutePath}` }],
    };
  } catch (error: any) {
    console.error("Error in insertPictureTool:", error);
    return {
      content: [{ type: "text", text: `Failed to insert picture: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Inline Picture Size ---
const setInlinePictureSizeSchema = z.object({
  shapeIndex: z.number().int().min(1).describe("The 1-based index of the inline picture in the document's InlineShapes collection."),
  heightPoints: z.number().describe("Desired height in points. Use -1 or 0 to auto-size based on width and aspect ratio."),
  widthPoints: z.number().describe("Desired width in points. Use -1 or 0 to auto-size based on height and aspect ratio."),
  lockAspectRatio: z.boolean().optional().default(true).describe("Maintain the picture's aspect ratio when resizing."),
});

async function setInlinePictureSizeTool(args: z.infer<typeof setInlinePictureSizeSchema>): Promise<CallToolResult> {
  try {
    if (args.heightPoints <= 0 && args.widthPoints <= 0) {
        return {
            content: [{ type: "text", text: "No size change specified (height and width were not positive values)." }],
            isError: false, // Not an error, just no action taken
        };
    }
    await wordService.setInlinePictureSize(args.shapeIndex, args.heightPoints, args.widthPoints, args.lockAspectRatio);
    return {
      content: [{ type: "text", text: `Successfully resized inline picture at index ${args.shapeIndex}.` }],
    };
  } catch (error: any) {
    console.error("Error in setInlinePictureSizeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to resize inline picture: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerImageTools(server: McpServer) {
  server.tool(
    "word_insertPicture",
    "Inserts a picture from a file path into the active document at the selection point.",
    insertPictureSchema.shape,
    insertPictureTool
  );
  server.tool(
    "word_setInlinePictureSize",
    "Resizes an inline picture (identified by its index) in the active document.",
    setInlinePictureSizeSchema.shape,
    setInlinePictureSizeTool
  );
  // Add tools for floating shapes (positioning, etc.) if needed later
}
