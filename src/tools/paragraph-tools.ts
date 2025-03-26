import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";

// --- Tool: Set Paragraph Alignment ---
const setAlignmentSchema = z.object({
  alignment: z.number().int().min(0).max(3).describe("Alignment type (0=Left, 1=Center, 2=Right, 3=Justify). Corresponds to WdParagraphAlignment enum."),
});

async function setAlignmentTool(args: z.infer<typeof setAlignmentSchema>): Promise<CallToolResult> {
  try {
    await wordService.setParagraphAlignment(args.alignment);
    const alignmentMap: { [key: number]: string } = { 0: "Left", 1: "Center", 2: "Right", 3: "Justify" };
    return {
      content: [{ type: "text", text: `Successfully set paragraph alignment to ${alignmentMap[args.alignment]}.` }],
    };
  } catch (error: any) {
    console.error("Error in setAlignmentTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set paragraph alignment: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paragraph Left Indent ---
const setLeftIndentSchema = z.object({
  indentPoints: z.number().describe("Left indentation value in points."),
});

async function setLeftIndentTool(args: z.infer<typeof setLeftIndentSchema>): Promise<CallToolResult> {
  try {
    await wordService.setParagraphLeftIndent(args.indentPoints);
    return {
      content: [{ type: "text", text: `Successfully set left indent to ${args.indentPoints} points.` }],
    };
  } catch (error: any) {
    console.error("Error in setLeftIndentTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set left indent: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paragraph Right Indent ---
const setRightIndentSchema = z.object({
  indentPoints: z.number().describe("Right indentation value in points."),
});

async function setRightIndentTool(args: z.infer<typeof setRightIndentSchema>): Promise<CallToolResult> {
  try {
    await wordService.setParagraphRightIndent(args.indentPoints);
    return {
      content: [{ type: "text", text: `Successfully set right indent to ${args.indentPoints} points.` }],
    };
  } catch (error: any) {
    console.error("Error in setRightIndentTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set right indent: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paragraph First Line Indent ---
const setFirstLineIndentSchema = z.object({
    indentPoints: z.number().describe("First line indentation in points (positive for indent, negative for hanging indent)."),
});

async function setFirstLineIndentTool(args: z.infer<typeof setFirstLineIndentSchema>): Promise<CallToolResult> {
    try {
        await wordService.setParagraphFirstLineIndent(args.indentPoints);
        const indentType = args.indentPoints >= 0 ? "indent" : "hanging indent";
        return {
            content: [{ type: "text", text: `Successfully set first line ${indentType} to ${Math.abs(args.indentPoints)} points.` }],
        };
    } catch (error: any) {
        console.error("Error in setFirstLineIndentTool:", error);
        return {
            content: [{ type: "text", text: `Failed to set first line indent: ${error.message}` }],
            isError: true,
        };
    }
}

// --- Tool: Set Paragraph Space Before ---
const setSpaceBeforeSchema = z.object({
  spacePoints: z.number().min(0).describe("Space before paragraph in points."),
});

async function setSpaceBeforeTool(args: z.infer<typeof setSpaceBeforeSchema>): Promise<CallToolResult> {
  try {
    await wordService.setParagraphSpaceBefore(args.spacePoints);
    return {
      content: [{ type: "text", text: `Successfully set space before paragraph to ${args.spacePoints} points.` }],
    };
  } catch (error: any) {
    console.error("Error in setSpaceBeforeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set space before paragraph: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paragraph Space After ---
const setSpaceAfterSchema = z.object({
  spacePoints: z.number().min(0).describe("Space after paragraph in points."),
});

async function setSpaceAfterTool(args: z.infer<typeof setSpaceAfterSchema>): Promise<CallToolResult> {
  try {
    await wordService.setParagraphSpaceAfter(args.spacePoints);
    return {
      content: [{ type: "text", text: `Successfully set space after paragraph to ${args.spacePoints} points.` }],
    };
  } catch (error: any) {
    console.error("Error in setSpaceAfterTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set space after paragraph: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Paragraph Line Spacing ---
const setLineSpacingSchema = z.object({
  lineSpacingRule: z.number().int().min(0).max(5).describe("Line spacing rule (0=Single, 1=1.5, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple). Corresponds to WdLineSpacing enum."),
  lineSpacingValue: z.number().optional().describe("Required value (in points or multiplier) if rule is AtLeast(3), Exactly(4), or Multiple(5)."),
});

async function setLineSpacingTool(args: z.infer<typeof setLineSpacingSchema>): Promise<CallToolResult> {
  try {
    if (args.lineSpacingRule >= 3 && args.lineSpacingValue === undefined) {
        throw new Error("lineSpacingValue is required when lineSpacingRule is AtLeast, Exactly, or Multiple.");
    }
    await wordService.setParagraphLineSpacing(args.lineSpacingRule, args.lineSpacingValue);
    const ruleMap: { [key: number]: string } = { 0: "Single", 1: "1.5 Lines", 2: "Double", 3: "At Least", 4: "Exactly", 5: "Multiple" };
    let message = `Successfully set line spacing rule to ${ruleMap[args.lineSpacingRule]}.`;
    if (args.lineSpacingValue !== undefined && args.lineSpacingRule >= 3) {
        const unit = args.lineSpacingRule === 5 ? 'x' : ' points';
        message += ` Value: ${args.lineSpacingValue}${unit}.`;
    }
    return {
      content: [{ type: "text", text: message }],
    };
  } catch (error: any) {
    console.error("Error in setLineSpacingTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set line spacing: ${error.message}` }],
      isError: true,
    };
  }
}


// --- Register Tools ---
export function registerParagraphTools(server: McpServer) {
  server.tool(
    "word_setParagraphAlignment",
    "Sets the alignment for the selected paragraph(s).",
    setAlignmentSchema.shape,
    setAlignmentTool
  );
  server.tool(
    "word_setParagraphLeftIndent",
    "Sets the left indent for the selected paragraph(s).",
    setLeftIndentSchema.shape,
    setLeftIndentTool
  );
  server.tool(
    "word_setParagraphRightIndent",
    "Sets the right indent for the selected paragraph(s).",
    setRightIndentSchema.shape,
    setRightIndentTool
  );
  server.tool(
    "word_setParagraphFirstLineIndent",
    "Sets the first line indent (or hanging indent) for the selected paragraph(s).",
    setFirstLineIndentSchema.shape,
    setFirstLineIndentTool
  );
  server.tool(
    "word_setParagraphSpaceBefore",
    "Sets the spacing before the selected paragraph(s).",
    setSpaceBeforeSchema.shape,
    setSpaceBeforeTool
  );
  server.tool(
    "word_setParagraphSpaceAfter",
    "Sets the spacing after the selected paragraph(s).",
    setSpaceAfterSchema.shape,
    setSpaceAfterTool
  );
  server.tool(
    "word_setParagraphLineSpacing",
    "Sets the line spacing for the selected paragraph(s).",
    setLineSpacingSchema.shape,
    setLineSpacingTool
  );
}
