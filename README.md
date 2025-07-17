[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/mario-andreschak-mcp-msoffice-interop-word-badge.png)](https://mseep.ai/app/mario-andreschak-mcp-msoffice-interop-word)

# MCP Office Interop Word Server

This project implements a [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that allows interaction with Microsoft Word documents using COM Interop on Windows.

It provides MCP tools to perform common Word processing tasks programmatically.

## Features

*   Wraps common Microsoft Word operations via COM Interop (`winax`).
*   Exposes functionality as MCP tools.
*   Supports both `stdio` and `sse` transports for MCP communication.
*   Built with TypeScript and uses the `@modelcontextprotocol/sdk`.

## Prerequisites

*   Node.js (v18 or later recommended)
*   npm
*   Microsoft Word installed on a Windows machine.

## Installation

1.  Clone the repository or download the source code.
2.  Navigate to the project directory in your terminal.
3.  Install dependencies:
    ```bash
    npm install
    ```

## Building

To compile the TypeScript code to JavaScript:

```bash
npm run build
```

This will output the compiled files to the `dist` directory.

## Running the Server

The server can run using two different MCP transports: `stdio` or `sse`.

### stdio Transport

This is the default mode. It's suitable for local clients that communicate via standard input/output.

```bash
npm start
```

or

```bash
node dist/index.js
```

Connect your MCP client (e.g., MCP Inspector) using the stdio method, pointing to the `node dist/index.js` command.

### SSE (Server-Sent Events) Transport

This mode uses HTTP and Server-Sent Events, suitable for web-based or remote clients.

**PowerShell:**

```powershell
$env:MCP_TRANSPORT="sse"; npm start
```

**Bash / Cmd:**

```bash
MCP_TRANSPORT=sse npm start
```

The server will start an HTTP server, typically on port 3001 (or the port specified by the `PORT` environment variable).

*   **SSE Endpoint:** `http://localhost:3001/sse`
*   **Message Endpoint (for client POSTs):** `http://localhost:3001/messages`

Connect your MCP client using the SSE method, providing the SSE endpoint URL.

## Available Tools

The server exposes the following tools (tool names are prefixed with `word_`):

**Document Operations:**

*   `word_createDocument`: Creates a new, blank Word document.
*   `word_openDocument`: Opens an existing document.
    *   `filePath` (string): Absolute path to the document.
*   `word_saveActiveDocument`: Saves the currently active document.
*   `word_saveActiveDocumentAs`: Saves the active document to a new path/format.
    *   `filePath` (string): Absolute path to save to.
    *   `fileFormat` (number, optional): Numeric `WdSaveFormat` value (e.g., 16 for docx, 17 for pdf).
*   `word_closeActiveDocument`: Closes the active document.
    *   `saveChanges` (number, optional): `WdSaveOptions` value (0=No, -1=Yes, -2=Prompt). Default: 0.

**Text Manipulation:**

*   `word_insertText`: Inserts text at the selection.
    *   `text` (string): Text to insert.
*   `word_deleteText`: Deletes text relative to the selection.
    *   `count` (number, optional): Number of units to delete (default: 1). Positive=forward, negative=backward.
    *   `unit` (number, optional): `WdUnits` value (1=Char, 2=Word, etc.). Default: 1.
*   `word_findAndReplace`: Finds and replaces text.
    *   `findText` (string): Text to find.
    *   `replaceText` (string): Replacement text.
    *   `matchCase` (boolean, optional): Default: false.
    *   `matchWholeWord` (boolean, optional): Default: false.
    *   `replaceAll` (boolean, optional): Default: true.
*   `word_toggleBold`: Toggles bold formatting for the selection.
*   `word_toggleItalic`: Toggles italic formatting for the selection.
*   `word_toggleUnderline`: Toggles underline formatting for the selection.
    *   `underlineStyle` (number, optional): `WdUnderline` value (default: 1=Single).

**Paragraph Formatting:**

*   `word_setParagraphAlignment`: Sets paragraph alignment.
    *   `alignment` (number): `WdParagraphAlignment` value (0=Left, 1=Center, 2=Right, 3=Justify).
*   `word_setParagraphLeftIndent`: Sets left indent.
    *   `indentPoints` (number): Indent value in points.
*   `word_setParagraphRightIndent`: Sets right indent.
    *   `indentPoints` (number): Indent value in points.
*   `word_setParagraphFirstLineIndent`: Sets first line/hanging indent.
    *   `indentPoints` (number): Indent value in points (positive=indent, negative=hanging).
*   `word_setParagraphSpaceBefore`: Sets space before paragraphs.
    *   `spacePoints` (number): Space value in points.
*   `word_setParagraphSpaceAfter`: Sets space after paragraphs.
    *   `spacePoints` (number): Space value in points.
*   `word_setParagraphLineSpacing`: Sets line spacing.
    *   `lineSpacingRule` (number): `WdLineSpacing` value (0=Single, 1=1.5, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple).
    *   `lineSpacingValue` (number, optional): Value needed for rules 3, 4, 5.

**Table Operations:**

*   `word_addTable`: Adds a table at the selection.
    *   `numRows` (number): Number of rows.
    *   `numCols` (number): Number of columns.
*   `word_setTableCellText`: Sets text in a table cell.
    *   `tableIndex` (number): 1-based table index.
    *   `rowIndex` (number): 1-based row index.
    *   `colIndex` (number): 1-based column index.
    *   `text` (string): Text to set.
*   `word_insertTableRow`: Inserts a row into a table.
    *   `tableIndex` (number): 1-based table index.
    *   `beforeRowIndex` (number, optional): Insert before this 1-based row index (or at end if omitted).
*   `word_insertTableColumn`: Inserts a column into a table.
    *   `tableIndex` (number): 1-based table index.
    *   `beforeColIndex` (number, optional): Insert before this 1-based column index (or at right end if omitted).
*   `word_applyTableAutoFormat`: Applies a style to a table.
    *   `tableIndex` (number): 1-based table index.
    *   `formatName` (string | number): Style name or `WdTableFormat` value.

**Image Operations:**

*   `word_insertPicture`: Inserts an inline picture.
    *   `filePath` (string): Absolute path to the image file.
    *   `linkToFile` (boolean, optional): Default: false.
    *   `saveWithDocument` (boolean, optional): Default: true.
*   `word_setInlinePictureSize`: Resizes an inline picture.
    *   `shapeIndex` (number): 1-based index of the inline shape.
    *   `heightPoints` (number): Height in points (-1 or 0 to auto-size).
    *   `widthPoints` (number): Width in points (-1 or 0 to auto-size).
    *   `lockAspectRatio` (boolean, optional): Default: true.

**Header/Footer Operations:**

*   `word_setHeaderFooterText`: Sets text in a header or footer.
    *   `text` (string): Text content.
    *   `isHeader` (boolean): True for header, false for footer.
    *   `sectionIndex` (number, optional): 1-based section index (default: 1).
    *   `headerFooterType` (number, optional): `WdHeaderFooterIndex` value (1=Primary, 2=FirstPage, 3=EvenPages). Default: 1.

**Page Setup Operations:**

*   `word_setPageMargins`: Sets page margins.
    *   `topPoints` (number): Top margin in points.
    *   `bottomPoints` (number): Bottom margin in points.
    *   `leftPoints` (number): Left margin in points.
    *   `rightPoints` (number): Right margin in points.
*   `word_setPageOrientation`: Sets page orientation.
    *   `orientation` (number): `WdOrientation` value (0=Portrait, 1=Landscape).
*   `word_setPaperSize`: Sets paper size.
    *   `paperSize` (number): `WdPaperSize` value (e.g., 1=Letter, 8=A4).

## Notes

*   This server requires Microsoft Word to be installed and accessible via COM Interop on the machine where the server runs.
*   Error handling for COM operations is basic. Robust production use might require more detailed error checking and recovery.
*   Word object model constants (like `WdSaveFormat`, `WdUnits`, etc.) are represented by their numeric values in the tool arguments. You may need to refer to the Word VBA documentation for specific values.
