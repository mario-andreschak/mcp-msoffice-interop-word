# MCP Office Interop Word

Microsoft Word automation for one trusted operator. Supports the published MCP 2026-07-28 protocol through the stable TypeScript SDK 2, legacy 2025-11-25 stdio/Streamable HTTP, and separately initialized legacy HTTP+SSE clients.

Use Node 22 or 24 and licensed Microsoft Word in an interactive Windows session. `npm ci && npm run build` installs the optional native `winax 3.6.9` addon on Windows. Windows CI must execute the real addon; its installation failure is not silently accepted. Linux/macOS can run protocol discovery and `word_runtimeStatus`, but document tools return an explicit platform error.

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["C:/path/to/mcp-msoffice-interop-word/dist/index.js"],
      "env": { "WORD_ALLOWED_ROOTS": "[\"C:/Documents/WordWorkspace\"]" }
    }
  }
}
```

The npm tarball also provides the `mcp-office-interop` command. Stdio is the default and writes only JSON-RPC to stdout. Native COM calls run serially in a child process so blocked Office dialogs cannot freeze MCP. Each call has a 30-second deadline including queue time; the queue holds at most 32 waiting calls and argument/result messages are bounded to 1 MiB.

The server creates its own Word application, activates new/opened documents, and uses its current selection. It never attaches to an arbitrary existing Word instance. A timeout or cancellation of an active call stops the worker and marks its outcome uncertain; restart the server before further Word operations. The Word window remains available to recover unsaved work. Server shutdown does not globally terminate Word processes or discard those documents.

`WORD_ALLOWED_ROOTS` is a JSON array of existing absolute local directories; default is the server's working directory. Read/open/image and save destinations must remain within those real directories, including symlink/junction resolution. Network/device paths and alternate data streams are rejected. Save As requires a new destination; `word_saveActiveDocument` saves the active existing document. New unsaved documents need Save As first. Save prompts are rejected. Word macros and automatic link updates are disabled for the owned instance. These controls do not turn Office into a sandbox for untrusted documents.

For HTTP, set `MCP_TRANSPORT=http`, `MCP_AUTH_TOKEN` to a 32–256 character printable secret, and optionally `PORT` (default 3001). Binding defaults to `127.0.0.1`. Every sensitive route requires the bearer header; URL tokens are not accepted. Use `/mcp` for modern/legacy Streamable HTTP or `/sse` and `/messages` for old SSE clients. `MCP_TRANSPORT=sse` retains the same endpoints. Exact Host/Origin checks reject hostile and opaque origins. Remote binding requires an HTTPS `MCP_PUBLIC_ORIGIN` and a TLS reverse proxy configured to preserve the expected Host. All authenticated connections share this operator's Word session; deploy separate processes/workspaces for separate operators.

`word_runtimeStatus` reports platform/native-addon availability without opening Word. It explicitly leaves Word installation unverified.

## Validation and issues

`npm test` exercises HTTP authentication/Host/Origin, modern and legacy catalogs, independent SSE clients, COM selection/format calls, path confinement and native worker lifecycle. `npm run test:protocol` executes the compiled or supplied installed entrypoint in both protocol eras, including errors and EOF. Windows tests execute `Scripting.Dictionary` through the actual native addon; this proves COM interoperability, not licensed Word availability.

For actual Word acceptance, after building, run `WORD_LIVE_TEST=1 node scripts/verify-word-live.mjs` from an interactive Windows desktop (set the variable using PowerShell's `$env:WORD_LIVE_TEST='1'`). It creates only its own temporary document and checks text/selection/bold/italic/paragraph/table/DOCX-save/close, with a 60-second outer deadline. This live test is separate from ordinary CI.

Issue 1: new/opened document activation and current application selection are explicit; native references never pass through Promise assimilation or IPC serialization; named table styles use `Table.Style`, and numeric `AutoFormat` options use the documented boolean positions. These repairs have deterministic regressions. Full issue reproduction and Word-version acceptance still require the explicit live test. [Microsoft Table.Style](https://learn.microsoft.com/en-us/office/vba/api/word.table.style), [AutoFormat](https://learn.microsoft.com/en-us/office/vba/api/word.table.autoformat), [Font.Bold](https://learn.microsoft.com/en-us/office/vba/api/Word.Font.Bold).

Issue 2: the Dockerfile runs the MCP discovery/status/error paths as a non-root Linux user. It cannot provide Microsoft Word COM; document automation requires Windows. Glama ownership/listing validation is an external action and is not claimed by a successful image build.

As of September 2026 this uses supported Node 22/24 and SDK 2 with actual modern/legacy boundary tests. End-of-2026 fitness remains conditional on passing native Word acceptance and keeping the runtime/addon/Office patches current. [winax](https://github.com/durs/node-activex), [Microsoft AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity).

## Available Tools

The server exposes the following tools (tool names are prefixed with `word_`):

**Document Operations:**

- `word_createDocument`: Creates a new, blank Word document.
- `word_openDocument`: Opens an existing document.
  - `filePath` (string): Absolute path to the document.
- `word_saveActiveDocument`: Saves the currently active document.
- `word_saveActiveDocumentAs`: Saves the active document to a new path/format.
  - `filePath` (string): Absolute path to save to.
  - `fileFormat` (number, optional): Numeric `WdSaveFormat` value (e.g., 16 for docx, 17 for pdf).
- `word_closeActiveDocument`: Closes the active document.
  - `saveChanges` (number, optional): `WdSaveOptions` value (0=No, -1=Yes). Default: 0; modal prompts are rejected.

**Text Manipulation:**

- `word_insertText`: Inserts text at the selection.
  - `text` (string): Text to insert.
- `word_deleteText`: Deletes text relative to the selection.
  - `count` (number, optional): Number of units to delete (default: 1). Positive=forward, negative=backward.
  - `unit` (number, optional): `WdUnits` value (1=Char, 2=Word, etc.). Default: 1.
- `word_findAndReplace`: Finds and replaces text.
  - `findText` (string): Text to find.
  - `replaceText` (string): Replacement text.
  - `matchCase` (boolean, optional): Default: false.
  - `matchWholeWord` (boolean, optional): Default: false.
  - `replaceAll` (boolean, optional): Default: true.
- `word_toggleBold`: Toggles bold formatting for the selection.
- `word_toggleItalic`: Toggles italic formatting for the selection.
- `word_toggleUnderline`: Toggles underline formatting for the selection.
  - `underlineStyle` (number, optional): `WdUnderline` value (default: 1=Single).

**Paragraph Formatting:**

- `word_setParagraphAlignment`: Sets paragraph alignment.
  - `alignment` (number): `WdParagraphAlignment` value (0=Left, 1=Center, 2=Right, 3=Justify).
- `word_setParagraphLeftIndent`: Sets left indent.
  - `indentPoints` (number): Indent value in points.
- `word_setParagraphRightIndent`: Sets right indent.
  - `indentPoints` (number): Indent value in points.
- `word_setParagraphFirstLineIndent`: Sets first line/hanging indent.
  - `indentPoints` (number): Indent value in points (positive=indent, negative=hanging).
- `word_setParagraphSpaceBefore`: Sets space before paragraphs.
  - `spacePoints` (number): Space value in points.
- `word_setParagraphSpaceAfter`: Sets space after paragraphs.
  - `spacePoints` (number): Space value in points.
- `word_setParagraphLineSpacing`: Sets line spacing.
  - `lineSpacingRule` (number): `WdLineSpacing` value (0=Single, 1=1.5, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple).
  - `lineSpacingValue` (number, optional): Value needed for rules 3, 4, 5.

**Table Operations:**

- `word_addTable`: Adds a table at the selection.
  - `numRows` (number): Number of rows.
  - `numCols` (number): Number of columns.
- `word_setTableCellText`: Sets text in a table cell.
  - `tableIndex` (number): 1-based table index.
  - `rowIndex` (number): 1-based row index.
  - `colIndex` (number): 1-based column index.
  - `text` (string): Text to set.
- `word_insertTableRow`: Inserts a row into a table.
  - `tableIndex` (number): 1-based table index.
  - `beforeRowIndex` (number, optional): Insert before this 1-based row index (or at end if omitted).
- `word_insertTableColumn`: Inserts a column into a table.
  - `tableIndex` (number): 1-based table index.
  - `beforeColIndex` (number, optional): Insert before this 1-based column index (or at right end if omitted).
- `word_applyTableAutoFormat`: Applies a style to a table.
  - `tableIndex` (number): 1-based table index.
  - `formatName` (string | number): Style name or `WdTableFormat` value.

**Image Operations:**

- `word_insertPicture`: Inserts an inline picture.
  - `filePath` (string): Absolute path to the image file.
  - `linkToFile` (boolean, optional): Default: false.
  - `saveWithDocument` (boolean, optional): Default: true.
- `word_setInlinePictureSize`: Resizes an inline picture.
  - `shapeIndex` (number): 1-based index of the inline shape.
  - `heightPoints` (number): Height in points (-1 or 0 to auto-size).
  - `widthPoints` (number): Width in points (-1 or 0 to auto-size).
  - `lockAspectRatio` (boolean, optional): Default: true.

**Header/Footer Operations:**

- `word_setHeaderFooterText`: Sets text in a header or footer.
  - `text` (string): Text content.
  - `isHeader` (boolean): True for header, false for footer.
  - `sectionIndex` (number, optional): 1-based section index (default: 1).
  - `headerFooterType` (number, optional): `WdHeaderFooterIndex` value (1=Primary, 2=FirstPage, 3=EvenPages). Default: 1.

**Page Setup Operations:**

- `word_setPageMargins`: Sets page margins.
  - `topPoints` (number): Top margin in points.
  - `bottomPoints` (number): Bottom margin in points.
  - `leftPoints` (number): Left margin in points.
  - `rightPoints` (number): Right margin in points.
- `word_setPageOrientation`: Sets page orientation.
  - `orientation` (number): `WdOrientation` value (0=Portrait, 1=Landscape).
- `word_setPaperSize`: Sets paper size.
  - `paperSize` (number): `WdPaperSize` value (e.g., 1=Letter, 8=A4).

### Windows native build toolchain

Use npm 11 or newer with Node 22/24, Python and the Visual Studio Desktop development with C++ workload. Node 22’s bundled older npm/node-gyp does not recognize Visual Studio 2026. Run `npm install --global npm@11.19.1` before installing this package. Native installation now fails clearly if `winax` could not build; use `npm install --foreground-scripts` to expose compiler diagnostics. Linux discovery does not load Word COM.

The package includes the MIT-licensed winax3.6.9 native runtime source with a narrow V8 optional-member detection repair. Upstream’s current release otherwise fails to compile on Node22. Its source integrity, exclusions and exact change are recorded in `vendor/winax/PROVENANCE.md`; no Office macro fixtures or prebuilt binaries are shipped. A pinned node-gyp13.0.2 builds the addon during Windows installation. Keep optional dependencies enabled so this compiler driver is installed. Python/Visual Studio remain operator prerequisites; native Windows CI executes actual COM on Node22 and24.
