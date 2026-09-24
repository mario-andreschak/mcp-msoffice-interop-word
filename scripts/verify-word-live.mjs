// Explicit acceptance for a licensed, interactive Windows Word installation.
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
if (process.env.WORD_LIVE_TEST !== "1" || process.platform !== "win32")
  throw new Error(
    "Set WORD_LIVE_TEST=1 on a Windows desktop with licensed Word to run this explicit acceptance test.",
  );
if (!process.argv.includes("--native-child")) {
  const child = spawn(
    process.execPath,
    [fileURLToPath(import.meta.url), "--native-child"],
    { stdio: "inherit", windowsHide: true },
  );
  const timer = setTimeout(() => child.kill(), 60000);
  child.once("exit", (code) => {
    clearTimeout(timer);
    process.exitCode = code ?? 1;
  });
} else {
  const { WordService } = await import("../dist/word/word-service.js");
  const root = mkdtempSync(path.join(tmpdir(), "word-live-"));
  process.env.WORD_ALLOWED_ROOTS = JSON.stringify([root]);
  const service = new WordService();
  try {
    service.createDocument();
    service.insertText("Word formatting — ✓");
    service.selectAll();
    service.toggleBold();
    service.toggleItalic();
    service.setParagraphAlignment(1);
    const doc = service.getActiveDocument();
    assert.equal(doc.Content.Font.Bold, -1);
    assert.equal(doc.Content.Font.Italic, -1);
    assert.equal(doc.Content.ParagraphFormat.Alignment, 1);
    service.moveCursorToEnd();
    service.insertText("\n");
    service.addTable(2, 2);
    service.setTableCellText(1, 1, 1, "Unicode — ✓");
    service.applyTableAutoFormat(1, 16);
    const output = path.join(root, "verified.docx");
    service.saveActiveDocumentAs(output, 16);
    assert.equal(readFileSync(output).subarray(0, 2).toString(), "PK");
    service.closeActiveDocument(0);
    console.log(
      "Real Word create/insert/select/font/paragraph/table/save/close acceptance passed.",
    );
  } finally {
    service.quitWord();
    rmSync(root, { recursive: true, force: true });
  }
}
