import test from "node:test";
import assert from "node:assert/strict";
import {
  mkdtempSync,
  writeFileSync,
  rmSync,
  mkdirSync,
  symlinkSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { WordService } from "../src/word/word-service.js";
import { confineRead, confineWrite } from "../src/word/paths.js";
function fixture() {
  const calls: any[] = [];
  const table: any = {
    AutoFormat: (...args: any[]) => calls.push(["AutoFormat", ...args]),
  };
  const font: any = { Bold: 0, Italic: 0, Underline: 0 };
  const selection: any = {
    Font: font,
    ParagraphFormat: {},
    TypeText: (text: string) => calls.push(["TypeText", text]),
    InlineShapes: {
      AddPicture: (...args: any[]) => calls.push(["Picture", ...args]),
    },
  };
  const doc: any = {
    Path: "",
    FullName: "",
    Activate: () => calls.push(["Activate"]),
    Close: (value: number) => calls.push(["Close", value]),
    Save: () => calls.push(["Save"]),
    SaveAs2: (...args: any[]) => calls.push(["SaveAs", ...args]),
    Tables: { Count: 1, Item: () => table },
  };
  // Reject accidental Promise assimilation of native COM objects.
  const com = (value: any) =>
    new Proxy(value, {
      get(target, key) {
        if (key === "then") throw new Error("COM object cannot be awaited");
        return target[key];
      },
    });
  const app: any = {
    Visible: false,
    Documents: {
      Count: 1,
      Add: () => com(doc),
      Open: (...args: any[]) => {
        calls.push(["Open", ...args]);
        return com(doc);
      },
    },
    Options: {},
    Selection: selection,
    ActiveDocument: com(doc),
    Quit: () => calls.push(["Quit"]),
  };
  return {
    service: new WordService(() => com(app)),
    calls,
    app,
    doc,
    table,
    font,
    selection,
  };
}
test("native COM references stay synchronous; creation activates the document and disables macros/link updates", () => {
  const f = fixture();
  f.service.createDocument();
  f.service.insertText("Hello — ✓");
  f.service.toggleBold();
  f.service.toggleItalic();
  f.service.toggleUnderline();
  f.service.toggleUnderline();
  f.service.setParagraphAlignment(2);
  assert.equal(f.app.AutomationSecurity, 3);
  assert.equal(f.app.Options.UpdateLinksAtOpen, false);
  assert.deepEqual(f.calls, [["Activate"], ["TypeText", "Hello — ✓"]]);
  assert.equal(f.font.Bold, 9999998);
  assert.equal(f.font.Italic, 9999998);
  assert.equal(f.font.Underline, 0);
  assert.equal(f.selection.ParagraphFormat.Alignment, 2);
});
test("named table styles use Style and numeric AutoFormat receives real boolean arguments", () => {
  const f = fixture();
  f.service.applyTableAutoFormat(1, "Table Grid");
  assert.equal(f.table.Style, "Table Grid");
  f.service.applyTableAutoFormat(1, 16);
  f.service.applyTableAutoFormat(1, 16, 1 | 32 | 256);
  assert.deepEqual(f.calls, [
    ["AutoFormat", 16],
    [
      "AutoFormat",
      16,
      true,
      false,
      false,
      false,
      true,
      true,
      false,
      false,
      false,
    ],
  ]);
});
test("close reports COM failure and rejects modal save prompts", () => {
  const f = fixture();
  assert.throws(() => f.service.closeActiveDocument(-2), /prompts/);
  f.doc.Close = () => {
    throw new Error("RPC rejected");
  };
  assert.throws(() => f.service.closeActiveDocument(0), /Could not close/);
});
test("file operations use configured real directories and reject escapes, links and overwrite", (t) => {
  const root = mkdtempSync(path.join(tmpdir(), "word-paths-"));
  const allowed = path.join(root, "allowed"),
    outside = path.join(root, "outside");
  mkdirSync(allowed);
  mkdirSync(outside);
  const old = process.env.WORD_ALLOWED_ROOTS;
  process.env.WORD_ALLOWED_ROOTS = JSON.stringify([allowed]);
  t.after(() => {
    if (old === undefined) delete process.env.WORD_ALLOWED_ROOTS;
    else process.env.WORD_ALLOWED_ROOTS = old;
    rmSync(root, { recursive: true, force: true });
  });
  const doc = path.join(allowed, "hello.docx");
  writeFileSync(doc, "fixture");
  writeFileSync(path.join(outside, "private.docx"), "outside");
  assert.equal(confineRead(doc), doc);
  assert.throws(
    () => confineRead(path.join(outside, "private.docx")),
    /outside/,
  );
  assert.throws(() => confineRead("relative.docx"), /absolute/);
  assert.throws(() => confineWrite(doc), /already exists/);
  symlinkSync(
    outside,
    path.join(allowed, "linked"),
    process.platform === "win32" ? "junction" : "dir",
  );
  assert.throws(
    () => confineRead(path.join(allowed, "linked/private.docx")),
    /outside/,
  );
  assert.throws(
    () => confineWrite(path.join(allowed, "linked/new.docx")),
    /outside/,
  );
  const f = fixture();
  f.service.openDocument(doc);
  assert.deepEqual(f.calls[0], ["Open", doc, false, false, false]);
  assert.throws(() => f.service.saveActiveDocument(), /no path/);
  f.service.saveActiveDocumentAs(path.join(allowed, "new.docx"), 16);
  assert.deepEqual(f.calls.at(-1), [
    "SaveAs",
    path.join(allowed, "new.docx"),
    16,
  ]);
});

test("an unavailable owned Word instance does not attach to or silently create another application",()=>{
 const f=fixture();f.service.createDocument();Object.defineProperty(f.app,"Visible",{get(){throw new Error("Word is busy");}});assert.throws(()=>f.service.getWordApplication(),/No replacement instance/);assert.equal(f.calls.filter(c=>c[0]==="Activate").length,1);
});
