import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
test(
  "Windows installs and executes the actual winax COM addon",
  { skip: process.platform !== "win32" },
  () => {
    const winax = createRequire(import.meta.url)("../vendor/winax/index.js");
    const dictionary = new winax.Object("Scripting.Dictionary", {
      activate: false,
    });
    try {
      dictionary.Add("Unicode — ✓", 9999998);
      assert.equal(dictionary.Count, 1);
      assert.equal(dictionary.Item("Unicode — ✓"), 9999998);
      dictionary.RemoveAll();
      assert.equal(dictionary.Count, 0);
    } finally {
      winax.release(dictionary);
    }
  },
);
