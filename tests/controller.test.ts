import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { WordController } from "../src/word/controller.js";
async function fixture(t: test.TestContext, timeoutMs = 3000) {
  const dir = await mkdtemp(path.join(tmpdir(), "word-worker-"));
  const worker = path.join(dir, "worker.cjs");
  await writeFile(
    worker,
    `let busy=false;process.on('message',job=>{console.log('native library noise');if(job.args[0]==='hang')return;if(busy){process.send({id:job.id,error:'concurrent'});return;}busy=true;setTimeout(()=>{busy=false;process.send({id:job.id,result:job.args[0]});},50);});process.on('disconnect',()=>process.exit());`,
  );
  const controller = new WordController({ worker, timeoutMs });
  t.after(async () => {
    await controller.close();
    await rm(dir, { recursive: true, force: true });
  });
  return controller;
}
test("native operations are serialized and JSON results survive noisy worker stdout", async (t) => {
  const c = await fixture(t);
  assert.deepEqual(
    await Promise.all([
      c.invoke("insertText", ["one"]),
      c.invoke("insertText", ["two"]),
      c.invoke("insertText", ["three"]),
    ]),
    ["one", "two", "three"],
  );
});
test("queued cancellation leaves the active operation usable", async (t) => {
  const c = await fixture(t);
  const first = c.invoke("insertText", ["one"]);
  const abort = new AbortController();
  const second = c.invoke("insertText", ["two"], abort.signal);
  const rejected = assert.rejects(second, /cancelled/);
  abort.abort();
  await rejected;
  assert.equal(await first, "one");
  assert.equal(await c.invoke("insertText", ["three"]), "three");
});
test("a blocked native call has a deadline and rejects queued work without silently creating a new Word session", async (t) => {
  const c = await fixture(t, 250);
  const first = assert.rejects(c.invoke("insertText", ["hang"]), /deadline/);
  const second = assert.rejects(c.invoke("insertText", ["queued"]), /deadline/);
  await Promise.all([first, second]);
  await assert.rejects(c.invoke("insertText", ["new"]), /restart/);
});
test("active cancellation stops the native worker and shutdown settles pending calls", async (t) => {
  const c = await fixture(t);
  const abort = new AbortController();
  const failed = assert.rejects(
    c.invoke("insertText", ["hang"], abort.signal),
    /cancelled/,
  );
  abort.abort();
  await failed;
  await c.close();
  const d = await fixture(t);
  const pending = assert.rejects(
    d.invoke("insertText", ["hang"]),
    /shutting down/,
  );
  await d.close();
  await pending;
});
