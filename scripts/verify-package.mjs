import { spawnSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";
import assert from "node:assert/strict";
const npm = process.env.npm_execpath;
if (!npm) throw new Error("Run through npm run test:package.");
const repo = process.cwd();
function run(args, cwd = repo) {
  const p = spawnSync(process.execPath, [npm, ...args], {
    cwd,
    encoding: "utf8",
    timeout: 180000,
  });
  if (p.status !== 0)
    throw new Error(`${args[0]} failed: ${p.stderr} ${p.stdout}`);
  return p.stdout;
}
const pack = JSON.parse(run(["pack", "--ignore-scripts", "--json"]));
const tarball = path.resolve(pack[0].filename);
const temp = await mkdtemp(path.join(tmpdir(), "word-package-"));
console.log("Production installation:", temp);
try {
  await writeFile(
    path.join(temp, "package.json"),
    JSON.stringify({
      name: "word-production-gate",
      version: "1.0.0",
      private: true,
    }),
  );
  run(
    [
      "install",
      "--omit=dev",
      ...(process.platform === "win32" ? [] : ["--ignore-scripts"]),
      tarball,
    ],
    temp,
  );
  const installed = path.join(temp, "node_modules/mcp-office-interop");
  const manifest = JSON.parse(
    await readFile(path.join(installed, "package.json"), "utf8"),
  );
  const entry = path.resolve(installed, manifest.bin["mcp-office-interop"]);
  assert.ok(entry.startsWith(installed + path.sep));
  const p = spawnSync(
    process.execPath,
    [path.join(repo, "scripts/verify-protocol.mjs"), entry],
    { encoding: "utf8", timeout: 60000 },
  );
  process.stdout.write(p.stdout ?? "");
  process.stderr.write(p.stderr ?? "");
  assert.equal(p.status, 0);
  const audit = JSON.parse(run(["audit", "--omit=dev", "--json"], temp));
  assert.equal(audit.metadata.vulnerabilities.total, 0);
  console.log("Installed production package and vulnerability audit passed.");
} finally {
  await rm(temp, {
    recursive: true,
    force: true,
    maxRetries: 5,
    retryDelay: 200,
  });
}
