import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
if (process.platform === "win32") {
  const require = createRequire(import.meta.url);
  try {
    const directory = fileURLToPath(new URL("../vendor/winax", import.meta.url));
    const gyp = require.resolve("node-gyp/bin/node-gyp.js");
    const build = spawnSync(process.execPath, [gyp, "rebuild", "--directory", directory], {
      stdio: "inherit", timeout: 600_000, windowsHide: true,
    });
    if (build.error || build.status !== 0) throw new Error("Native build failed");
    require("../vendor/winax/index.js");
  } catch {
    console.error("Word native addon did not build. Use Node22.22.2+/24.15+, npm11+, Python and Visual Studio Desktop development with C++; reinstall with --foreground-scripts to see compiler diagnostics. Keep optional dependencies enabled for the pinned native build tool. Linux discovery does not require Word COM.");
    process.exitCode = 1;
  }
}
