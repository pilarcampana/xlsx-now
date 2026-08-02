"use strict";

// Compiles src/ and test/ to CommonJS in dist-test/, which is what the test
// run loads. See tsconfig.test.json for why the tests do not run on the ESM
// in dist/: nyc measures coverage through Istanbul's `require` hook, and an
// ES module never goes through it.
//
// A plain script rather than a line in package.json because of the last step:
// the package is `"type": "module"`, so the CommonJS just emitted needs its
// own package.json marker next to it or Node reads it as ESM. Writing that
// from a shell one-liner is where quoting stops being portable, and this has
// to run on Windows too.

const { spawnSync } = require("node:child_process");
const { rmSync, writeFileSync } = require("node:fs");
const { join } = require("node:path");

const root = join(__dirname, "..");
const outDir = join(root, "dist-test");

// tsc leaves whatever it emitted last time in place, and a renamed test would
// otherwise keep running from its old name.
rmSync(outDir, { recursive: true, force: true });

const built = spawnSync(
    process.execPath,
    [require.resolve("typescript/bin/tsc"), "-p", join(root, "tsconfig.test.json")],
    { stdio: "inherit" }
);

if (built.status === 0) {
    writeFileSync(join(outDir, "package.json"), '{ "type": "commonjs" }\n');
} else {
    process.exitCode = built.status === null ? 1 : built.status;
}
