// Assembles dist/umd/xlsx-now.umd.js from the two halves of the UMD build:
// the single AMD file `tsc -p tsconfig.umd.json` emits out of src/core, and the
// hand-written envelope in src/umd/wrapper.js. `tsc` did the module work; this
// script only splices the pieces, shifts the source map to match, and marks the
// output directory as CommonJS.
//
// Paths are resolved from the repo root — run via `npm run build`.
import { readFile, writeFile, rm } from 'node:fs/promises';
import { resolve } from 'node:path';

/** Line of src/umd/wrapper.js the AMD modules take the place of. */
const INJECTION_MARKER = '    // <<< tsc AMD output is injected here by scripts/build-umd.ts >>>';

const WRAPPER_FILE = resolve('src/umd/wrapper.js');
const AMD_FILE = resolve('dist/umd/xlsx-now.amd.js');
const OUT_FILE = resolve('dist/umd/xlsx-now.umd.js');
const OUT_MAP_NAME = 'xlsx-now.umd.js.map';

/**
 * The fields of the source map this script touches. Everything else — notably
 * `sources`, which stays valid because the map does not change directory — is
 * carried over untouched.
 */
interface SourceMap {
    file: string;
    mappings: string;
    [field: string]: unknown;
}

/**
 * `tsc` ends the file with the reference to its own map and no trailing
 * newline, which would comment out whatever is appended after it.
 */
function stripSourceMappingUrl(code: string): string {
    return code.replace(/\n?\/\/# sourceMappingURL=.*$/, '\n');
}

async function main(): Promise<void> {
    const wrapper = await readFile(WRAPPER_FILE, 'utf8');
    const halves = wrapper.split(`${INJECTION_MARKER}\n`);
    const [head, tail] = halves;
    if (halves.length !== 2 || head === undefined || tail === undefined) {
        throw new Error(`${WRAPPER_FILE} must contain the injection marker exactly once:\n${INJECTION_MARKER}`);
    }

    const amd = stripSourceMappingUrl(await readFile(AMD_FILE, 'utf8'));
    await writeFile(OUT_FILE, `${head}${amd}${tail}//# sourceMappingURL=${OUT_MAP_NAME}\n`);

    // Source map mappings are line-oriented and separated by ";", so pushing
    // the generated code down by the lines of the wrapper's head is just a
    // matter of prepending that many empty lines.
    const map = JSON.parse(await readFile(`${AMD_FILE}.map`, 'utf8')) as SourceMap;
    const headLines = head.split('\n').length - 1;
    map.file = 'xlsx-now.umd.js';
    map.mappings = ';'.repeat(headLines) + map.mappings;
    await writeFile(resolve('dist/umd', OUT_MAP_NAME), JSON.stringify(map));

    // The package is "type": "module", which would make Node parse this .js as
    // ESM and choke on the UMD wrapper's `module`/`exports`. A package.json
    // scoped to this directory opts just it back into CommonJS, while the file
    // keeps the .js extension a <script> tag needs.
    await writeFile(resolve('dist/umd/package.json'), `{ "type": "commonjs" }\n`);

    // The AMD half is an intermediate artifact; dist/umd is published whole.
    await rm(AMD_FILE);
    await rm(`${AMD_FILE}.map`);
}

await main();
