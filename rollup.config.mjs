// Builds the UMD bundles. Rollup runs on the JS that `tsc` already emitted
// (`npm run build` runs `tsc` first), so TypeScript compilation stays in one
// place and this config only bundles this package's own modules into one file
// and converts the module format.
//
// There are two, because they cover different environments: the core bundle is
// the writer, which runs anywhere, and the browser bundle is only the helpers
// that reach for the DOM. A page loads the core one and, if it wants the
// helpers, the browser one on top.
import { writeFile } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';

/**
 * The UMD bundle keeps the `.js` extension so it can be served and consumed
 * from a plain `<script>` tag, but this package is `"type": "module"` — which
 * would make Node parse that `.js` as ESM and `require()` of it fail on the
 * UMD wrapper's `module`/`exports` references. A `package.json` scoped to the
 * output directory opts just that directory back into CommonJS.
 */
const commonjsMarker = {
    name: 'umd-commonjs-marker',
    async writeBundle(options) {
        await writeFile(join(dirname(options.file), 'package.json'), `{ "type": "commonjs" }\n`);
    },
};

/**
 * `src/browser` reaches the writer through relative imports into `src/core`,
 * and the core already ships as a bundle of its own. Treating those ids as
 * external is what keeps the browser bundle down to just the helpers, instead
 * of carrying a second copy of the writer that a page would load twice.
 *
 * Rollup asks about the raw specifier first and about the resolved absolute
 * path afterwards; only the second one can be matched, which is enough.
 */
const CORE_DIR = resolve('dist/src/core') + sep;
const isCore = (id) => id.startsWith(CORE_DIR);

/** How the core bundle is named on each side: a module to require, a global. */
const CORE_MODULE = 'xlsx-now';
const CORE_GLOBAL = 'xlsxNow';

export default [
    {
        input: 'dist/src/core/index.js',
        // Dependencies stay out of the bundle: only this package's own modules
        // are bundled. `fflate` is left as a call to whatever `require`/`define`
        // the host provides, exactly like the ESM build does — nothing is
        // inlined, and no copy of a third-party library is duplicated into this
        // file.
        external: ['fflate'],
        output: {
            file: 'dist/umd/xlsx-now.umd.js',
            format: 'umd',
            name: 'xlsxNow',
            sourcemap: true,
            // Name of the external under a plain <script> tag: fflate's own UMD
            // build publishes itself as the `fflate` global, so it has to be
            // loaded before this bundle.
            globals: { fflate: 'fflate' },
        },
        plugins: [commonjsMarker],
    },
    {
        // `downloadXlsx` and `createXlsxBlob`: the save dialog and the Blob,
        // which is the part of the package a page cannot get from the core
        // bundle. Browser-only by nature — it reads `window` and `document`,
        // so unlike the core bundle it is not meant to be required from Node.
        input: 'dist/src/browser/index.js',
        external: isCore,
        // The core is external by absolute path, and Rollup would render such
        // an external as a path relative to the output file. It is not a file
        // to reach from here, it is the package: keep the id as it is so
        // `paths` below can put the package's own name in its place.
        makeAbsoluteExternalsRelative: false,
        output: {
            file: 'dist/umd/xlsx-now-browser.umd.js',
            format: 'umd',
            name: 'xlsxNowBrowser',
            sourcemap: true,
            // The core, named for each way this file can be loaded: required
            // by the package's name under CommonJS and AMD, and read off the
            // global the core bundle publishes under a <script> tag.
            paths: (id) => (isCore(id) ? CORE_MODULE : id),
            globals: (id) => (isCore(id) ? CORE_GLOBAL : id),
        },
        plugins: [commonjsMarker],
    },
];
