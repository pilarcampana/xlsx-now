// Builds the UMD bundle. Rollup runs on the JS that `tsc` already emitted
// (`npm run build` runs `tsc` first), so TypeScript compilation stays in one
// place and this config only bundles this package's own modules into one file
// and converts the module format.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';

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

export default {
    input: 'dist/src/core/index.js',
    // Dependencies stay out of the bundle: only this package's own modules are
    // bundled. `fflate` is left as a call to whatever `require`/`define` the
    // host provides, exactly like the ESM build does — nothing is inlined, and
    // no copy of a third-party library is duplicated into this file.
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
};
