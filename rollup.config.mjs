// Builds the UMD bundle. Rollup runs on the JS that `tsc` already emitted
// (`npm run build` runs `tsc` first), so TypeScript compilation stays in one
// place and this config only does module-format conversion + dependency
// inlining.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import nodeResolve from '@rollup/plugin-node-resolve';

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
    output: {
        file: 'dist/umd/xlsx-now.umd.js',
        format: 'umd',
        name: 'xlsxNow',
        sourcemap: true,
        // `fflate` is inlined so the bundle works from a script tag with
        // nothing else to load.
        inlineDynamicImports: true,
    },
    plugins: [nodeResolve({ browser: true }), commonjsMarker],
};
