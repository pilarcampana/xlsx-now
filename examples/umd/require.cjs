// UMD consumer, CommonJS side: a plain `require()` of the bundle — no ESM,
// no import map, no bundler. `fflate` is inlined in the bundle, so there is
// nothing else to load.
//
// Run via `npm run example:umd:node` (which builds first).
const { writeFile, mkdir } = require('node:fs/promises');
const { resolve } = require('node:path');

const { createXlsxStream } = require('../../dist/umd/xlsx-now.umd.js');

const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
];

async function* fetchRowsFromUpstream() {
    for (let i = 1; i <= 200; i++) {
        yield { id: i, name: `Widget ${i}`, price: Math.round(i * 3.33 * 100) / 100, inStock: i % 5 !== 0 };
        if (i % 50 === 0) await new Promise((r) => setTimeout(r, 5));
    }
}

async function main() {
    const stream = createXlsxStream({
        columns,
        rows: fetchRowsFromUpstream(),
        sheetName: 'Widgets',
    });

    // Paths are resolved from the repo root — run via `npm run example:umd:node`.
    await mkdir(resolve('out'), { recursive: true });
    const outPath = resolve('out/example-umd-node.xlsx');
    await writeFile(outPath, Buffer.from(await new Response(stream).arrayBuffer()));
    console.log(`Wrote ${outPath}`);
}

main().catch((err) => {
    console.error(err);
    process.exitCode = 1;
});
