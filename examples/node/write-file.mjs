// Node-side proof: rows arrive one at a time (simulating a DB cursor or an
// upstream NDJSON response) and are streamed straight into an .xlsx file —
// the full row set is never held in memory at once.
import { createWriteStream } from 'node:fs';
import { Readable } from 'node:stream';
import { makeZip } from 'client-zip';
import { createXlsxStream } from '../../src/core/createXlsxStream.js';

const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
    { name: 'created_at', key: 'createdAt' },
];

async function* fetchRowsFromUpstream() {
    for (let i = 1; i <= 200; i++) {
        yield {
            id: i,
            name: `Widget ${i}`,
            price: Math.round(i * 3.33 * 100) / 100,
            inStock: i % 5 !== 0,
            createdAt: new Date(Date.UTC(2026, 0, 1 + (i % 28))),
        };
        // Simulate rows trickling in over the network instead of arriving
        // all at once — this is the case the whole design targets.
        if (i % 50 === 0) await new Promise((r) => setTimeout(r, 5));
    }
}

const webStream = createXlsxStream({
    columns,
    rows: fetchRowsFromUpstream(),
    sheetName: 'Widgets',
    makeZip,
});

const outPath = new URL('../../out/example-node.xlsx', import.meta.url);
await import('node:fs/promises').then((fs) => fs.mkdir(new URL('../../out/', import.meta.url), { recursive: true }));

await new Promise((resolve, reject) => {
    Readable.fromWeb(webStream)
        .pipe(createWriteStream(outPath))
        .on('finish', resolve)
        .on('error', reject);
});

console.log(`Wrote ${outPath.pathname}`);
