// Node-side proof: rows arrive one at a time (simulating a DB cursor or an
// upstream NDJSON response) and are streamed straight into an .xlsx file —
// the full row set is never held in memory at once.
import { createWriteStream } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Readable } from 'node:stream';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';
import { makeZip } from 'client-zip';
import { createXlsxStream } from '../../src/core/createXlsxStream.js';
import type { Column, Row } from '../../src/core/types.js';

const columns: Column[] = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
    { name: 'created_at', key: 'createdAt' },
];

async function* fetchRowsFromUpstream(): AsyncGenerator<Row> {
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

// Paths are resolved from the repo root — run via `npm run example:node`.
await mkdir(resolve('out'), { recursive: true });
const outPath = resolve('out/example-node.xlsx');

await new Promise<void>((resolvePromise, reject) => {
    // Same object at runtime; the cast only bridges the DOM lib's
    // ReadableStream declaration and node:stream/web's.
    Readable.fromWeb(webStream as NodeWebReadableStream<Uint8Array>)
        .pipe(createWriteStream(outPath))
        .on('finish', resolvePromise)
        .on('error', reject);
});

console.log(`Wrote ${outPath}`);
