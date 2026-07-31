// Node-side proof: rows arrive in batches (simulating a DB cursor or an
// upstream NDJSON response) and the writer sits in the middle of a plain pipe
// chain, straight into an .xlsx file — the full row set is never held in
// memory at once.
//
// For records that are not already a stream there is `writeXlsxFile` in the
// same module, which does array-or-generator -> file in one call.
import { createWriteStream } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { createXlsxDuplex } from '../../src/node/index.js';
import type { Column, Row } from '../../src/core/types.js';

const columns: Column[] = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
    { name: 'created_at', key: 'createdAt' },
];

const BATCH_SIZE = 50;

// A cursor hands over a page of rows at a time, which is why the writer takes
// `Row[]` and not one record per chunk.
async function* fetchRowBatchesFromUpstream(): AsyncGenerator<Row[]> {
    for (let start = 1; start <= 200; start += BATCH_SIZE) {
        const batch: Row[] = [];
        for (let i = start; i < start + BATCH_SIZE; i++) {
            batch.push({
                id: i,
                name: `Widget ${i}`,
                price: Math.round(i * 3.33 * 100) / 100,
                inStock: i % 5 !== 0,
                createdAt: new Date(Date.UTC(2026, 0, 1 + (i % 28))),
            });
        }
        yield batch;
        // Simulate batches trickling in over the network instead of arriving
        // all at once — this is the case the whole design targets.
        await new Promise((r) => setTimeout(r, 5));
    }
}

// Paths are resolved from the repo root — run via `npm run example:node`.
await mkdir(resolve('out'), { recursive: true });
const outPath = resolve('out/example-node.xlsx');

await pipeline(
    Readable.from(fetchRowBatchesFromUpstream()),
    createXlsxDuplex({ columns, sheetName: 'Widgets' }),
    createWriteStream(outPath),
);

console.log(`Wrote ${outPath}`);
