// Node-side proof: the writer sits in the middle of a plain pipe chain, one
// record per chunk in and an .xlsx file out — the full row set is never held
// in memory at once. Upstream would normally be a file being split into lines
// and parsed, or a database cursor; here it is simulated.
//
// For records that are not already a stream there is `writeXlsxFile` in the
// same module, which does array-or-generator -> file in one call.
import { createWriteStream } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { XlsxTransform } from '../../src/node/index.js';
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

// Paths are resolved from the repo root — run via `npm run example:node`.
await mkdir(resolve('out'), { recursive: true });
const outPath = resolve('out/example-node.xlsx');

await pipeline(
    Readable.from(fetchRowsFromUpstream()),
    new XlsxTransform({ columns, sheetName: 'Widgets' }),
    createWriteStream(outPath),
);

console.log(`Wrote ${outPath}`);
