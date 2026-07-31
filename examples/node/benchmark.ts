// Measures what the compression change actually costs and buys: output size
// and wall time per compression level, plus peak memory, on a row count large
// enough that buffering the workbook would be visible.
//
// Run via `npm run benchmark` (optionally `ROWS=1000000 npm run benchmark`).
import { createWriteStream } from 'node:fs';
import { mkdir, stat, rm } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Readable } from 'node:stream';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';
import { createXlsxStream } from '../../src/core/createXlsxStream.js';
import type { CompressionLevel } from '../../src/core/zip.js';
import type { Column, Row } from '../../src/core/types.js';

const ROWS = Number(process.env['ROWS'] ?? 200_000);
const LEVELS: CompressionLevel[] = [0, 1, 6, 9];

const columns: Column[] = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
    { name: 'created_at', key: 'createdAt' },
];

function* generateRows(count: number): Generator<Row> {
    for (let i = 1; i <= count; i++) {
        yield {
            id: i,
            name: `Widget ${i}`,
            price: Math.round(i * 3.33 * 100) / 100,
            inStock: i % 5 !== 0,
            createdAt: new Date(Date.UTC(2026, 0, 1 + (i % 28))),
        };
    }
}

interface Result {
    level: CompressionLevel;
    bytes: number;
    seconds: number;
    peakRssMb: number;
}

async function run(level: CompressionLevel): Promise<Result> {
    const outPath = resolve(`out/benchmark-level-${level}.xlsx`);

    let peakRss = 0;
    const sampler = setInterval(() => {
        peakRss = Math.max(peakRss, process.memoryUsage().rss);
    }, 20);

    const started = performance.now();
    const webStream = createXlsxStream({
        columns,
        rows: generateRows(ROWS),
        sheetName: 'Widgets',
        compressionLevel: level,
    });

    await new Promise<void>((resolvePromise, reject) => {
        Readable.fromWeb(webStream as NodeWebReadableStream<Uint8Array>)
            .pipe(createWriteStream(outPath))
            .on('finish', resolvePromise)
            .on('error', reject);
    });
    const seconds = (performance.now() - started) / 1000;
    clearInterval(sampler);

    const { size } = await stat(outPath);
    await rm(outPath);

    return { level, bytes: size, seconds, peakRssMb: peakRss / 1024 / 1024 };
}

const mb = (bytes: number): string => `${(bytes / 1024 / 1024).toFixed(1)} MB`;

await mkdir(resolve('out'), { recursive: true });
console.log(`${ROWS.toLocaleString()} rows x ${columns.length} columns\n`);

const results: Result[] = [];
for (const level of LEVELS) results.push(await run(level));

const stored = results.find((r) => r.level === 0)!;
console.log('| level | file size | vs. stored | time | peak RSS |');
console.log('| --- | --- | --- | --- | --- |');
for (const r of results) {
    const label = r.level === 0 ? '0 (stored)' : String(r.level);
    const vs = r.level === 0 ? '—' : `${((r.bytes / stored.bytes) * 100).toFixed(1)}%`;
    console.log(
        `| ${label} | ${mb(r.bytes)} | ${vs} | ${r.seconds.toFixed(1)} s | ${r.peakRssMb.toFixed(0)} MB |`,
    );
}
