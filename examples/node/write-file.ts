// Node-side proof, and the shape the writer is designed for: an NDJSON file
// read as a stream, split into lines, parsed, and written out as an .xlsx —
// the full row set is never held in memory at once.
//
// Every link is a Web Streams `TransformStream`, including the two written
// here: the same chain runs in the browser, and `XlsxStream` is the same
// class the browser example uses. Nothing in it is Node-specific except the
// two ends, the file being read and the file being written.
//
// For records that are not already a stream there is `writeXlsxFile` in
// `xlsx-now/node`, which does array-or-generator -> file in one call.
import { openAsBlob } from 'node:fs';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { XlsxStream } from '../../src/core/xlsxStream.js';
import { createFileWritable } from '../../src/node/index.js';
import type { Column, Row } from '../../src/core/types.js';

const columns: Column[] = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
    { name: 'created_at', key: 'createdAt' },
];

/** Splits a stream of text into lines, keeping the tail between chunks. */
class LineSplitStream extends TransformStream<string, string> {
    constructor() {
        let rest = '';
        super({
            transform(chunk, controller) {
                const lines = (rest + chunk).split('\n');
                rest = lines.pop() ?? '';
                for (const line of lines) if (line) controller.enqueue(line);
            },
            flush(controller) {
                if (rest) controller.enqueue(rest);
            },
        });
    }
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}T/;

class JsonParseStream extends TransformStream<string, Row> {
    constructor() {
        super({
            transform(line, controller) {
                // JSON has no date type, so they arrive as text and a real
                // pipeline revives them here — which is what gets `created_at`
                // written as a date and not as a string.
                const revive = (_key: string, value: unknown): unknown =>
                    typeof value === 'string' && ISO_DATE.test(value) ? new Date(value) : value;
                controller.enqueue(JSON.parse(line, revive) as Row);
            },
        });
    }
}

// Paths are resolved from the repo root — run via `npm run example:node`.
await mkdir(resolve('out'), { recursive: true });
const inPath = resolve('out/example-node.ndjson');
const outPath = resolve('out/example-node.xlsx');

// Stands in for whatever really produces the records: an export from another
// service, a dump, a cursor written out line by line.
await writeFile(
    inPath,
    Array.from({ length: 200 }, (_, k) =>
        JSON.stringify({
            id: k + 1,
            name: `Widget ${k + 1}`,
            price: Math.round((k + 1) * 3.33 * 100) / 100,
            inStock: (k + 1) % 5 !== 0,
            createdAt: new Date(Date.UTC(2026, 0, 1 + ((k + 1) % 28))),
        }),
    ).join('\n'),
);

await (await openAsBlob(inPath))
    .stream()
    .pipeThrough(new TextDecoderStream())
    .pipeThrough(new LineSplitStream())
    .pipeThrough(new JsonParseStream())
    .pipeThrough(new XlsxStream({ columns, sheetName: 'Widgets' }))
    .pipeTo(createFileWritable(outPath));

console.log(`Wrote ${outPath}`);
