import assert from 'node:assert/strict';
import {
    createXlsxStream,
    type CreateXlsxStreamOptions,
    type XlsxSheet,
} from '../src/core/createXlsxStream.js';
import type { CellRow, Column } from '../src/core/types.js';
import { asAsyncIterable, asIterable, collect } from './helpers/streams.js';
import { readXlsx, sheetRows } from './helpers/zip.js';

const COLUMNS: readonly Column[] = [{ name: 'id', pk: true }, { name: 'name' }];
const RECORDS = [
    { id: 1, name: 'Ana' },
    { id: 2, name: 'Beto' },
];

describe('createXlsxStream', () => {
    it('takes an array of records', async () => {
        const { sheet } = await readXlsx(
            await collect(createXlsxStream({ columns: COLUMNS, rows: RECORDS })),
        );
        assert.equal(sheetRows(sheet).length, 3); // header + 2
        assert.ok(sheet.includes('>Beto<'));
    });

    it('takes a sync generator', async () => {
        const { sheet } = await readXlsx(
            await collect(createXlsxStream({ columns: COLUMNS, rows: asIterable(RECORDS) })),
        );
        assert.equal(sheetRows(sheet).length, 3);
    });

    it('takes an async source, one row at a time', async () => {
        const { sheet } = await readXlsx(
            await collect(createXlsxStream({ columns: COLUMNS, rows: asAsyncIterable(RECORDS) })),
        );
        assert.equal(sheetRows(sheet).length, 3);
        assert.ok(sheet.includes('>Ana<'));
    });

    it('takes rows arrays with no columns declared', async () => {
        const rows: CellRow[] = [['a', 1], ['b', 2]];
        const { sheet } = await readXlsx(await collect(createXlsxStream({ rows })));
        assert.equal(sheetRows(sheet).length, 2);
    });

    it('produces a readable file out of no rows at all', async () => {
        const { names, sheet } = await readXlsx(await collect(createXlsxStream({ rows: [] })));
        assert.equal(names.length, 6);
        assert.ok(sheet.includes('<sheetData></sheetData>'));
    });

    it('reads records only as the consumer asks for bytes', async () => {
        let read = 0;
        async function* counted(): AsyncGenerator<CellRow> {
            for (let i = 0; i < 10000; i++) {
                read++;
                yield [i, 'a fairly long value, so a batch fills in a few thousand rows'];
            }
        }

        const reader = createXlsxStream({ rows: counted() }).getReader();
        // The first chunks are the parts written before any record — the
        // workbook, the styles — so read on until the worksheet needs one.
        for (let chunk = 0; chunk < 100 && read === 0; chunk++) await reader.read();
        const consumed = read;
        await reader.cancel();

        assert.ok(consumed > 0, 'no record was ever read');
        assert.ok(consumed < 10000, `the whole source was drained (${consumed})`);
    });

    it('one pull always makes progress, however small the rows are', async () => {
        // Rows this short take many of them to fill a batch; the pull loop is
        // what keeps a read from returning nothing.
        const stream = createXlsxStream({ rows: asIterable(Array.from({ length: 20000 }, () => [1])) });
        const reader = stream.getReader();
        const first = await reader.read();
        assert.equal(first.done, false);
        assert.ok((first.value?.length ?? 0) > 0, 'a read returned an empty chunk');
        await reader.cancel();
    });

    it('tells the source it is done when the consumer cancels', async () => {
        let returned: unknown = 'not called';
        const rows: Iterable<CellRow> = {
            [Symbol.iterator]: () => ({
                next: () => ({ done: false, value: ['x'] }),
                return: (value?: unknown) => {
                    returned = value;
                    return { done: true as const, value: undefined };
                },
            }),
        };

        const stream = createXlsxStream({ rows });
        const reader = stream.getReader();
        await reader.read();
        await reader.cancel('enough');
        assert.equal(returned, 'enough');
    });

    it('survives a source with nothing to return to', async () => {
        const rows: Iterable<CellRow> = {
            [Symbol.iterator]: () => ({ next: () => ({ done: false, value: ['x'] }) }),
        };
        const reader = createXlsxStream({ rows }).getReader();
        await reader.read();
        await reader.cancel();
    });

    it('lets a failing source through to the consumer', async () => {
        async function* failing(): AsyncGenerator<CellRow> {
            yield ['a'];
            throw new Error('the cursor died');
        }
        await assert.rejects(collect(createXlsxStream({ rows: failing() })), /the cursor died/);
    });
});

describe('createXlsxStream with sheets', () => {
    const SHEETS: XlsxSheet[] = [
        { name: 'Ventas', columns: COLUMNS, rows: RECORDS },
        { name: 'Costos', columns: COLUMNS, rows: [{ id: 3, name: 'Caro' }] },
    ];

    it('writes one worksheet per sheet, in order', async () => {
        const { sheetNames, sheets } = await readXlsx(
            await collect(createXlsxStream({ sheets: SHEETS })),
        );
        assert.deepEqual(sheetNames, ['Ventas', 'Costos']);
        assert.equal(sheetRows(sheets[0] ?? '').length, 3); // header + 2
        assert.equal(sheetRows(sheets[1] ?? '').length, 2); // header + 1
        assert.ok((sheets[1] ?? '').includes('>Caro<'));
    });

    it('takes an array of sheets, and an async source of them alike', async () => {
        const fromArray = await readXlsx(await collect(createXlsxStream({ sheets: SHEETS })));
        const fromAsync = await readXlsx(
            await collect(createXlsxStream({ sheets: asAsyncIterable(SHEETS) })),
        );
        assert.deepEqual(fromAsync.sheetNames, fromArray.sheetNames);
        assert.deepEqual(fromAsync.sheets, fromArray.sheets);
    });

    it('takes rows a sheet pulls for itself, not only an array', async () => {
        const sheets = [{ name: 'Cursor', columns: COLUMNS, rows: asAsyncIterable(RECORDS) }];
        const { sheetNames, sheets: parts } = await readXlsx(
            await collect(createXlsxStream({ sheets })),
        );
        assert.deepEqual(sheetNames, ['Cursor']);
        assert.equal(sheetRows(parts[0] ?? '').length, 3);
    });

    it('carries every sheet option, not just the columns', async () => {
        const sheets: XlsxSheet[] = [
            {
                name: 'Congelada',
                columns: COLUMNS,
                freezeRows: 2,
                freezeColumns: 1,
                columnFormats: { B: { width: 30 } },
                rows: RECORDS,
            },
        ];
        const { sheets: parts } = await readXlsx(await collect(createXlsxStream({ sheets })));
        const sheet = parts[0] ?? '';
        assert.ok(sheet.includes('ySplit="2"'), 'freezeRows did not reach the sheet');
        assert.ok(sheet.includes('xSplit="1"'), 'freezeColumns did not reach the sheet');
        assert.ok(sheet.includes('width="30"'), 'columnFormats did not reach the sheet');
    });

    it('falls back to the workbook options for whatever a sheet leaves out', async () => {
        const { sheets } = await readXlsx(
            await collect(
                createXlsxStream({
                    columns: COLUMNS,
                    sheets: [{ name: 'Hereda', rows: RECORDS }],
                }),
            ),
        );
        // The records were read by the workbook's columns, header row and all.
        assert.equal(sheetRows(sheets[0] ?? '').length, 3);
        assert.ok((sheets[0] ?? '').includes('>Ana<'));
    });

    it('reads a sheet only while it is the one being written', async () => {
        const started: string[] = [];
        function rowsOf(name: string): AsyncGenerator<CellRow> {
            return (async function* () {
                started.push(name);
                yield [name];
            })();
        }
        const sheets = [
            { name: 'Primera', rows: rowsOf('Primera') },
            { name: 'Segunda', rows: rowsOf('Segunda') },
        ];

        // An async generator does not run until it is iterated, so nothing has
        // started before the stream is drained, and then only in order.
        assert.deepEqual(started, []);
        await collect(createXlsxStream({ sheets }));
        assert.deepEqual(started, ['Primera', 'Segunda']);
    });

    it('tells the sheet and its rows they are done when the consumer cancels', async () => {
        let rowsStarted = false;
        let sheetsReturned = false;
        let rowsReturned = false;
        const rows: Iterable<CellRow> = {
            [Symbol.iterator]: () => {
                rowsStarted = true;
                return {
                    next: () => ({ done: false, value: ['x'] }),
                    return: () => {
                        rowsReturned = true;
                        return { done: true as const, value: undefined };
                    },
                };
            },
        };
        const sheets: Iterable<XlsxSheet> = {
            [Symbol.iterator]: () => ({
                next: () => ({ done: false, value: { name: 'Infinita', rows } }),
                return: () => {
                    sheetsReturned = true;
                    return { done: true as const, value: undefined };
                },
            }),
        };

        // The first chunks are the parts written before any message, so read
        // on until the sheet being cancelled is really the one open.
        const reader = createXlsxStream({ sheets }).getReader();
        for (let chunk = 0; chunk < 100 && !rowsStarted; chunk++) await reader.read();
        assert.ok(rowsStarted, 'the sheet was never opened');

        await reader.cancel();
        assert.ok(rowsReturned, 'the open sheet was never closed');
        assert.ok(sheetsReturned, 'the source of sheets was never closed');
    });

    it('lets a failing sheet source through to the consumer', async () => {
        async function* failing(): AsyncGenerator<XlsxSheet> {
            yield { name: 'Buena', rows: [['a']] };
            throw new Error('the report list died');
        }
        await assert.rejects(
            collect(createXlsxStream({ sheets: failing() })),
            /the report list died/,
        );
    });

    it('refuses rows and sheets together, and neither', async () => {
        assert.throws(
            () =>
                createXlsxStream({
                    rows: RECORDS,
                    sheets: SHEETS,
                } as unknown as CreateXlsxStreamOptions),
            /two ways to say the same thing/,
        );
        assert.throws(
            () => createXlsxStream({} as unknown as CreateXlsxStreamOptions),
            /Nothing to write/,
        );
    });
});
