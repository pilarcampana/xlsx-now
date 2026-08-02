import assert from 'node:assert/strict';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
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
