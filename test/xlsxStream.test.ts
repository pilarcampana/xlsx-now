import assert from 'node:assert/strict';
import { XlsxStream } from '../src/core/xlsxStream.js';
import type { CellRow, Column, Row } from '../src/core/types.js';
import { asReadable, collect } from './helpers/streams.js';
import { readXlsx, sheetRows } from './helpers/zip.js';

const COLUMNS: readonly Column[] = [{ name: 'id', pk: true }, { name: 'name' }];

describe('XlsxStream', () => {
    it('turns records going in into file bytes coming out', async () => {
        const records: Row[] = [
            { id: 1, name: 'Ana' },
            { id: 2, name: 'Beto' },
        ];
        const bytes = await collect(
            asReadable(records).pipeThrough(new XlsxStream({ columns: COLUMNS })),
        );

        const { names, sheet } = await readXlsx(bytes);
        assert.equal(names.length, 6);
        assert.equal(sheetRows(sheet).length, 3); // header + 2
        assert.ok(sheet.includes('>Beto<'));
    });

    it('takes rows arrays with no columns declared', async () => {
        const rows: CellRow[] = [['a', 1]];
        const bytes = await collect(asReadable(rows).pipeThrough(new XlsxStream({})));
        assert.equal(sheetRows((await readXlsx(bytes)).sheet).length, 1);
    });

    it('closes the file when the writable side closes with nothing in it', async () => {
        const bytes = await collect(
            asReadable<Row>([]).pipeThrough(new XlsxStream({ columns: COLUMNS })),
        );
        const { sheet } = await readXlsx(bytes);
        // The header row is still there, and so is the end of the archive.
        assert.equal(sheetRows(sheet).length, 1);
    });

    it('is a TransformStream, with both sides of one', () => {
        const stream = new XlsxStream({ columns: COLUMNS });
        assert.ok(stream instanceof TransformStream);
        assert.ok(stream.readable instanceof ReadableStream);
        assert.ok(stream.writable instanceof WritableStream);
    });
});
