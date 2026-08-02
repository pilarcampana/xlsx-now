import assert from 'node:assert/strict';
import { XlsxWriter, type RowOf, type XlsxWriterOptions } from '../src/core/xlsxWriter.js';
import type { CellRow, Column } from '../src/core/types.js';
import { recordingSink } from './helpers/streams.js';
import { METHOD_DEFLATE, METHOD_STORE, SHEET_PART, readXlsx, sheetRows } from './helpers/zip.js';

const COLUMNS: readonly Column[] = [
    { name: 'id', pk: true },
    { name: 'Full name', key: 'full_name' },
];

/** A whole file, out of the writer alone. */
function write<O extends XlsxWriterOptions>(options: O, rows: readonly RowOf<O>[]): Buffer {
    const { sink, bytes } = recordingSink();
    const writer = new XlsxWriter<O>(sink, options);
    for (const row of rows) writer.writeRow(row);
    writer.finish();
    return bytes();
}

describe('XlsxWriter: the package it writes', () => {
    it('carries the five parts an .xlsx needs, worksheet last', async () => {
        const { names } = await readXlsx(write({}, []));
        assert.deepEqual(names, [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/workbook.xml',
            'xl/styles.xml',
            'xl/_rels/workbook.xml.rels',
            SHEET_PART,
        ]);
    });

    it('declares in [Content_Types].xml every part it wrote', async () => {
        const { names, byName } = await readXlsx(write({}, []));
        const contentTypes = byName.get('[Content_Types].xml')?.text ?? '';
        for (const name of names) {
            if (name.endsWith('.rels') || name === '[Content_Types].xml') continue;
            assert.ok(contentTypes.includes(`PartName="/${name}"`), `${name} is not declared`);
        }
    });

    it('names the sheet Sheet1 unless told otherwise', async () => {
        const named = await readXlsx(write({ sheetName: 'Ventas & más' }, []));
        assert.ok(named.byName.get('xl/workbook.xml')?.text.includes('name="Ventas &amp; más"'));

        const unnamed = await readXlsx(write({}, []));
        assert.ok(unnamed.byName.get('xl/workbook.xml')?.text.includes('name="Sheet1"'));
    });

    it('deflates by default and stores at level 0', async () => {
        const deflated = await readXlsx(write({}, []));
        assert.equal(deflated.byName.get(SHEET_PART)?.method, METHOD_DEFLATE);

        const stored = await readXlsx(write({ compressionLevel: 0 }, []));
        assert.equal(stored.byName.get(SHEET_PART)?.method, METHOD_STORE);
    });
});

describe('XlsxWriter: the rows mode', () => {
    it('writes one row per array, numbered from 1', async () => {
        const rows: CellRow[] = [['a', 1], [true, new Date(0)]];
        const { sheet } = await readXlsx(write({}, rows));
        assert.deepEqual(sheetRows(sheet), [
            '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">a</t></is></c>' +
                '<c r="B1" t="n"><v>1</v></c></row>',
            '<row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2" t="n"><v>25569</v></c></row>',
        ]);
    });

    it('declares no header row of its own, so nothing is frozen', async () => {
        const { sheet } = await readXlsx(write({}, [['a']]));
        assert.equal(sheetRows(sheet).length, 1);
        assert.ok(sheet.includes('<sheetView workbookViewId="0"/>'), 'a pane was written');
    });

    it('writes a well-formed empty sheet with no rows at all', async () => {
        const { sheet } = await readXlsx(write({}, []));
        assert.ok(sheet.includes('<sheetData></sheetData>'));
        assert.ok(sheet.endsWith('</worksheet>'));
    });
});

describe('XlsxWriter: the columns mode', () => {
    it('writes the header row first, then one row per record', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS }, [{ id: 1, full_name: 'Ana' }]),
        );
        const rows = sheetRows(sheet);
        assert.equal(rows.length, 2);
        assert.ok(rows[0]?.includes('>id<'), 'the header is not the column names');
        assert.ok(rows[0]?.includes('>Full name<'));
        assert.ok(rows[1]?.includes('>Ana<'), 'the record was not read by key');
    });

    it('styles the header bold, and the pk column highlighted throughout', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS }, [{ id: 1, full_name: 'Ana' }]),
        );
        const [header, data] = sheetRows(sheet);
        // 3 is bold + highlight, 1 is bold, 2 is highlight.
        assert.ok(header?.includes('<c r="A1" t="inlineStr" s="3">'), header ?? 'no header row');
        assert.ok(header?.includes('<c r="B1" t="inlineStr" s="1">'), header ?? 'no header row');
        assert.ok(data?.includes('<c r="A2" t="n" s="2">'), data ?? 'no data row');
        assert.ok(data?.includes('<c r="B2" t="inlineStr">'), 'a plain cell got a style');
    });

    it('freezes the header row and the leading pk columns', async () => {
        const { sheet } = await readXlsx(write({ columns: COLUMNS }, []));
        assert.ok(sheet.includes('<pane xSplit="1" ySplit="1" topLeftCell="B2"'), sheet);
    });

    it('writes the header row even when no record follows', async () => {
        const { sheet } = await readXlsx(write({ columns: COLUMNS }, []));
        assert.equal(sheetRows(sheet).length, 1);
    });

    it('writes the header row even when it alone fills a batch', async () => {
        // Enough columns that the header reaches the zip before the first
        // record does, which is the one path where writing it can fail.
        const columns = Array.from({ length: 5000 }, (_, i) => ({ name: `column ${i}` }));
        const { sheet } = await readXlsx(write({ columns }, [{}]));
        assert.equal(sheetRows(sheet).length, 2);
        assert.ok(sheet.includes('>column 4999<'));
    });
});

describe('XlsxWriter: explicit freezes', () => {
    it('freezes what it is asked to in the rows mode', async () => {
        const { sheet } = await readXlsx(write({ freezeRows: 2, freezeColumns: 3 }, []));
        assert.ok(sheet.includes('<pane xSplit="3" ySplit="2" topLeftCell="D3"'), sheet);
    });

    it('overrides what the columns imply', async () => {
        const { sheet } = await readXlsx(write({ columns: COLUMNS, freezeColumns: 0 }, []));
        assert.ok(sheet.includes('<pane ySplit="1" topLeftCell="A2"'), sheet);
        assert.doesNotMatch(sheet, /xSplit/);
    });

    it('takes a zero as a freeze of nothing, not as no answer', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS, freezeRows: 0, freezeColumns: 0 }, []),
        );
        assert.ok(sheet.includes('<sheetView workbookViewId="0"/>'), sheet);
    });
});

describe('XlsxWriter: streaming and failure', () => {
    it('hands bytes out while the rows are still coming', () => {
        const { sink, chunks } = recordingSink();
        const writer = new XlsxWriter(sink, { compressionLevel: 0 });
        for (let i = 0; i < 5000; i++) writer.writeRow([i, 'a fairly long value to fill the batch']);
        const during = chunks.length;
        writer.finish();
        assert.ok(during > 0, 'nothing came out before finish()');
        assert.ok(chunks.length > during, 'finish() produced nothing');
    });

    it('keeps numbering rows across batches', async () => {
        const rows = Array.from({ length: 2000 }, (_, i) => [i, 'x'.repeat(60)]);
        const { sheet } = await readXlsx(write({}, rows));
        const written = sheetRows(sheet);
        assert.equal(written.length, 2000);
        assert.ok(written[1999]?.startsWith('<row r="2000">'), written[1999] ?? 'no row 2000');
    });

    it('lets a failing sink through, from the constructor on', () => {
        assert.throws(
            () =>
                new XlsxWriter(() => {
                    throw new Error('sink is full');
                }, { compressionLevel: 0 }),
            /sink is full/,
        );
    });

    it('lets a failing sink through from writeRow and from finish', () => {
        let failing = false;
        const sink = (): void => {
            if (failing) throw new Error('sink is full');
        };

        const fromWriteRow = new XlsxWriter(sink, { compressionLevel: 0 });
        failing = true;
        assert.throws(() => {
            for (let i = 0; i < 5000; i++) fromWriteRow.writeRow([i, 'a fairly long value to fill the batch']);
        }, /sink is full/);

        failing = false;
        const fromFinish = new XlsxWriter(sink, { compressionLevel: 0 });
        failing = true;
        assert.throws(() => fromFinish.finish(), /sink is full/);
    });

    it('abandons the archive after a failure instead of leaving half a file behind', async () => {
        let failing = false;
        const { sink, bytes } = recordingSink();
        const writer = new XlsxWriter((chunk) => {
            if (failing) throw new Error('sink is full');
            sink(chunk);
        }, { compressionLevel: 0 });

        failing = true;
        assert.throws(() => {
            for (let i = 0; i < 5000; i++) writer.writeRow([i, 'a fairly long value to fill the batch']);
        }, /sink is full/);

        // Any failure mid-write leaves the archive unreadable, and what came
        // out before it is not a file anyone can open.
        await assert.rejects(readXlsx(bytes()));
    });
});
