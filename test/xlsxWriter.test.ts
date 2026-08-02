import assert from 'node:assert/strict';
import { XlsxWriter, type XlsxWriterOptions } from '../src/core/xlsxWriter.js';
import type { SheetInput } from '../src/core/command.js';
import type { CellRow, Column } from '../src/core/types.js';
import { recordingSink } from './helpers/streams.js';
import { METHOD_DEFLATE, METHOD_STORE, SHEET_PART, readXlsx, sheetRows } from './helpers/zip.js';

const COLUMNS: readonly Column[] = [
    { name: 'id', pk: true },
    { name: 'Full name', key: 'full_name' },
];

/** A whole file, out of the writer alone. */
function write(options: XlsxWriterOptions, rows: readonly SheetInput[]): Buffer {
    const { sink, bytes } = recordingSink();
    const writer = new XlsxWriter(sink, options);
    for (const row of rows) writer.writeRow(row);
    writer.finish();
    return bytes();
}

describe('XlsxWriter: the package it writes', () => {
    it('carries the five parts an .xlsx needs, the sheet before what describes it', async () => {
        // The styles and the workbook come after the sheet because neither is
        // known until the last row is in: what a cell asked for, and how many
        // sheets there turned out to be. Order inside the archive is nobody's
        // business but the central directory's.
        const { names } = await readXlsx(write({}, []));
        assert.deepEqual(names, [
            '[Content_Types].xml',
            '_rels/.rels',
            SHEET_PART,
            'xl/styles.xml',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
        ]);
    });

    it('keeps [Content_Types].xml first, where a package reader wants it', async () => {
        const { names } = await readXlsx(write({}, []));
        assert.equal(names[0], '[Content_Types].xml');
    });

    it('declares in [Content_Types].xml every part it wrote', async () => {
        const { names, byName } = await readXlsx(write({}, []));
        const contentTypes = byName.get('[Content_Types].xml')?.text ?? '';
        for (const name of names) {
            if (name.endsWith('.rels') || name === '[Content_Types].xml') continue;
            // The worksheets are typed by extension: how many of them there
            // are is not known when this part goes out.
            const declared = name.startsWith('xl/worksheets/')
                ? contentTypes.includes('<Default Extension="xml"')
                : contentTypes.includes(`PartName="/${name}"`);
            assert.ok(declared, `${name} is not declared`);
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
            '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>' +
                '<c r="B1"><v>1</v></c></row>',
            '<row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2" s="1"><v>25569</v></c></row>',
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
        // The indexes are handed out as the cells ask: the pk header first,
        // then the plain header, then the pk fill on its own.
        assert.ok(header?.includes('<c r="A1" t="inlineStr" s="1">'), header ?? 'no header row');
        assert.ok(header?.includes('<c r="B1" t="inlineStr" s="2">'), header ?? 'no header row');
        assert.ok(data?.includes('<c r="A2" s="3">'), data ?? 'no data row');
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

describe('XlsxWriter: the #worksheet command', () => {
    it('closes the sheet being written and opens the next one', async () => {
        const { sheetNames, sheets } = await readXlsx(
            write({ sheetName: 'Enero' }, [['a'], { '#worksheet': 'Febrero' }, ['b'], ['c']]),
        );
        const [enero = '', febrero = ''] = sheets;
        assert.deepEqual(sheetNames, ['Enero', 'Febrero']);
        assert.equal(sheetRows(enero).length, 1);
        assert.equal(sheetRows(febrero).length, 2);
        assert.ok(febrero.includes('>c<'));
        assert.ok(!enero.includes('>c<'), 'a row went to the wrong sheet');
    });

    it('writes one part per sheet, and relates every one of them', async () => {
        const { names, byName } = await readXlsx(
            write({}, [['a'], { '#worksheet': 'B' }, ['b'], { '#worksheet': 'C' }, ['c']]),
        );
        assert.ok(names.includes('xl/worksheets/sheet3.xml'), names.join(', '));
        const rels = byName.get('xl/_rels/workbook.xml.rels')?.text ?? '';
        assert.match(rels, /Id="rId3"[^>]*Target="worksheets\/sheet3\.xml"/);
        assert.match(rels, /Id="rId4"[^>]*Target="styles\.xml"/);
    });

    it('numbers the rows of every sheet from one', async () => {
        const { sheets } = await readXlsx(
            write({}, [['a'], ['b'], { '#worksheet': 'Second' }, ['c']]),
        );
        const second = sheets[1] ?? '';
        assert.ok(second.startsWith('<?xml'));
        assert.ok(sheetRows(second)[0]?.startsWith('<row r="1">'), second);
    });

    it('configures the first sheet when it arrives before any row', async () => {
        // Nothing is written until the first message, so a stream that
        // declares itself on the way in leaves no empty sheet behind.
        const { sheetNames, sheets } = await readXlsx(
            write({}, [{ '#worksheet': 'Ventas', columns: COLUMNS }, { id: 1, full_name: 'Ana' }]),
        );
        const ventas = sheets[0] ?? '';
        assert.deepEqual(sheetNames, ['Ventas']);
        assert.equal(sheetRows(ventas).length, 2);
        assert.ok(ventas.includes('>Full name<'));
    });

    it('gives every sheet the columns its own command declares', async () => {
        const { sheets } = await readXlsx(
            write({ columns: COLUMNS }, [
                { id: 1, full_name: 'Ana' },
                { '#worksheet': 'Otra', columns: [{ name: 'total' }] },
                { total: 7 },
            ]),
        );
        const [first = '', second = ''] = sheets;
        assert.ok(first.includes('>Full name<'));
        assert.ok(second.includes('>total<'), second);
        assert.ok(!second.includes('>Full name<'), 'the old columns stayed on');
    });

    it('carries the workbook columns over to a sheet that declares none', async () => {
        const { sheets } = await readXlsx(
            write({ columns: COLUMNS }, [
                { '#worksheet': 'Segunda' },
                { id: 2, full_name: 'Beto' },
            ]),
        );
        const segunda = sheets[0] ?? '';
        assert.ok(segunda.includes('>Full name<'), 'the header row is missing');
        assert.ok(segunda.includes('>Beto<'), 'the record was not read by key');
    });

    it('takes an empty list of columns as a sheet of plain rows', async () => {
        const { sheets } = await readXlsx(
            write({ columns: COLUMNS }, [{ '#worksheet': 'Libre', columns: [] }, ['a', 'b']]),
        );
        const libre = sheets[0] ?? '';
        assert.equal(sheetRows(libre).length, 1, 'a header row was written');
        assert.ok(libre.includes('<sheetView workbookViewId="0"/>'), 'a pane was written');
    });

    it('freezes per sheet what its own command asks for', async () => {
        const { sheets } = await readXlsx(
            write({}, [{ '#worksheet': 'Fija', freezeRows: 1, freezeColumns: 2 }]),
        );
        const fija = sheets[0] ?? '';
        assert.ok(fija.includes('<pane xSplit="2" ySplit="1" topLeftCell="C2"'), fija);
    });

    it('leaves a well-formed empty sheet when nothing follows the command', async () => {
        const { sheetNames, sheets } = await readXlsx(write({}, [['a'], { '#worksheet': 'Vacía' }]));
        const vacia = sheets[1] ?? '';
        assert.deepEqual(sheetNames, ['Sheet1', 'Vacía']);
        assert.ok(vacia.includes('<sheetData></sheetData>'), vacia);
        assert.ok(vacia.endsWith('</worksheet>'));
    });

    it('leaves the first sheet blank when one command follows another', async () => {
        const { sheetNames, sheets } = await readXlsx(
            write({}, [{ '#worksheet': 'hoja1' }, { '#worksheet': 'hoja2' }, ['a']]),
        );
        // The first command is the first sheet rather than a second one, so
        // the second command closes a sheet nothing was written in — which is
        // a sheet in the workbook all the same.
        assert.deepEqual(sheetNames, ['hoja1', 'hoja2']);
        assert.ok(sheets[0]?.includes('<sheetData></sheetData>'), sheets[0] ?? 'no first sheet');
        assert.equal(sheetRows(sheets[1] ?? '').length, 1);
    });

    it('names the first sheet after the command, not after the options', async () => {
        const { sheetNames } = await readXlsx(
            write({ sheetName: 'Ignorada' }, [{ '#worksheet': 'La primera' }, ['a']]),
        );
        assert.deepEqual(sheetNames, ['La primera']);
    });

    it('makes a name Excel would refuse fit, rather than refusing it', async () => {
        // A file Excel will not open is worse than a sheet whose name lost a
        // slash, and by the time anyone finds out the rows are gone.
        const { sheetNames } = await readXlsx(
            write({ sheetName: 'Enero/Febrero' }, [
                ['a'],
                { '#worksheet': '' },
                { '#worksheet': 'x'.repeat(40) },
                { '#worksheet': 'enero/febrero' },
            ]),
        );
        assert.deepEqual(sheetNames, [
            'EneroFebrero',
            'Sheet2',
            'x'.repeat(31),
            'enerofebrero (2)',
        ]);
    });
});

describe('XlsxWriter: the #line command', () => {
    it('writes a record read by the columns, said outright', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS }, [{ '#line': 'row', values: { id: 1, full_name: 'Ana' } }]),
        );
        assert.ok(sheetRows(sheet)[1]?.includes('>Ana<'), sheet);
    });

    it('writes an array said outright, the same as a bare one', async () => {
        const spelled = await readXlsx(write({}, [{ '#line': 'array', values: ['a', 1] }]));
        const bare = await readXlsx(write({}, [['a', 1]]));
        assert.deepEqual(sheetRows(spelled.sheet), sheetRows(bare.sheet));
    });

    it('writes an empty line, the same as an empty array does', async () => {
        const { sheet } = await readXlsx(write({}, [['a'], { '#line': 'empty' }, ['b']]));
        const rows = sheetRows(sheet);
        assert.equal(rows.length, 3);
        assert.equal(rows[1], '<row r="2"></row>');
    });

    it('writes the columns the cells name, and nothing in between', async () => {
        const { sheet } = await readXlsx(
            write({}, [{ '#line': 'array', values: [{ v: 'first', col: 'A' }, { v: 4, col: 'D' }] }]),
        );
        const [row = ''] = sheetRows(sheet);
        assert.ok(row.includes('r="A1"'), row);
        assert.ok(row.includes('r="D1"'), row);
        assert.ok(!row.includes('r="B1"'), 'a gap was written as a cell');
    });

    it('gives the row the height, the style and the hiding it asks for', async () => {
        const { sheet } = await readXlsx(
            write({}, [
                { '#line': 'array', values: ['Total'], height: 22, s: { bold: true } },
                { '#line': 'empty', hidden: true },
            ]),
        );
        const [total = '', hidden = ''] = sheetRows(sheet);
        assert.ok(total.startsWith('<row r="1" s="1" customFormat="1" ht="22" customHeight="1">'), total);
        assert.ok(hidden.startsWith('<row r="2" hidden="1">'), hidden);
    });

    it('leaves a bare row unadorned, as it has nowhere to say otherwise', async () => {
        const { sheet } = await readXlsx(write({}, [['a']]));
        assert.ok(sheetRows(sheet)[0]?.startsWith('<row r="1">'), sheet);
    });

    it('says what a line nobody knows was asked to be', () => {
        assert.throws(() => write({}, [{ '#line': 'colum' } as never]), /Unknown line/);
    });

    it('needs the sheet\'s columns for a record, like a bare one does', () => {
        assert.throws(() => write({}, [{ '#line': 'row', values: { id: 1 } }]), /needs columns/);
    });

    it('opens the first sheet, as any other line does', async () => {
        const { sheetNames, sheet } = await readXlsx(write({}, [{ '#line': 'empty' }]));
        assert.deepEqual(sheetNames, ['Sheet1']);
        assert.equal(sheetRows(sheet).length, 1);
    });
});

describe('XlsxWriter: rows of cells and records, together', () => {
    it('takes both on the same sheet, in any order', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS }, [
                { id: 1, full_name: 'Ana' },
                ['—', { v: 'a note across the sheet', s: { bold: true } }],
                { id: 2, full_name: 'Beto' },
            ]),
        );
        const rows = sheetRows(sheet);
        assert.equal(rows.length, 4);
        assert.ok(rows[2]?.includes('>a note across the sheet<'), rows[2] ?? 'no third row');
        assert.ok(rows[3]?.includes('>Beto<'), rows[3] ?? 'no fourth row');
    });

    it('says what a record on a sheet with no columns is missing', () => {
        assert.throws(() => write({}, [{ id: 1 }]), /needs columns/);
    });

    it('names a command nobody knows rather than writing a blank row', () => {
        assert.throws(() => write({}, [{ '#worksheets': 'Sheet2' }]), /Unknown command/);
        // Columns or not: a `#` key is a command that was meant, not a
        // property nobody declared.
        assert.throws(
            () => write({ columns: COLUMNS }, [{ '#lines': 'empty' }]),
            /Unknown command/,
        );
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

describe('XlsxWriter: the styles the workbook carries', () => {
    it('writes the part at the end, with everything the rows asked for in it', async () => {
        const { byName } = await readXlsx(
            write({ styles: { title: { bold: true, size: 20 } } }, [
                [{ v: 'Report', s: 'title' }],
                [{ v: 1234.5, s: { numFmt: '#,##0.00' } }],
            ]),
        );
        const styles = byName.get('xl/styles.xml')?.text ?? '';
        assert.ok(styles.includes('<sz val="20"/>'), styles);
        assert.ok(styles.includes('formatCode="#,##0.00"'), styles);
    });

    it('gives one index to one style, however many cells ask for it', async () => {
        const { sheet, byName } = await readXlsx(
            write({ styles: { title: { bold: true } } }, [
                [{ v: 'a', s: 'title' }, { v: 'b', s: 'title' }, { v: 'c', s: { bold: true } }],
            ]),
        );
        assert.equal((sheet.match(/ s="1"/g) ?? []).length, 3);
        assert.equal((byName.get('xl/styles.xml')?.text.match(/<xf /g) ?? []).length, 3); // 1 in cellStyleXfs, 2 in cellXfs
    });

    it('says what a style nobody declared was, and writes no file', () => {
        assert.throws(() => write({}, [[{ v: 1, s: 'nope' }]]), /Unknown style "nope"/);
    });

    it('formats a date so it is not read back as a number', async () => {
        const { sheet, byName } = await readXlsx(write({}, [[new Date(2024, 0, 15)]]));
        assert.ok(sheet.includes(' s="1"'), sheet);
        const styles = byName.get('xl/styles.xml')?.text ?? '';
        // The built-in short date: an id, so there is no format code in the
        // file at all — what the day looks like is the reader's own.
        assert.match(styles, /<cellXfs[^>]*><xf [^>]*\/><xf numFmtId="14"/, styles);
        assert.ok(!styles.includes('numFmts'), styles);
    });

    it('refuses a date format nothing can be added to before it writes a byte', () => {
        const { sink, bytes } = recordingSink();
        assert.throws(() => new XlsxWriter(sink, { dateFormat: 15 }), /dateTimeFormat/);
        assert.equal(bytes().length, 0);
    });

    it('takes the date format the workbook asked for, time of day and all', async () => {
        const { byName } = await readXlsx(
            write({ dateFormat: 'dd/mm/yyyy' }, [
                [new Date(2024, 0, 15)],
                [new Date(2024, 0, 15, 12, 30)],
            ]),
        );
        const styles = byName.get('xl/styles.xml')?.text ?? '';
        assert.ok(styles.includes('formatCode="dd/mm/yyyy"'), styles);
        assert.ok(styles.includes('formatCode="dd/mm/yyyy hh:mm:ss"'), styles);
    });
});

describe('XlsxWriter: the columns of the sheet', () => {
    it('writes what the options declare, before the first row', async () => {
        const { sheet } = await readXlsx(
            write({ columnFormats: { B: { width: 24 }, D: { hidden: true } } }, [['a']]),
        );
        assert.ok(sheet.includes('<col min="2" max="2" width="24" customWidth="1"/>'), sheet);
        assert.ok(sheet.includes('<col min="4" max="4" hidden="1"/>'), sheet);
        assert.ok(sheet.indexOf('<cols>') < sheet.indexOf('<sheetData>'));
    });

    it('lets a #worksheet command give the sheet its own', async () => {
        const { sheets } = await readXlsx(
            write({ columnFormats: [{ width: 8 }] }, [
                ['first'],
                { '#worksheet': 'Wide', columnFormats: [{ width: 40 }] },
                ['second'],
            ]),
        );
        assert.ok(sheets[0]?.includes('width="8"'), sheets[0] ?? 'no first sheet');
        assert.ok(sheets[1]?.includes('width="40"'), sheets[1] ?? 'no second sheet');
    });

    it('falls back to the options for the sheet that declares none', async () => {
        const { sheets } = await readXlsx(
            write({ columnFormats: [{ width: 8 }] }, [['first'], { '#worksheet': 'Same' }, ['second']]),
        );
        assert.ok(sheets[1]?.includes('width="8"'), sheets[1] ?? 'no second sheet');
    });

    it('registers a column style in the same table the cells use', async () => {
        const { sheet, byName } = await readXlsx(
            write({ styles: { money: { numFmt: '#,##0.00' } }, columnFormats: { C: { s: 'money' } } }, [
                [{ v: 1, s: 'money' }],
            ]),
        );
        assert.ok(sheet.includes('<col min="3" max="3" style="1"/>'), sheet);
        assert.ok(sheet.includes('<c r="A1" s="1">'), sheet);
        assert.equal((byName.get('xl/styles.xml')?.text.match(/formatCode/g) ?? []).length, 1);
    });
});

describe('XlsxWriter: columns sized by what they hold', () => {
    it('writes nothing about the columns when nobody asked for a width', async () => {
        const { sheet } = await readXlsx(write({}, [['a very long line of text']]));
        assert.ok(!sheet.includes('<cols>'), sheet);
    });

    it('writes no <cols> for a sheet that had nothing to measure', async () => {
        // An empty `<cols/>` is a sheet Excel refuses to open.
        const { sheet } = await readXlsx(write({ autoWidthMax: 20 }, [[null, '']]));
        assert.ok(!sheet.includes('cols'), sheet);
    });

    it('sizes each column by its longest cell, header row included', async () => {
        const { sheet } = await readXlsx(
            write({ columns: COLUMNS, autoWidthMax: 50 }, [
                { id: 1, full_name: 'Ana' },
                { id: 22222, full_name: 'Bernardino' },
            ]),
        );
        // "id" is the header, and 22222 is longer; "Full name" is 9, and
        // "Bernardino" is 10.
        assert.ok(sheet.includes('<col min="1" max="1" width="5" customWidth="1"/>'), sheet);
        assert.ok(sheet.includes('<col min="2" max="2" width="10" customWidth="1"/>'), sheet);
    });

    it('keeps the <cols> ahead of the rows it was worked out from', async () => {
        const { sheet } = await readXlsx(write({ autoWidthMax: 20 }, [['abc']]));
        assert.ok(sheet.indexOf('<cols>') < sheet.indexOf('<sheetData>'), sheet);
        assert.ok(sheet.indexOf('<sheetData>') < sheet.indexOf('<row '), sheet);
    });

    it('stops a column at the maximum, and leaves the shorter ones alone', async () => {
        const { sheet } = await readXlsx(
            write({ autoWidthMax: 8 }, [['abcd', 'a line nobody would want a column of']]),
        );
        assert.ok(sheet.includes('<col min="1" max="1" width="4" customWidth="1"/>'), sheet);
        assert.ok(sheet.includes('<col min="2" max="2" width="8" customWidth="1"/>'), sheet);
    });

    it('leaves a column that was given a width at the one it was given', async () => {
        const { sheet } = await readXlsx(
            write({ autoWidthMax: 50, columnFormats: { A: { width: 3 } } }, [
                ['a much longer line', 'other'],
            ]),
        );
        assert.ok(sheet.includes('<col min="1" max="1" width="3" customWidth="1"/>'), sheet);
        assert.ok(sheet.includes('<col min="2" max="2" width="5" customWidth="1"/>'), sheet);
    });

    it('measures a column whose format says everything but the width', async () => {
        const { sheet } = await readXlsx(
            write({ autoWidthMax: 50, columnFormats: { A: { hidden: true } } }, [['abcdef']]),
        );
        assert.ok(sheet.includes('<col min="1" max="1" width="6" customWidth="1" hidden="1"/>'), sheet);
    });

    it('measures the column a cell says it goes in, not the one it came in', async () => {
        const { sheet } = await readXlsx(
            write({ autoWidthMax: 50 }, [[{ v: 'abcdefg', col: 'D' }]]),
        );
        assert.ok(sheet.includes('<col min="4" max="4" width="7" customWidth="1"/>'), sheet);
        assert.ok(!sheet.includes('min="1"'), sheet);
    });

    it('measures the result a formula carries, and nothing when it carries none', async () => {
        const { sheet } = await readXlsx(
            write({ autoWidthMax: 50 }, [[{ f: 'SUM(B1:B9)', v: 1234.5 }, { f: 'NOW()' }]]),
        );
        assert.ok(sheet.includes('<col min="1" max="1" width="6" customWidth="1"/>'), sheet);
        assert.ok(!sheet.includes('min="2"'), sheet);
    });

    it('gives every sheet its own widths, and takes the max from the command that opened it', async () => {
        const { sheets } = await readXlsx(
            write({ autoWidthMax: 50 }, [
                ['abcdef'],
                { '#worksheet': 'Narrow', autoWidthMax: 2 },
                ['abcdef'],
                { '#worksheet': 'Plain', autoWidthMax: undefined },
                ['abcdef'],
            ]),
        );
        assert.ok(sheets[0]?.includes('width="6"'), sheets[0] ?? 'no first sheet');
        assert.ok(sheets[1]?.includes('width="2"'), sheets[1] ?? 'no second sheet');
        // `undefined` is a field the command left out, so the workbook's own
        // max is what the third sheet gets.
        assert.ok(sheets[2]?.includes('width="6"'), sheets[2] ?? 'no third sheet');
    });

    it('leaves a blank sheet behind when one #worksheet follows another', async () => {
        const { sheetNames, sheets } = await readXlsx(
            write({ autoWidthMax: 20 }, [
                { '#worksheet': 'hoja1' },
                { '#worksheet': 'hoja2' },
                ['abc'],
            ]),
        );
        // The first command opens the first sheet; the second one closes it
        // with nothing in it, which is a sheet in the workbook all the same.
        assert.deepEqual(sheetNames, ['hoja1', 'hoja2']);
        assert.ok(sheets[0]?.includes('<sheetData></sheetData>'), sheets[0] ?? 'no first sheet');
        assert.ok(!sheets[0]?.includes('cols'), sheets[0] ?? 'no first sheet');
        assert.ok(sheets[1]?.includes('width="3"'), sheets[1] ?? 'no second sheet');
    });

    it('writes the sheet whole, in the place it would have taken anyway', async () => {
        const { names, sheet } = await readXlsx(
            write({ autoWidthMax: 20 }, [['abc'], { '#worksheet': 'Second' }, ['de']]),
        );
        assert.deepEqual(names, [
            '[Content_Types].xml',
            '_rels/.rels',
            SHEET_PART,
            'xl/worksheets/sheet2.xml',
            'xl/styles.xml',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
        ]);
        assert.ok(sheet.endsWith('</sheetData></worksheet>'), sheet);
    });

    it('refuses a maximum no column can be sized by', () => {
        assert.throws(() => write({ autoWidthMax: 0 }, [['a']]), /autoWidthMax/);
    });
});
