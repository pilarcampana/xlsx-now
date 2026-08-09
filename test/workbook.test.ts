// End to end, read back by a reader that had nothing to do with writing:
// `exceljs` for the workbook, `yauzl` plus Node's own `zlib` for the
// container. This is `scripts/validate-xlsx.ts`'s job, run against files the
// tests generate rather than against the examples' output.
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { columnWidth } from '../src/core/autoWidth.js';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
import { DEFAULT_DATETIME_FORMAT } from '../src/core/styles.js';
import type { Column, Row } from '../src/core/types.js';
import { defaultTypes, withType } from '../src/core/valueTypes.js';
import { collect } from './helpers/streams.js';
import { MAX_VERSION, METHOD_DEFLATE, SHEET_PART, ZIP64_VERSION, readXlsx, readZipEntries } from './helpers/zip.js';

const PK_FILL_ARGB = 'FFFFE699';

const COLUMNS: readonly Column[] = [
    { name: 'id', pk: true },
    { name: 'Full name', key: 'full_name' },
    { name: 'score' },
    { name: 'when' },
    { name: 'active' },
];

const WHEN = new Date('2024-01-15T12:30:00.000Z');

const RECORDS: Row[] = [
    { id: 1, full_name: 'Ana & Co <1>', score: 10.5, when: WHEN, active: true },
    { id: 2, full_name: 'Beto', score: -3, when: WHEN, active: false },
    { id: 3, full_name: null, score: 0, when: null, active: null },
];

async function generate(columns: readonly Column[], rows: Row[]): Promise<Buffer> {
    return collect(createXlsxStream({ columns, rows, sheetName: 'People' }));
}

async function open(bytes: Buffer): Promise<ExcelJS.Worksheet> {
    const workbook = new ExcelJS.Workbook();
    // `exceljs` declares its own, older `Buffer`; the bytes are the same ones.
    await workbook.xlsx.load(bytes as unknown as Parameters<typeof workbook.xlsx.load>[0]);
    const sheet = workbook.worksheets[0];
    assert.ok(sheet, 'the workbook has no worksheets');
    return sheet;
}

function isPkFilled(cell: ExcelJS.Cell): boolean {
    return cell.fill?.type === 'pattern' && cell.fill.fgColor?.argb === PK_FILL_ARGB;
}

/** The sheet's first view, as the freeze it may or may not be. */
function view(sheet: ExcelJS.Worksheet): { state?: string; xSplit?: number; ySplit?: number } {
    return sheet.views[0] ?? {};
}

describe('a generated workbook, read back with exceljs', () => {
    let sheet: ExcelJS.Worksheet;

    before(async () => {
        sheet = await open(await generate(COLUMNS, RECORDS));
    });

    it('carries the sheet under the name it was given', () => {
        assert.equal(sheet.name, 'People');
    });

    it('has a header row plus one row per record', () => {
        assert.equal(sheet.rowCount, RECORDS.length + 1);
        assert.equal(sheet.columnCount, COLUMNS.length);
    });

    it('heads every column with its name, in bold', () => {
        const header = sheet.getRow(1);
        for (const [index, column] of COLUMNS.entries()) {
            const cell = header.getCell(index + 1);
            assert.equal(cell.value, column.name);
            assert.ok(cell.font?.bold, `header cell ${cell.address} is not bold`);
        }
    });

    it('fills the pk column, in the header and in the data alike', () => {
        assert.ok(isPkFilled(sheet.getRow(1).getCell(1)), 'the pk header is not filled');
        assert.ok(isPkFilled(sheet.getRow(2).getCell(1)), 'a pk cell is not filled');
    });

    it('leaves the other columns unstyled', () => {
        const cell = sheet.getRow(2).getCell(2);
        assert.ok(cell.fill === undefined || cell.fill.type === undefined, 'a plain cell is filled');
        assert.ok(!cell.font?.bold, 'a data row is bold');
    });

    it('reads every value back as the type it went in as', () => {
        const row = sheet.getRow(2);
        assert.equal(row.getCell(1).value, 1);
        assert.equal(row.getCell(2).value, 'Ana & Co <1>');
        assert.equal(row.getCell(3).value, 10.5);
        assert.equal(row.getCell(5).value, true);
    });

    it('writes a Date as a date, and a reader reads one back', () => {
        // The serial alone would be right and unreadable: a number with no
        // format is shown as the five digits it is, so a date that asked for
        // no style of its own gets the one that makes it a date.
        const value = sheet.getRow(2).getCell(4).value;
        assert.ok(value instanceof Date, `read back as ${typeof value}`);
        assert.equal(value.getTime(), WHEN.getTime());
    });

    it('shows the time too, since the value carries one', () => {
        // What the file carries is an id — the built-in short date and time —
        // and not a format code: the spelling of a date is the reader's own,
        // and `exceljs` answers with its own table's entry for that id.
        const numFmt = sheet.getRow(2).getCell(4).numFmt;
        assert.match(
            numFmt,
            /y.*h/i,
            `built-in ${DEFAULT_DATETIME_FORMAT} was read back as "${numFmt}"`,
        );
    });

    it('reads a negative number and a false back too', () => {
        const row = sheet.getRow(3);
        assert.equal(row.getCell(3).value, -3);
        assert.equal(row.getCell(5).value, false);
    });

    it('leaves an empty value empty, and a zero a zero', () => {
        const row = sheet.getRow(4);
        assert.equal(row.getCell(2).value, null);
        assert.equal(row.getCell(3).value, 0);
        assert.equal(row.getCell(4).value, null);
        // The pk column is styled, so its cell is there even when empty.
        assert.ok(isPkFilled(row.getCell(1)));
    });

    it('freezes the header row and the leading pk column', () => {
        assert.equal(view(sheet).state, 'frozen');
        assert.equal(view(sheet).ySplit, 1);
        assert.equal(view(sheet).xSplit, 1);
    });
});

describe('a generated workbook, with the pk columns elsewhere', () => {
    it('freezes no column when a pk is not one of the first', async () => {
        const sheet = await open(
            await generate([{ name: 'name' }, { name: 'id', pk: true }], [{ name: 'Ana', id: 1 }]),
        );
        assert.equal(view(sheet).state, 'frozen');
        assert.equal(view(sheet).ySplit, 1);
        assert.ok(!view(sheet).xSplit, `xSplit is ${JSON.stringify(view(sheet))}`);
    });

    it('freezes no column when every column is a pk', async () => {
        const sheet = await open(
            await generate(
                [{ name: 'id', pk: true }, { name: 'year', pk: true }],
                [{ id: 1, year: 2024 }],
            ),
        );
        assert.ok(!view(sheet).xSplit, `xSplit is ${JSON.stringify(view(sheet))}`);
    });
});

describe('a workbook of several sheets, read back with exceljs', () => {
    let workbook: ExcelJS.Workbook;

    before(async () => {
        // One stream, three sheets: the first named in the options, the other
        // two by the commands that open them — one of them with columns of
        // its own, the other with none at all.
        const bytes = await collect(
            createXlsxStream({
                sheetName: 'People',
                columns: COLUMNS,
                rows: [
                    RECORDS[0] as Row,
                    { '#worksheet': 'Totals', columns: [{ name: 'label', pk: true }, { name: 'total' }] },
                    { label: 'sum', total: 7.5 },
                    { '#worksheet': 'Notas', columns: [], freezeRows: 0 },
                    ['a note', 2],
                    { '#line': 'empty' },
                    { '#line': 'array', values: ['a tall title'], height: 30, s: { bold: true } },
                    { '#line': 'array', values: [{ v: 'far right', col: 'C' }] },
                    { '#line': 'array', values: ['out of sight'], hidden: true },
                ],
            }),
        );
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(bytes as unknown as Parameters<typeof workbook.xlsx.load>[0]);
    });

    it('carries every sheet, in the order the stream declared them', () => {
        assert.deepEqual(
            workbook.worksheets.map((sheet) => sheet.name),
            ['People', 'Totals', 'Notas'],
        );
    });

    it('sends every row to the sheet that was open when it arrived', () => {
        const [people, totals, notas] = workbook.worksheets;
        assert.equal(people?.rowCount, 2); // header + one record
        assert.equal(totals?.rowCount, 2);
        assert.equal(notas?.rowCount, 5); // no header of its own
        assert.equal(totals?.getRow(2).getCell(2).value, 7.5);
        assert.equal(notas?.getRow(1).getCell(1).value, 'a note');
    });

    it('reads back what a #line asked of the row itself', () => {
        const notas = workbook.worksheets[2];
        assert.ok(notas, 'the third sheet is missing');
        // Row 2 is the empty line, so the title is row 3.
        const title = notas.getRow(3);
        assert.equal(title.getCell(1).value, 'a tall title');
        assert.equal(title.height, 30);
        assert.ok(title.font?.bold, 'the row style did not reach the row');
        assert.equal(notas.getRow(5).hidden, true);
    });

    it('puts a sparse line at the column its letter names', () => {
        const notas = workbook.worksheets[2];
        const sparse = notas?.getRow(4);
        assert.equal(sparse?.getCell(3).value, 'far right');
        assert.equal(sparse?.getCell(1).value, null, 'the gap was filled in');
    });

    it('heads each sheet with its own columns, and styles them as its own', () => {
        const totals = workbook.worksheets[1];
        assert.equal(totals?.getRow(1).getCell(1).value, 'label');
        assert.ok(totals?.getRow(1).getCell(1).font?.bold, 'the header is not bold');
        assert.ok(isPkFilled(totals.getRow(2).getCell(1)), 'the pk cell is not filled');
    });

    it('freezes each sheet on its own terms', () => {
        const [people, totals, notas] = workbook.worksheets;
        assert.equal(view(people as ExcelJS.Worksheet).ySplit, 1);
        assert.equal(view(totals as ExcelJS.Worksheet).ySplit, 1);
        assert.equal(view(totals as ExcelJS.Worksheet).xSplit, 1);
        assert.notEqual(view(notas as ExcelJS.Worksheet).state, 'frozen');
    });
});

describe('a generated container, read back with yauzl', () => {
    it('is a ZIP Office can open, written without knowing the sizes', async () => {
        const entries = await readZipEntries(await generate(COLUMNS, RECORDS));

        for (const entry of entries) {
            assert.ok(entry.crcMatches, `${entry.name}: CRC mismatch`);
            assert.ok(
                entry.madeByVersion <= MAX_VERSION,
                `${entry.name}: made with ZIP ${entry.madeByVersion / 10}`,
            );
            assert.ok(
                entry.versionNeededToExtract < ZIP64_VERSION,
                `${entry.name}: ZIP64 entry`,
            );
            assert.ok(entry.streamed, `${entry.name}: sizes were known upfront`);
            assert.ok(entry.localSizesZeroed, `${entry.name}: local header carries sizes`);
        }

        const sheet = entries.find((entry) => entry.name === SHEET_PART);
        assert.ok(sheet, `missing ${SHEET_PART}`);
        assert.equal(sheet.method, METHOD_DEFLATE);
    });

    it('has no ZIP64 record at the end of it', async () => {
        const bytes = await generate(COLUMNS, RECORDS);
        const tail = bytes.subarray(-128);
        assert.ok(!tail.includes(Buffer.from('PK\x06\x06')), 'ZIP64 end of central directory');
        assert.ok(!tail.includes(Buffer.from('PK\x06\x07')), 'ZIP64 locator');
    });
});

describe('a workbook whose columns sized themselves, read back with exceljs', () => {
    let sheet: ExcelJS.Worksheet;

    before(async () => {
        const bytes = await collect(
            createXlsxStream({
                sheetName: 'People',
                columns: COLUMNS,
                autoWidthMax: 12,
                // C is said outright, so it stays at 30 however short its
                // cells turn out to be.
                columnFormats: { C: { width: 30 } },
                rows: RECORDS,
            }),
        );
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(bytes as unknown as Parameters<typeof workbook.xlsx.load>[0]);
        const first = workbook.worksheets[0];
        assert.ok(first, 'the workbook has no worksheets');
        sheet = first;
    });

    it('gives each column the width of its longest cell, up to the maximum', () => {
        // The width a column is written with is the characters it measured
        // plus the padding Excel measures a column by, which is what keeps
        // the value it was sized for from being clipped.
        // "id" against 1, 2 and 3; "Full name" against "Ana & Co <1>", which
        // is 12 and reaches the maximum exactly.
        assert.equal(sheet.getColumn(1).width, columnWidth(2));
        assert.equal(sheet.getColumn(2).width, columnWidth(12));
        // The dates are longer than the maximum.
        assert.equal(sheet.getColumn(4).width, columnWidth(12));
        // TRUE, FALSE, and the header longer than both.
        assert.equal(sheet.getColumn(5).width, columnWidth(6));
    });

    it('leaves the column that was given a width at the one it was given', () => {
        assert.equal(sheet.getColumn(3).width, 30);
    });

    it('still holds every row it was handed', () => {
        assert.equal(sheet.rowCount, RECORDS.length + 1);
        assert.equal(sheet.getRow(2).getCell(2).value, 'Ana & Co <1>');
    });
});

describe('a workbook that was taught a type of its own, read back with exceljs', () => {
    class HourRange {
        constructor(
            readonly from: string,
            readonly to: string,
        ) {}
        toString(): string {
            return `${this.from} a ${this.to}`;
        }
    }

    const TYPES = withType(defaultTypes, HourRange, {
        convert: (range) => ({ v: range.toString() }),
    });

    let sheet: ExcelJS.Worksheet;

    before(async () => {
        sheet = await open(
            await collect(
                createXlsxStream({
                    columns: [{ name: 'class' }, { name: 'hours' }, { name: 'seats' }],
                    rows: [
                        { class: 'Bases de datos', hours: new HourRange('8:00', '10:30'), seats: 60n },
                        { class: 'Algoritmos', hours: new HourRange('14:00', '16:00'), seats: 45n },
                    ],
                    types: TYPES,
                }),
            ),
        );
    });

    it('writes an instance of the class as what its conversion made of it', () => {
        assert.equal(sheet.getRow(2).getCell(2).value, '8:00 a 10:30');
        assert.equal(sheet.getRow(3).getCell(2).value, '14:00 a 16:00');
    });

    it('still knows the types it was based on', () => {
        assert.equal(sheet.getRow(2).getCell(3).value, 60);
    });

    it('refuses the same class when the workbook was never given it', async () => {
        await assert.rejects(
            collect(
                createXlsxStream({
                    columns: [{ name: 'hours' }],
                    rows: [{ hours: new HourRange('8:00', '10:30') }],
                }),
            ),
            /"HourRange" is not one of them/,
        );
    });
});

describe('a row style and a column style, read back with exceljs', () => {
    // The shape of a file made by hand in Excel: column A bold, row 1 bold,
    // one cell bold on its own. Excel writes the style on every cell it
    // reaches — a `<c>` with no `s` is style 0, not a cell waiting to inherit
    // one — and keeps the `<row>` and `<col>` for the cells that are not in
    // the file at all.
    let sheet: ExcelJS.Worksheet;
    let xml: string;

    before(async () => {
        const bytes = await collect(
            createXlsxStream({
                columnFormats: [{ s: { bold: true } }],
                rows: [
                    {
                        '#line': 'array',
                        values: ['both', 'by line', 'by line 2'],
                        s: { bold: true },
                    },
                    ['by column', 'normal', 'normal 2'],
                    ['by column 2', 'normal 3', { v: 'in cell', s: { bold: true } }],
                    [{ v: 'both, merged', s: { italic: true } }],
                ],
            } as Parameters<typeof createXlsxStream>[0]),
        );
        xml = (await readXlsx(bytes)).sheet;
        sheet = await open(bytes as Buffer);
    });

    /** Which cells of the first three rows and columns came back bold. */
    function bold(): string[] {
        return [1, 2, 3].map((row) =>
            [1, 2, 3].map((col) => (sheet.getRow(row).getCell(col).font?.bold ? 'B' : '.')).join(''),
        );
    }

    it('reaches every cell the row, the column and the cell itself ask for', () => {
        assert.deepEqual(bold(), ['BBB', 'B..', 'B.B']);
    });

    it('writes the style on the cells, which is the only place it is read from', () => {
        assert.match(xml, /<c r="A1"[^>]* s="1"/);
        assert.match(xml, /<c r="C3"[^>]* s="1"/);
    });

    it('merges what the column says with what the cell says, rather than one winning', () => {
        // A4 falls under the bold column and asks for italic on its own: a
        // cell that says one thing does not give up what it was already under.
        const cell = sheet.getRow(4).getCell(1);
        assert.ok(cell.font?.bold, 'lost the bold the column gave it');
        assert.ok(cell.font?.italic, 'lost the italic it asked for');
    });

    it('leaves a cell that falls under nothing with no style attribute at all', () => {
        // Not `s="0"`: the default is what a `<c>` says by saying nothing.
        assert.match(xml, /<c r="B2" t="inlineStr">/);
    });

    it('still says it on the row and the column, for the cells not in the file', () => {
        // This is what makes a value typed into an empty cell of that row —
        // or that column — come out styled, as it does in Excel.
        assert.match(xml, /<row r="1" s="\d+" customFormat="1">/);
        assert.match(xml, /<col min="1" max="1" style="\d+"\/>/);
    });
});

describe('a workbook of merged cells, read back with exceljs', () => {
    // A header spanning three columns, a category spanning three rows, and a
    // block spanning both — the three shapes Excel merges.
    let sheet: ExcelJS.Worksheet;
    let xml: string;

    before(async () => {
        const bytes = await collect(
            createXlsxStream({
                rows: [
                    [{ v: 'Ventas 2024', s: { bold: true, border: { all: 'thin' } }, colSpan: 3 }],
                    [{ v: 'Fruta', rowSpan: 3 }, 'manzana', 10],
                    [undefined, 'pera', 20],
                    [undefined, 'banana', 30],
                    [{ v: 'Total', colSpan: 2, rowSpan: 2 }, undefined, 60],
                    [undefined, undefined, 15],
                ],
            } as Parameters<typeof createXlsxStream>[0]),
        );
        xml = (await readXlsx(bytes)).sheet;
        sheet = await open(bytes as Buffer);
    });

    it('comes back as the ranges it was written as', () => {
        assert.deepEqual(sheet.model.merges, ['A1:C1', 'A2:A4', 'A5:B6']);
    });

    it('puts the merged ranges after the sheet data, which is where they go', () => {
        assert.match(xml, /<\/sheetData><mergeCells count="3">/);
    });

    it('shows the value of the first cell of the range, and nothing under it', () => {
        assert.equal(sheet.getCell('A1').value, 'Ventas 2024');
        assert.equal(sheet.getCell('A2').value, 'Fruta');
        // exceljs reads a covered cell as bound to the one that declared the
        // merge, which is the same thing Excel shows.
        assert.equal(sheet.getCell('B1').master.address, 'A1');
        assert.equal(sheet.getCell('A4').master.address, 'A2');
    });

    it('leaves the rest of the sheet where the merges did not move it', () => {
        assert.equal(sheet.getCell('B3').value, 'pera');
        assert.equal(sheet.getCell('C5').value, 60);
    });

    it('carries the border of the range around the whole of it', () => {
        // The one thing a merge needs from the cells it covers: xlsx has no
        // border around a range, only borders around cells.
        for (const ref of ['A1', 'B1', 'C1']) {
            assert.ok(sheet.getCell(ref).border?.top, `${ref} lost the top border`);
        }
    });
});

describe('a workbook whose data carried what XML cannot, read back with exceljs', () => {
    // One character XML 1.0 has no place for is enough to make the whole file
    // unreadable — and it arrives in the data, not in the code: a truncated
    // text field, a stray byte of something binary. The file has to open.
    const DIRTY = `ab${String.fromCharCode(0)}cd${String.fromCharCode(7)}ef`;

    it('opens, and shows the value without what could not be written', async () => {
        const bytes = await collect(
            createXlsxStream({
                sheetName: `Hoja${String.fromCharCode(0)}1`,
                rows: [
                    [DIRTY],
                    [{ v: DIRTY, s: { numFmt: `0"${DIRTY}"` } }],
                    [{ v: 1, f: `CONCATENATE("${DIRTY}")` }],
                ],
            } as Parameters<typeof createXlsxStream>[0]),
        );
        // Reading it back at all is the assertion: exceljs parses the XML, and
        // before this it failed with "disallowed character".
        const sheet = await open(bytes as Buffer);
        assert.equal(sheet.getCell('A1').value, 'abcdef');
        assert.equal(sheet.getCell('A2').value, 'abcdef');
        assert.equal(sheet.name, 'Hoja1');
    });

    it('leaves the file with none of those characters in it', async () => {
        const bytes = await collect(
            createXlsxStream({ rows: [[DIRTY]] } as Parameters<typeof createXlsxStream>[0]),
        );
        const { sheet } = await readXlsx(bytes);
        assert.doesNotMatch(sheet, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/);
    });
});
