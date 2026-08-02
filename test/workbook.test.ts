// End to end, read back by a reader that had nothing to do with writing:
// `exceljs` for the workbook, `yauzl` plus Node's own `zlib` for the
// container. This is `scripts/validate-xlsx.ts`'s job, run against files the
// tests generate rather than against the examples' output.
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
import type { Column, Row } from '../src/core/types.js';
import { collect } from './helpers/streams.js';
import { MAX_VERSION, METHOD_DEFLATE, SHEET_PART, ZIP64_VERSION, readZipEntries } from './helpers/zip.js';

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

    it('writes a Date as its Excel serial number, and as nothing more', () => {
        // Without a date `numFmt` — which the writer does not do yet — the
        // value is right and a reader has no way to know it is a date.
        assert.equal(sheet.getRow(2).getCell(4).value, WHEN.getTime() / 86400000 + 25569);
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
                    ['another', 3],
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
        assert.equal(notas?.rowCount, 2); // no header of its own
        assert.equal(totals?.getRow(2).getCell(2).value, 7.5);
        assert.equal(notas?.getRow(1).getCell(1).value, 'a note');
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
