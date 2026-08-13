// The reader end to end, against three kinds of file: the ones this package
// writes, ones built by hand here, and one written by `exceljs` — which is
// the only one of the three that had nothing to do with this repository, and
// so the only one that can say the reader agrees with anything but itself.
import 'temporal-polyfill/global';
import assert from 'node:assert/strict';
import { Temporal } from 'temporal-polyfill';
import ExcelJS from 'exceljs';
import { createXlsxStream, type CreateXlsxStreamOptions } from '../../src/core/createXlsxStream.js';
import { openXlsx, readXlsx } from '../../src/core/read/readXlsx.js';
import type { ReadDates } from '../../src/core/read/dates.js';
import type { ReadValue } from '../../src/core/read/types.js';
import type { StyledCell } from '../../src/core/types.js';
import { stylesOf, xlsxPackage, zipOf } from '../helpers/package.js';
import { collect } from '../helpers/streams.js';

/** A workbook this package wrote, as bytes. */
async function written(options: CreateXlsxStreamOptions): Promise<Uint8Array> {
    return new Uint8Array(await collect(createXlsxStream(options)));
}

describe('readXlsx: what this package wrote', () => {
    it('reads the values back as they went in', async () => {
        const bytes = await written({
            sheetName: 'Datos',
            rows: [
                ['texto', 42, true],
                ['otra', -3.5, false],
            ],
        });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.name, 'Datos');
        assert.deepEqual(sheet?.cells, [
            ['texto', 42, true],
            ['otra', -3.5, false],
        ]);
        assert.deepEqual([sheet?.maxRow, sheet?.maxCol], [2, 3]);
    });

    it('reads a date back as the same date', async () => {
        const dates = [new Date(2024, 0, 15), new Date(2024, 6, 1, 12, 30), new Date(1900, 0, 1)];
        const bytes = await written({ rows: [dates] });
        // The writer reads a `Date` by the local clock and `localDate` is the
        // same clock the other way round, so this is the round trip.
        const [sheet] = await readXlsx(bytes, { dates: 'localDate' });
        assert.deepEqual(
            sheet?.cells[0]?.map((value) => (value as Date).getTime()),
            dates.map((date) => date.getTime()),
        );
    });

    it('reads a date as a Temporal value when nothing said otherwise', async () => {
        const bytes = await written({
            rows: [
                [
                    Temporal.PlainDate.from('2024-01-15'),
                    Temporal.PlainDateTime.from('2024-01-15T12:00'),
                    Temporal.PlainTime.from('10:30'),
                ],
            ],
        });
        const [sheet] = await readXlsx(bytes);
        assert.deepEqual(
            sheet?.cells[0]?.map((value) => String(value)),
            ['2024-01-15', '2024-01-15T12:00:00', '10:30:00'],
        );
        const [day, when, time] = sheet?.cells[0] ?? [];
        assert.ok(day instanceof Temporal.PlainDate);
        assert.ok(when instanceof Temporal.PlainDateTime);
        assert.ok(time instanceof Temporal.PlainTime);
    });

    it('reads the same date four ways, and they all say the same thing', async () => {
        const bytes = await written({ rows: [[Temporal.PlainDateTime.from('2024-01-15T12:00')]] });
        const cellOf = async (dates: ReadDates): Promise<ReadValue | undefined> =>
            (await readXlsx(bytes, { dates }))[0]?.cells[0]?.[0];

        assert.equal(String(await cellOf('temporal')), '2024-01-15T12:00:00');
        assert.equal(await cellOf('isoString'), '2024-01-15T12:00:00');
        assert.equal((await cellOf('utcDate') as Date).toISOString(), '2024-01-15T12:00:00.000Z');
        const local = (await cellOf('localDate')) as Date;
        assert.deepEqual(
            [local.getFullYear(), local.getMonth(), local.getDate(), local.getHours()],
            [2024, 0, 15, 12],
        );
    });

    it('gives back Temporal values a workbook can be written from again', async () => {
        const rows = [
            [
                Temporal.PlainDate.from('2024-01-15'),
                Temporal.PlainDateTime.from('2024-01-15T12:00'),
                Temporal.PlainTime.from('10:30'),
            ],
        ];
        const once = await readXlsx(await written({ rows }), { mode: 'cells' });
        const again = await readXlsx(await written({ rows: once[0]?.cells ?? [] }), {
            mode: 'cells',
        });
        assert.deepEqual(again[0]?.cells, once[0]?.cells);
        // And the formats came back with them: a day, a day and an hour, and
        // an hour on its own are three different things to show.
        assert.deepEqual(
            (once[0]?.cells[0] ?? []).map((cell) => (cell as StyledCell).s),
            [{ numFmt: 14 }, { numFmt: 22 }, { numFmt: 21 }],
        );
    });

    it('keeps a hole a hole, and tells it from a cell that holds nothing', async () => {
        const bytes = await written({
            rows: [['a', undefined, { v: null, s: { numFmt: '0.00' } }, 'd']],
        });
        const [sheet] = await readXlsx(bytes);
        const row = sheet?.cells[0] ?? [];
        assert.equal(row[0], 'a');
        assert.equal(row[1], undefined, 'a position with no cell should stay empty');
        assert.equal(row[2], null, 'a cell that holds nothing should be null');
        assert.equal(row[3], 'd');
    });

    it('reads every sheet, in the order the workbook declares them', async () => {
        const bytes = await written({
            sheets: [
                { name: 'Uno', rows: [[1]] },
                { name: 'Dos', rows: [[2]] },
                { name: 'Tres', rows: [[3]] },
            ],
        });
        const sheets = await readXlsx(bytes);
        assert.deepEqual(sheets.map((sheet) => sheet.name), ['Uno', 'Dos', 'Tres']);
        assert.deepEqual(sheets.map((sheet) => sheet.cells[0]?.[0]), [1, 2, 3]);
    });

    it('reads the columns mode as the header row it writes', async () => {
        const bytes = await written({
            columns: [{ name: 'nombre' }, { name: 'edad' }],
            rows: [{ nombre: 'Ana', edad: 33 }],
        });
        const [sheet] = await readXlsx(bytes);
        assert.deepEqual(sheet?.cells, [
            ['nombre', 'edad'],
            ['Ana', 33],
        ]);
    });

    it('gives back cells a workbook can be written from again', async () => {
        const bytes = await written({
            rows: [[{ v: 1234.5, s: { numFmt: '#,##0.00' } }, new Date(2024, 0, 15), 'texto']],
        });
        const [sheet] = await readXlsx(bytes, { mode: 'cells', dates: 'localDate' });
        const row = (sheet?.cells[0] ?? []) as StyledCell[];
        assert.deepEqual(row[0], { v: 1234.5, s: { numFmt: '#,##0.00' } });
        assert.deepEqual(row[1]?.s, { numFmt: 14 });
        assert.deepEqual(row[2], { v: 'texto' });

        // And what came out goes back in and reads the same, which is what
        // makes the two modes one round trip rather than two shapes.
        const again = await readXlsx(await written({ rows: [row] }), {
            mode: 'cells',
            dates: 'localDate',
        });
        assert.deepEqual(again[0]?.cells[0], sheet?.cells[0]);
    });

    it('carries a formula and the result the file cached for it', async () => {
        const bytes = await written({ rows: [[{ v: 3, f: 'SUM(A2:A3)' }]] });
        const [values] = await readXlsx(bytes);
        assert.equal(values?.cells[0]?.[0], 3);
        const [cells] = await readXlsx(bytes, { mode: 'cells' });
        assert.deepEqual(cells?.cells[0]?.[0], { v: 3, f: 'SUM(A2:A3)' });
    });

    it('reads a sheet a row reached across with col', async () => {
        const bytes = await written({ rows: [[{ v: 'lejos', col: 'E' }]] });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.maxCol, 5);
        assert.equal(sheet?.cells[0]?.[4], 'lejos');
        assert.equal(sheet?.cells[0]?.[0], undefined);
    });
});

describe('readXlsx: files built by hand', () => {
    it('reads a workbook that keeps its strings in the shared table', async () => {
        const bytes = xlsxPackage({
            sheets: {
                Hoja: '<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>',
            },
            sharedStrings: '<si><t>primero</t></si><si><t>segundo</t></si>',
        });
        const [sheet] = await readXlsx(bytes);
        assert.deepEqual(sheet?.cells, [['segundo', 'primero']]);
    });

    it('reads a workbook whose sheets are not in the order of their parts', async () => {
        const bytes = xlsxPackage({
            sheets: { Primera: '<row r="1"><c><v>1</v></c></row>', Segunda: '<row r="1"><c><v>2</v></c></row>' },
        });
        const sheets = await readXlsx(bytes);
        assert.deepEqual(sheets.map((sheet) => sheet.name), ['Primera', 'Segunda']);
    });

    it('reads a workbook that counts its days from 1904', async () => {
        const bytes = xlsxPackage({
            sheets: { H: '<row r="1"><c r="A1" s="0"><v>43844</v></c></row>' },
            styles: stylesOf([14]),
            workbookPr: 'date1904="1"',
        });
        const [sheet] = await readXlsx(bytes, { dates: 'localDate' });
        const date = sheet?.cells[0]?.[0] as Date;
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
    });

    it('reads an error cell as the error it shows', async () => {
        const bytes = xlsxPackage({
            sheets: { H: '<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>' },
        });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.cells[0]?.[0], '#DIV/0!');
        const [cells] = await readXlsx(bytes, { mode: 'cells' });
        assert.deepEqual(cells?.cells[0]?.[0], { v: '#DIV/0!', t: 'e' });
    });

    it('is dense in rows and ragged in columns', async () => {
        const bytes = xlsxPackage({
            sheets: {
                H:
                    '<row r="1"><c r="A1"><v>1</v></c></row>' +
                    '<row r="4"><c r="C4"><v>3</v></c></row>',
            },
        });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.maxRow, 4);
        assert.equal(sheet?.maxCol, 3);
        assert.deepEqual(sheet?.cells[1], [], 'a row the file skipped should be an empty row');
        assert.deepEqual(sheet?.cells[2], []);
        assert.equal(sheet?.cells[0]?.length, 1, 'a row should end at its own last cell');
    });

    it('reads a sheet with nothing in it', async () => {
        const [sheet] = await readXlsx(xlsxPackage({ sheets: { Vacía: '' } }));
        assert.deepEqual(sheet, { name: 'Vacía', cells: [], maxCol: 0, maxRow: 0 });
    });

    it('reads a file whose elements carry a namespace prefix', async () => {
        const bytes = xlsxPackage({
            sheets: { H: '<x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row>' },
            prefix: 'x',
        });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.cells[0]?.[0], 7);
    });
});

describe('readXlsx: what exceljs wrote', () => {
    /** A workbook written by an implementation that is not this one. */
    async function byExcelJs(build: (sheet: ExcelJS.Worksheet) => void): Promise<Uint8Array> {
        const workbook = new ExcelJS.Workbook();
        build(workbook.addWorksheet('Externa'));
        return new Uint8Array(await workbook.xlsx.writeBuffer());
    }

    it('reads the values of a file this package did not write', async () => {
        const bytes = await byExcelJs((sheet) => {
            sheet.addRow(['texto', 123, true]);
            sheet.addRow(['texto', 4.5, false]);
        });
        const [sheet] = await readXlsx(bytes);
        assert.equal(sheet?.name, 'Externa');
        assert.deepEqual(sheet?.cells, [
            ['texto', 123, true],
            ['texto', 4.5, false],
        ]);
    });

    it('reads a date exceljs wrote as the instant it wrote', async () => {
        const when = new Date(Date.UTC(2024, 0, 15, 12, 30));
        const bytes = await byExcelJs((sheet) => {
            sheet.addRow([when]);
        });
        const [sheet] = await readXlsx(bytes, { dates: 'utcDate' });
        const read = sheet?.cells[0]?.[0] as Date;
        assert.ok(read instanceof Date, 'a date came back as something else');
        // `exceljs` writes the serial of the UTC clock, so reading it back by
        // that same clock is the instant it wrote, to the millisecond — and
        // what matters here is that a serial under a date format was read as a
        // date at all.
        assert.equal(read.getTime(), when.getTime());
    });

    it('reads a formula and the value cached with it', async () => {
        const bytes = await byExcelJs((sheet) => {
            sheet.addRow([1]);
            sheet.addRow([2]);
            sheet.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 3 };
        });
        const [sheet] = await readXlsx(bytes, { mode: 'cells' });
        assert.deepEqual(sheet?.cells[2]?.[0], { v: 3, f: 'SUM(A1:A2)' });
    });

    it('reads the several sheets of a workbook it did not write', async () => {
        const workbook = new ExcelJS.Workbook();
        workbook.addWorksheet('Primera').addRow(['a']);
        workbook.addWorksheet('Segunda').addRow(['b']);
        const sheets = await readXlsx(new Uint8Array(await workbook.xlsx.writeBuffer()));
        assert.deepEqual(sheets.map((sheet) => sheet.name), ['Primera', 'Segunda']);
        assert.deepEqual(sheets.map((sheet) => sheet.cells[0]?.[0]), ['a', 'b']);
    });
});

describe('openXlsx', () => {
    it('names the sheets without reading a single row', async () => {
        const bytes = await written({
            sheets: [
                { name: 'Uno', rows: [[1]] },
                { name: 'Dos', rows: [[2]] },
            ],
        });
        const reader = await openXlsx(bytes);
        assert.deepEqual(reader.sheets.map((sheet) => sheet.name), ['Uno', 'Dos']);
    });

    it('reads one sheet without reading the ones before it', async () => {
        const bytes = await written({
            sheets: [
                { name: 'Uno', rows: [[1]] },
                { name: 'Dos', rows: [[2]] },
            ],
        });
        const reader = await openXlsx(bytes);
        const second = reader.sheets[1];
        assert.deepEqual((await second?.read())?.cells, [[2]]);
    });

    it('hands out the rows one at a time, and stops when the caller does', async () => {
        const bytes = await written({ rows: Array.from({ length: 500 }, (_row, index) => [index]) });
        const reader = await openXlsx(bytes);
        const first: number[] = [];
        for await (const row of reader.sheets[0]?.rows() ?? []) {
            first.push(row.cells[0] as number);
            if (first.length === 3) break;
        }
        assert.deepEqual(first, [0, 1, 2]);
    });

    it('reads the rows of a sheet by the numbers the sheet gives them', async () => {
        const bytes = xlsxPackage({
            sheets: { H: '<row r="1"><c><v>1</v></c></row><row r="9"><c><v>9</v></c></row>' },
        });
        const reader = await openXlsx(bytes);
        const rows = [];
        for await (const row of reader.sheets[0]?.rows() ?? []) rows.push(row.index);
        assert.deepEqual(rows, [1, 9]);
    });

    it('refuses a package that is missing a part it declares', async () => {
        const bytes = zipOf({
            '_rels/.rels':
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>',
        });
        await assert.rejects(readXlsx(bytes), /does not carry/);
    });

    it('refuses a sheet that points at a relationship the workbook does not have', async () => {
        const bytes = zipOf({
            '_rels/.rels':
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>',
            'xl/_rels/workbook.xml.rels':
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            'xl/workbook.xml':
                '<workbook xmlns:r="x"><sheets><sheet name="H" r:id="rIdPerdida"/></sheets></workbook>',
        });
        await assert.rejects(readXlsx(bytes), /rIdPerdida/);
    });
});
