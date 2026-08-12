import assert from 'node:assert/strict';
import { NO_FORMATS, readNumberFormats } from '../../src/core/read/numberFormats.js';
import { localDates } from '../../src/core/read/dates.js';
import type { RawReadValue, ReadRow } from '../../src/core/read/types.js';
import {
    cellValue,
    parseCellReference,
    readRows,
    styledCell,
    type CellContext,
    type RawCell,
} from '../../src/core/read/worksheet.js';
import { stylesOf } from '../helpers/package.js';
import { asAsyncIterable } from '../helpers/streams.js';

const PLAIN: CellContext = {
    sharedStrings: [],
    formats: NO_FORMATS,
    date1904: false,
    // A `Date` is what these tests read dates as: the assertions here are
    // about the cell, not about which of the five types it comes back in.
    dates: localDates,
};

function context(overrides: Partial<CellContext>): CellContext {
    return { ...PLAIN, ...overrides };
}

/** A cell as the file spells it, with everything it does not say left out. */
function raw(cell: Partial<RawCell>): RawCell {
    return { type: undefined, style: undefined, value: undefined, formula: undefined, ...cell };
}

/** The rows a `<sheetData>` holds, read whole. */
async function rowsOf<C>(
    body: string,
    convert: (cell: RawCell) => C,
    chunks = 1,
): Promise<ReadRow<C>[]> {
    const xml = `<worksheet><sheetData>${body}</sheetData></worksheet>`;
    const pieces: string[] = [];
    const size = Math.ceil(xml.length / chunks);
    for (let at = 0; at < xml.length; at += size) pieces.push(xml.slice(at, at + size));
    const rows: ReadRow<C>[] = [];
    for await (const row of readRows(asAsyncIterable(pieces), convert, 'sheet1.xml')) rows.push(row);
    return rows;
}

const values = (body: string, ctx = PLAIN, chunks?: number): Promise<ReadRow<RawReadValue>[]> =>
    rowsOf(body, (cell) => cellValue(cell, ctx), chunks);

describe('parseCellReference', () => {
    it('is the coordinates a cell is named by', () => {
        assert.deepEqual(parseCellReference('A1'), { column: 0, row: 1 });
        assert.deepEqual(parseCellReference('B12'), { column: 1, row: 12 });
        assert.deepEqual(parseCellReference('AA100'), { column: 26, row: 100 });
    });

    it('is nothing for what is not a reference', () => {
        for (const text of ['', '1A', 'A', '12', 'A1:B2', '$A$1']) {
            assert.equal(parseCellReference(text), undefined, `${text} was read as a reference`);
        }
    });
});

describe('cellValue', () => {
    it('reads a number, which is what a cell that says nothing holds', () => {
        assert.equal(cellValue(raw({ value: '42.5' }), PLAIN), 42.5);
        assert.equal(cellValue(raw({ type: 'n', value: '-3' }), PLAIN), -3);
    });

    it('reads a boolean out of the 1 or 0 it is written as', () => {
        assert.equal(cellValue(raw({ type: 'b', value: '1' }), PLAIN), true);
        assert.equal(cellValue(raw({ type: 'b', value: '0' }), PLAIN), false);
    });

    it('reads a shared string through the table it points into', () => {
        const ctx = context({ sharedStrings: ['uno', 'dos'] });
        assert.equal(cellValue(raw({ type: 's', value: '1' }), ctx), 'dos');
    });

    it('reads the string a cell carries itself, and the one a formula left', () => {
        assert.equal(cellValue(raw({ type: 'inlineStr', value: 'hola' }), PLAIN), 'hola');
        assert.equal(cellValue(raw({ type: 'str', value: 'hola' }), PLAIN), 'hola');
    });

    it('reads an error as what the cell shows, not as a value that went missing', () => {
        assert.equal(cellValue(raw({ type: 'e', value: '#DIV/0!' }), PLAIN), '#DIV/0!');
    });

    it('reads a date written out in full, which the spec allows and Excel does not write', () => {
        const value = cellValue(raw({ type: 'd', value: '2024-01-15T12:30:00' }), PLAIN);
        assert.ok(value instanceof Date);
        assert.deepEqual(
            [value.getFullYear(), value.getMonth(), value.getDate(), value.getHours()],
            [2024, 0, 15, 12],
        );
    });

    it('reads a date written out with a zone as the wall clock it shows', () => {
        // The zone is allowed and dropped: the rest of the sheet is serials,
        // which carry none, and a cell that moved by three hours where its
        // neighbours did not would be the odd one out in its own column.
        const withZone = cellValue(raw({ type: 'd', value: '2024-01-15T12:30:00Z' }), PLAIN);
        const without = cellValue(raw({ type: 'd', value: '2024-01-15T12:30:00' }), PLAIN);
        assert.deepEqual(withZone, without);
    });

    it('reads a date written out as a day alone, with no time to go with it', () => {
        const value = cellValue(raw({ type: 'd', value: '2024-01-15' }), PLAIN);
        assert.ok(value instanceof Date);
        assert.equal(value.getHours(), 0);
    });

    it('refuses a date cell holding something that is not one', () => {
        assert.throws(() => cellValue(raw({ type: 'd', value: 'ayer' }), PLAIN), /not a date/);
    });

    it('is a Date when the format under the number says it is one', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf([14, 0])) });
        const date = cellValue(raw({ value: '45306', style: 0 }), ctx);
        assert.ok(date instanceof Date);
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
        // The same number under a format that is not a date stays a number.
        assert.equal(cellValue(raw({ value: '45306', style: 1 }), ctx), 45306);
    });

    it('counts from 1904 when the workbook says it does', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf([14])), date1904: true });
        const date = cellValue(raw({ value: '43844', style: 0 }), ctx);
        assert.ok(date instanceof Date);
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
    });

    it('is null for a cell that is there and holds nothing', () => {
        assert.equal(cellValue(raw({ style: 3 }), PLAIN), null);
    });

    it('refuses a cell whose value is not what it says it is', () => {
        assert.throws(() => cellValue(raw({ value: 'no es un número' }), PLAIN), /not a number/);
        assert.throws(() => cellValue(raw({ type: 'd', value: 'ayer' }), PLAIN), /not a date/);
        assert.throws(() => cellValue(raw({ type: 'x', value: '1' }), PLAIN), /not a type/);
    });

    it('refuses a shared string the table does not have', () => {
        const ctx = context({ sharedStrings: ['uno'] });
        assert.throws(() => cellValue(raw({ type: 's', value: '5' }), ctx), /shared string 5/);
    });
});

describe('styledCell', () => {
    it('carries the value, and says no more than it has to', () => {
        assert.deepEqual(styledCell(raw({ value: '42' }), PLAIN), { v: 42 });
        assert.deepEqual(styledCell(raw({ type: 'b', value: '1' }), PLAIN), { v: true });
    });

    it('carries the number format as a style the writer takes back', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf(['#,##0.00', 14])) });
        assert.deepEqual(styledCell(raw({ value: '1234.5', style: 0 }), ctx), {
            v: 1234.5,
            s: { numFmt: '#,##0.00' },
        });
        const date = styledCell(raw({ value: '45306', style: 1 }), ctx);
        assert.deepEqual(date.s, { numFmt: 14 });
    });

    it('carries the formula, and the result the file cached for it', () => {
        assert.deepEqual(styledCell(raw({ value: '3', formula: '1+2' }), PLAIN), {
            v: 3,
            f: '1+2',
        });
    });

    it('says the type only where the writer would not work it out', () => {
        // Text that is the cached result of a formula has to stay text, and
        // an error is text that is not a value; a number, a boolean, a date
        // and a plain string all say what they are by being what they are.
        assert.equal(styledCell(raw({ type: 'str', value: 'x', formula: 'A1' }), PLAIN).t, 'str');
        assert.equal(styledCell(raw({ type: 'e', value: '#N/A' }), PLAIN).t, 'e');
        assert.equal(styledCell(raw({ type: 'inlineStr', value: 'x' }), PLAIN).t, undefined);
        assert.equal(styledCell(raw({ value: '1' }), PLAIN).t, undefined);
    });
});

describe('readRows', () => {
    it('reads the cells of a row into the columns they name', async () => {
        const rows = await values('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>');
        assert.equal(rows.length, 1);
        assert.equal(rows[0]?.index, 1);
        assert.deepEqual([...(rows[0]?.cells ?? [])], [1, undefined, 3]);
    });

    it('carries on from the last column for a cell that does not name one', async () => {
        const rows = await values('<row><c><v>1</v></c><c><v>2</v></c><c r="E1"><v>5</v></c><c><v>6</v></c></row>');
        assert.equal(rows[0]?.cells.length, 6);
        assert.deepEqual(rows[0]?.cells[4], 5);
        assert.deepEqual(rows[0]?.cells[5], 6);
    });

    it('numbers a row that does not number itself as the next one', async () => {
        const rows = await values('<row><c><v>1</v></c></row><row><c><v>2</v></c></row>');
        assert.deepEqual(rows.map((row) => row.index), [1, 2]);
    });

    it('keeps the numbers of rows the sheet skipped over', async () => {
        const rows = await values('<row r="1"><c><v>1</v></c></row><row r="9"><c><v>9</v></c></row>');
        assert.deepEqual(rows.map((row) => row.index), [1, 9]);
    });

    it('joins the runs of a string the cell carries itself', async () => {
        const rows = await values(
            '<row><c t="inlineStr"><is><r><t>en </t></r><r><t>partes</t></r></is></c></row>',
        );
        assert.equal(rows[0]?.cells[0], 'en partes');
    });

    it('reads the same rows however the file is cut into chunks', async () => {
        const body =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>una frase larga</t></is></c>' +
            '<c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>';
        const whole = await values(body);
        for (const chunks of [2, 5, 20, 100]) {
            assert.deepEqual(await values(body, PLAIN, chunks), whole, `cut into ${chunks}`);
        }
    });

    it('reads a file whose elements carry a namespace prefix', async () => {
        const xml =
            '<x:worksheet xmlns:x="urn:x"><x:sheetData><x:row r="1">' +
            '<x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData></x:worksheet>';
        const rows: ReadRow<RawReadValue>[] = [];
        for await (const row of readRows(
            asAsyncIterable([xml]),
            (cell) => cellValue(cell, PLAIN),
            'sheet1.xml',
        )) {
            rows.push(row);
        }
        assert.equal(rows[0]?.cells[0], 7);
    });

    it('leaves alone what is not the data', async () => {
        // `<dimension>` and `<mergeCells>` sit outside `<sheetData>` and have
        // an `r` of their own, which a reader that went by tag names alone
        // would take for a row.
        const xml =
            '<worksheet><dimension ref="A1:B2"/><sheetData><row r="1"><c><v>1</v></c></row>' +
            '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>';
        const rows: ReadRow<RawReadValue>[] = [];
        for await (const row of readRows(
            asAsyncIterable([xml]),
            (cell) => cellValue(cell, PLAIN),
            'sheet1.xml',
        )) {
            rows.push(row);
        }
        assert.equal(rows.length, 1);
    });

    it('reads a row that holds nothing as a row with no cells', async () => {
        const rows = await values('<row r="3"/>');
        assert.deepEqual(rows, [{ index: 3, cells: [] }]);
    });

    it('refuses a cell whose reference is not one', async () => {
        await assert.rejects(values('<row r="1"><c r="hola"><v>1</v></c></row>'), /not a cell reference/);
    });

    it('refuses a part that does not parse', async () => {
        await assert.rejects(values('<row r="1"><c><v>1</v></row>'));
    });
});
