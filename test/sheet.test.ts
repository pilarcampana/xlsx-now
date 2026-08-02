import assert from 'node:assert/strict';
import {
    cellRowXml,
    SHEET_FOOTER,
    sheetHeaderXml,
    type ColumnFormats,
    type RowOptions,
} from '../src/core/sheet.js';
import { StyleTable } from '../src/core/styles.js';
import type { CellRow } from '../src/core/types.js';

/**
 * A row and the table its styles went into. Every test here starts from an
 * empty table, so an index says which style was asked for first, not which
 * one it is: `bold` is 1 in a table where bold was the first thing asked for.
 */
function rowXml(rowNumber: number, row: CellRow, options?: RowOptions): string {
    return cellRowXml(rowNumber, row, new StyleTable(), options);
}

/** Just the `<sheetViews>` of a worksheet header. */
function views(rows: number, columns: number): string {
    const xml = sheetHeaderXml({ rows, columns }, new StyleTable());
    return /<sheetViews>.*<\/sheetViews>/.exec(xml)?.[0] ?? '';
}

/** Just the `<cols>` of one, which is where a column's own layout goes. */
function cols(columnFormats: ColumnFormats, styles = new StyleTable()): string {
    const xml = sheetHeaderXml({ rows: 0, columns: 0 }, styles, columnFormats);
    return /<cols>.*<\/cols>/.exec(xml)?.[0] ?? '';
}

describe('sheetHeaderXml', () => {
    it('opens the worksheet and leaves sheetData open', () => {
        const xml = sheetHeaderXml({ rows: 0, columns: 0 }, new StyleTable());
        assert.ok(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'));
        assert.ok(xml.includes('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'));
        assert.ok(xml.endsWith('<sheetData>'));
    });

    it('writes the view even with nothing frozen', () => {
        // A reader that goes looking for the view has to find one.
        assert.equal(views(0, 0), '<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
    });

    it('splits on rows alone below the frozen rows', () => {
        const xml = views(1, 0);
        assert.ok(xml.includes('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'));
        assert.doesNotMatch(xml, /xSplit/);
        assert.ok(xml.includes('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'));
    });

    it('splits on columns alone right of the frozen columns', () => {
        const xml = views(0, 2);
        assert.ok(xml.includes('<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>'));
        assert.doesNotMatch(xml, /ySplit/);
    });

    it('splits on both, and the scrolling area starts past the two', () => {
        const xml = views(2, 3);
        assert.ok(xml.includes('<pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>'));
        assert.ok(xml.includes('<selection pane="bottomRight" activeCell="D3" sqref="D3"/>'));
    });

    it('names the columns past 26 in the freeze too', () => {
        assert.ok(views(1, 26).includes('topLeftCell="AA2"'));
    });
});

describe('SHEET_FOOTER', () => {
    it('closes what the header opened', () => {
        assert.equal(SHEET_FOOTER, '</sheetData></worksheet>');
    });
});

describe('cellRowXml', () => {
    it('numbers the row and puts each value in its own column', () => {
        assert.equal(
            rowXml(3, [1, true]),
            '<row r="3"><c r="A3" t="n"><v>1</v></c><c r="B3" t="b"><v>1</v></c></row>',
        );
    });

    it('writes an empty row as an empty row', () => {
        assert.equal(rowXml(1, []), '<row r="1"></row>');
    });

    it('skips an undefined position without shifting the ones after it', () => {
        // The position *is* the column, so a hole has to stay a hole.
        assert.equal(
            rowXml(1, ['a', undefined, 'b']),
            '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">a</t></is></c>' +
                '<c r="C1" t="inlineStr"><is><t xml:space="preserve">b</t></is></c></row>',
        );
    });

    it('writes no cell for an unstyled empty value', () => {
        assert.equal(rowXml(1, [null, '']), '<row r="1"></row>');
    });

    it('writes a styled empty cell: asking for the style is asking for the cell', () => {
        assert.equal(
            rowXml(1, [{ v: null, s: { bold: true } }, { s: { italic: true } }]),
            '<row r="1"><c r="A1" s="1"/><c r="B1" s="2"/></row>',
        );
    });

    it('takes a cell as a value plus what it says about itself', () => {
        assert.equal(
            rowXml(2, [{ v: 7, s: { bold: true } }]),
            '<row r="2"><c r="A2" t="n" s="1"><v>7</v></c></row>',
        );
    });

    it('takes a cell with no style as the default style', () => {
        assert.equal(rowXml(2, [{ v: 7 }]), '<row r="2"><c r="A2" t="n"><v>7</v></c></row>');
    });

    it('writes the formula and the type a cell carries', () => {
        assert.equal(
            rowXml(1, [{ v: 3, f: '=A1+A2' }, { v: '007', t: 'inlineStr' }]),
            '<row r="1"><c r="A1"><f>A1+A2</f><v>3</v></c>' +
                '<c r="B1" t="inlineStr"><is><t xml:space="preserve">007</t></is></c></row>',
        );
    });

    it('takes a Date as a value, not as a cell that says more', () => {
        // It is the one object a cell can be on its own — and the one value
        // that gets a style nobody asked for, so it is not shown as a serial.
        assert.equal(
            rowXml(1, [new Date(1970, 0, 1)]),
            '<row r="1"><c r="A1" t="n" s="1"><v>25569</v></c></row>',
        );
    });

    it('says what an object that is no kind of cell was meant to be', () => {
        // Going in as a blank would hide it until someone opened the file.
        assert.throws(() => rowXml(1, [{ value: 7 } as never]), /"value"/);
        assert.throws(() => rowXml(1, [{} as never]), /this one is empty/);
    });
});

describe('cellRowXml: the column a cell asks for', () => {
    it('sends the cell to the column it names, by letter or by number', () => {
        assert.equal(
            rowXml(1, [{ v: 'far', col: 'D' }]),
            '<row r="1"><c r="D1" t="inlineStr"><is><t xml:space="preserve">far</t></is></c></row>',
        );
        // Columns are numbered from 1, as the sheet shows them.
        assert.equal(rowXml(1, [{ v: 1, col: 4 }]), '<row r="1"><c r="D1" t="n"><v>1</v></c></row>');
    });

    it('costs nothing for the columns it skips over', () => {
        const row = rowXml(1, [{ v: 'a', col: 'A' }, { v: 'z', col: 'BZ' }]);
        assert.ok(row.includes('r="A1"'), row);
        assert.ok(row.includes('r="BZ1"'), row);
        assert.equal((row.match(/<c /g) ?? []).length, 2);
    });

    it('carries on from there, so the cells after it need say nothing', () => {
        const row = rowXml(1, [{ v: 1, col: 'C' }, 2, 3]);
        assert.ok(row.includes('r="C1"') && row.includes('r="D1"') && row.includes('r="E1"'), row);
    });

    it('refuses a column the line has already gone past', () => {
        // Two cells in one column is a file Excel opens as one of them, and
        // which one is nobody's decision to leave to it.
        assert.throws(() => rowXml(5, ['a', 'b', { v: 'c', col: 'A' }]), /row 5 has already written/);
        assert.throws(() => rowXml(1, [{ v: 1, col: 'C' }, { v: 2, col: 'C' }]), /already written/);
    });

    it('says what is not a column at all', () => {
        assert.throws(() => rowXml(1, [{ v: 1, col: 'A1' }]), /is not a column/);
        assert.throws(() => rowXml(1, [{ v: 1, col: '' }]), /is not a column/);
        assert.throws(() => rowXml(1, [{ v: 1, col: 0 }]), /is not a column/);
        assert.throws(() => rowXml(1, [{ v: 1, col: 1.5 }]), /is not a column/);
    });
});

describe('cellRowXml: what the row itself asks for', () => {
    it('adds nothing when it asks for nothing', () => {
        assert.equal(rowXml(1, [], {}), '<row r="1"></row>');
        assert.equal(rowXml(1, [], { hidden: false }), '<row r="1"></row>');
    });

    it('marks a height as custom, which is what makes Excel apply it', () => {
        assert.equal(rowXml(1, [], { height: 22 }), '<row r="1" ht="22" customHeight="1"></row>');
    });

    it('hides the row', () => {
        assert.equal(rowXml(1, [], { hidden: true }), '<row r="1" hidden="1"></row>');
    });

    it('styles the whole row, under whatever its cells carry', () => {
        assert.equal(
            rowXml(1, [{ v: 'x', s: { italic: true } }], { s: { bold: true } }),
            '<row r="1" s="2" customFormat="1">' +
                '<c r="A1" t="inlineStr" s="1"><is><t xml:space="preserve">x</t></is></c></row>',
        );
    });

    it('takes all three at once', () => {
        assert.equal(
            rowXml(4, [], { height: 8, hidden: true, s: { bold: true } }),
            '<row r="4" s="1" customFormat="1" ht="8" customHeight="1" hidden="1"></row>',
        );
    });
});

describe('sheetHeaderXml: the columns of the sheet', () => {
    it('writes no <cols> when the sheet asks for none', () => {
        // An empty `<cols/>` is a sheet Excel refuses to open.
        assert.ok(!sheetHeaderXml({ rows: 0, columns: 0 }, new StyleTable()).includes('cols'));
        assert.equal(cols([]), '');
        assert.equal(cols({}), '');
    });

    it('comes before the rows, which is where a worksheet carries it', () => {
        const xml = sheetHeaderXml({ rows: 1, columns: 0 }, new StyleTable(), [{ width: 8 }]);
        assert.ok(xml.indexOf('<cols>') > xml.indexOf('</sheetViews>'), xml);
        assert.ok(xml.indexOf('<cols>') < xml.indexOf('<sheetData>'), xml);
    });

    it('marks a width as custom, which is what makes Excel apply it', () => {
        assert.equal(cols([{ width: 8 }]), '<cols><col min="1" max="1" width="8" customWidth="1"/></cols>');
    });

    it('spans one column per entry, at the position it was declared in', () => {
        assert.equal(
            cols([{ width: 8 }, undefined, { hidden: true }]),
            '<cols><col min="1" max="1" width="8" customWidth="1"/><col min="3" max="3" hidden="1"/></cols>',
        );
    });

    it('takes the same thing keyed by the column each one is for', () => {
        assert.equal(cols({ A: { width: 8 }, C: { hidden: true } }), cols([{ width: 8 }, undefined, { hidden: true }]));
    });

    it('writes them left to right, whatever order the keys came in', () => {
        assert.equal(cols({ C: { width: 3 }, A: { width: 1 } }), cols([{ width: 1 }, undefined, { width: 3 }]));
    });

    it('registers the column style, so its cells fall back to it', () => {
        const styles = new StyleTable({ money: { numFmt: '#,##0.00' } });
        assert.equal(cols({ B: { s: 'money' } }, styles), '<cols><col min="2" max="2" style="1"/></cols>');
    });

    it('reads a key that is a number as the column of that number', () => {
        assert.equal(cols({ '3': { width: 3 } }), cols({ C: { width: 3 } }));
    });

    it('says what is not a column', () => {
        assert.throws(() => cols({ 'A1': { width: 3 } }), /is not a column/);
    });
});

describe('sheetHeaderXml: the widths a sheet measured for itself', () => {
    /** Just the `<cols>`, out of what the formats say and what the cells measured. */
    function measured(
        autoWidths: readonly (number | undefined)[],
        columnFormats?: ColumnFormats,
    ): string {
        const xml = sheetHeaderXml({ rows: 0, columns: 0 }, new StyleTable(), columnFormats, autoWidths);
        return /<cols>.*<\/cols>/.exec(xml)?.[0] ?? '';
    }

    it('writes a width for every column that measured one', () => {
        assert.equal(measured([4, 12]), cols([{ width: 4 }, { width: 12 }]));
    });

    it('leaves out the columns nothing was written in', () => {
        assert.equal(measured([undefined, 12]), cols([undefined, { width: 12 }]));
        assert.equal(measured([]), '');
    });

    it('gives way to a width the sheet was given outright', () => {
        assert.equal(measured([4], [{ width: 30 }]), cols([{ width: 30 }]));
    });

    it('fills in the width of a column whose format says everything else', () => {
        assert.equal(measured([4], { A: { hidden: true } }), cols([{ width: 4, hidden: true }]));
    });

    it('writes both sources left to right, as one <cols>', () => {
        assert.equal(
            measured([undefined, undefined, 3], { A: { width: 8 } }),
            cols([{ width: 8 }, undefined, { width: 3 }]),
        );
    });
});
