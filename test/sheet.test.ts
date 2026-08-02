import assert from 'node:assert/strict';
import { cellRowXml, SHEET_FOOTER, sheetHeaderXml } from '../src/core/sheet.js';

/** Just the `<sheetViews>` of a worksheet header. */
function views(rows: number, columns: number): string {
    return /<sheetViews>.*<\/sheetViews>/.exec(sheetHeaderXml({ rows, columns }))?.[0] ?? '';
}

describe('sheetHeaderXml', () => {
    it('opens the worksheet and leaves sheetData open', () => {
        const xml = sheetHeaderXml({ rows: 0, columns: 0 });
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
            cellRowXml(3, [1, true]),
            '<row r="3"><c r="A3" t="n"><v>1</v></c><c r="B3" t="b"><v>1</v></c></row>',
        );
    });

    it('writes an empty row as an empty row', () => {
        assert.equal(cellRowXml(1, []), '<row r="1"></row>');
    });

    it('skips an undefined position without shifting the ones after it', () => {
        // The position *is* the column, so a hole has to stay a hole.
        assert.equal(
            cellRowXml(1, ['a', undefined, 'b']),
            '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">a</t></is></c>' +
                '<c r="C1" t="inlineStr"><is><t xml:space="preserve">b</t></is></c></row>',
        );
    });

    it('writes no cell for an unstyled empty value', () => {
        assert.equal(cellRowXml(1, [null, '']), '<row r="1"></row>');
    });

    it('writes a styled empty cell, however it was asked for', () => {
        assert.equal(
            cellRowXml(1, [{ value: null, style: { highlight: true } }, { value: undefined, style: { bold: true } }]),
            '<row r="1"><c r="A1" s="2"/><c r="B1" s="1"/></row>',
        );
    });

    it('takes a wrapper as a value plus a style', () => {
        assert.equal(
            cellRowXml(2, [{ value: 7, style: { bold: true, highlight: true } }]),
            '<row r="2"><c r="A2" t="n" s="3"><v>7</v></c></row>',
        );
    });

    it('takes a wrapper with no style as the default style', () => {
        assert.equal(cellRowXml(2, [{ value: 7 }]), '<row r="2"><c r="A2" t="n"><v>7</v></c></row>');
    });

    it('takes a Date as a value, not as a wrapper', () => {
        // It is the one object a cell can be without asking for a style.
        assert.equal(
            cellRowXml(1, [new Date(0)]),
            '<row r="1"><c r="A1" t="n"><v>25569</v></c></row>',
        );
    });
});
