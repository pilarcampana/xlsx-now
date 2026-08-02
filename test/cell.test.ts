import assert from 'node:assert/strict';
import { cellRef, cellXml, columnLetters, sanitizeText } from '../src/core/cell.js';
import { STYLE } from '../src/core/styles.js';

describe('sanitizeText', () => {
    it('escapes the five XML entities', () => {
        assert.equal(sanitizeText('&<>"\''), '&amp;&lt;&gt;&quot;&apos;');
    });

    it('escapes the ampersand first, so an entity is not built twice', () => {
        assert.equal(sanitizeText('&lt;'), '&amp;lt;');
    });

    it('leaves ordinary text alone', () => {
        assert.equal(sanitizeText('Muñoz, José — 100% ok'), 'Muñoz, José — 100% ok');
    });

    it('renders anything that is not a string as one', () => {
        assert.equal(sanitizeText(42), '42');
        assert.equal(sanitizeText(null), 'null');
        assert.equal(sanitizeText(undefined), 'undefined');
        assert.equal(sanitizeText({ toString: () => 'a & b' }), 'a &amp; b');
    });
});

describe('columnLetters', () => {
    it('numbers the first sheet of letters', () => {
        assert.equal(columnLetters(0), 'A');
        assert.equal(columnLetters(1), 'B');
        assert.equal(columnLetters(25), 'Z');
    });

    it('carries into two letters, which is not plain base 26', () => {
        assert.equal(columnLetters(26), 'AA');
        assert.equal(columnLetters(27), 'AB');
        assert.equal(columnLetters(51), 'AZ');
        assert.equal(columnLetters(52), 'BA');
        assert.equal(columnLetters(701), 'ZZ');
    });

    it('carries into three letters, up to Excel\'s last column', () => {
        assert.equal(columnLetters(702), 'AAA');
        // 16384 columns is the sheet's width; XFD is the last of them.
        assert.equal(columnLetters(16383), 'XFD');
    });
});

describe('cellRef', () => {
    it('joins the column letters to the row number', () => {
        assert.equal(cellRef(0, 1), 'A1');
        assert.equal(cellRef(26, 10), 'AA10');
        assert.equal(cellRef(16383, 1048576), 'XFD1048576');
    });
});

describe('cellXml', () => {
    it('writes nothing for an empty value with the default style', () => {
        assert.equal(cellXml(null, 'A1', STYLE.DEFAULT), '');
        assert.equal(cellXml(undefined, 'A1', STYLE.DEFAULT), '');
        assert.equal(cellXml('', 'A1', STYLE.DEFAULT), '');
    });

    it('writes an empty cell when it carries a style', () => {
        assert.equal(cellXml(null, 'A1', STYLE.HIGHLIGHT), '<c r="A1" s="2"/>');
        assert.equal(cellXml('', 'B2', STYLE.BOLD), '<c r="B2" s="1"/>');
    });

    it('writes numbers as numbers', () => {
        assert.equal(cellXml(0, 'A1', STYLE.DEFAULT), '<c r="A1" t="n"><v>0</v></c>');
        assert.equal(cellXml(-3.5, 'A1', STYLE.DEFAULT), '<c r="A1" t="n"><v>-3.5</v></c>');
    });

    it('falls back to text for a number XML cannot carry', () => {
        // NaN and the infinities have no numeric spelling in a sheet, so they
        // go in as what they read as.
        assert.match(cellXml(NaN, 'A1', STYLE.DEFAULT), /t="inlineStr"/);
        assert.match(cellXml(Infinity, 'A1', STYLE.DEFAULT), /<t xml:space="preserve">Infinity<\/t>/);
    });

    it('writes booleans as 1 and 0', () => {
        assert.equal(cellXml(true, 'A1', STYLE.DEFAULT), '<c r="A1" t="b"><v>1</v></c>');
        assert.equal(cellXml(false, 'A1', STYLE.DEFAULT), '<c r="A1" t="b"><v>0</v></c>');
    });

    it('writes a Date as an Excel serial number', () => {
        // 1970-01-01 is day 25569 of Excel's own epoch.
        assert.equal(
            cellXml(new Date(0), 'A1', STYLE.DEFAULT),
            '<c r="A1" t="n"><v>25569</v></c>',
        );
        assert.equal(
            cellXml(new Date('2024-01-15T12:30:00.000Z'), 'A1', STYLE.DEFAULT),
            '<c r="A1" t="n"><v>45306.52083333333</v></c>',
        );
    });

    it('writes strings inline, escaped, with the spaces preserved', () => {
        assert.equal(
            cellXml(' a & b ', 'A1', STYLE.DEFAULT),
            '<c r="A1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b </t></is></c>',
        );
    });

    it('carries the style index on every kind of cell', () => {
        assert.match(cellXml(1, 'A1', STYLE.BOLD), /^<c r="A1" t="n" s="1">/);
        assert.match(cellXml(true, 'A1', STYLE.HIGHLIGHT), /^<c r="A1" t="b" s="2">/);
        assert.match(cellXml('x', 'A1', STYLE.BOLD_HIGHLIGHT), /^<c r="A1" t="inlineStr" s="3">/);
        assert.match(cellXml(new Date(0), 'A1', STYLE.BOLD), /^<c r="A1" t="n" s="1">/);
    });

    it('leaves the style attribute out for the default style', () => {
        assert.doesNotMatch(cellXml(1, 'A1', STYLE.DEFAULT), / s=/);
    });
});
