import assert from 'node:assert/strict';
import {
    cellRef,
    cellXml,
    columnLetters,
    sanitizeText,
} from '../src/core/cell.js';

/** The two entries a cell can ask for here; the real ones come from the table. */
const DEFAULT = 0;
const STYLED = 2;

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

    it('drops the characters XML 1.0 has no place for', () => {
        // One of these is enough to make the whole file unreadable, and there
        // is no escape for them: XML leaves them out of `Char` altogether, so
        // `&#0;` is refused as flatly as the character itself.
        assert.equal(sanitizeText('ab\u0000cd\u0007ef'), 'abcdef');
        assert.equal(sanitizeText('\u0008\u000B\u000C\u001F\uFFFE\uFFFF'), '');
    });

    it('keeps the three control characters XML does allow', () => {
        assert.equal(sanitizeText('a\tb\nc\rd'), 'a\tb\nc\rd');
    });

    it('keeps a surrogate pair whole, emoji and all', () => {
        // The pair is two code units and one character: dropping either half
        // would break far more than it fixed.
        assert.equal(sanitizeText('ok \u{1F600}'), 'ok \u{1F600}');
        assert.equal(sanitizeText('\u{10FFFF}'), '\u{10FFFF}');
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
        assert.equal(cellXml(null, 'A1', DEFAULT), '');
        assert.equal(cellXml(undefined, 'A1', DEFAULT), '');
        assert.equal(cellXml('', 'A1', DEFAULT), '');
    });

    it('writes an empty cell when it carries a style', () => {
        assert.equal(cellXml(null, 'A1', STYLED), '<c r="A1" s="2"/>');
        assert.equal(cellXml('', 'B2', 1), '<c r="B2" s="1"/>');
    });

    it('writes numbers as numbers, with no type to say so', () => {
        // `n` is what a `<c>` holds without a `t`, and numbers are most of
        // what a sheet is made of: the attribute is the one worth leaving out.
        assert.equal(cellXml(0, 'A1', DEFAULT), '<c r="A1"><v>0</v></c>');
        assert.equal(cellXml(-3.5, 'A1', DEFAULT), '<c r="A1"><v>-3.5</v></c>');
    });

    it('falls back to text for a number XML cannot carry', () => {
        // NaN and the infinities have no numeric spelling in a sheet, so they
        // go in as what they read as.
        assert.match(cellXml(NaN, 'A1', DEFAULT), /t="inlineStr"/);
        assert.match(cellXml(Infinity, 'A1', DEFAULT), /<t>Infinity<\/t>/);
    });

    it('writes booleans as 1 and 0', () => {
        assert.equal(cellXml(true, 'A1', DEFAULT), '<c r="A1" t="b"><v>1</v></c>');
        assert.equal(cellXml(false, 'A1', DEFAULT), '<c r="A1" t="b"><v>0</v></c>');
    });

    it('writes strings inline, escaped', () => {
        assert.equal(
            cellXml('a & b', 'A1', DEFAULT),
            '<c r="A1" t="inlineStr"><is><t>a &amp; b</t></is></c>',
        );
    });

    it('asks for the spaces to be kept only when there are any to lose', () => {
        // Excel trims the edges of a `<t>` that does not say otherwise, and
        // nothing else about the text is at stake — so the attribute is
        // written where it changes what comes back, and nowhere else.
        for (const text of [' 007', '007 ', '\ta', 'b\n']) {
            assert.match(cellXml(text, 'A1', DEFAULT), /<t xml:space="preserve">/, text);
        }
        for (const text of ['007', 'a b', 'Ana & Co']) {
            assert.doesNotMatch(cellXml(text, 'A1', DEFAULT), /xml:space/, text);
        }
    });

    it('carries the style index on every kind of cell', () => {
        assert.match(cellXml(1, 'A1', 1), /^<c r="A1" s="1">/);
        assert.match(cellXml(true, 'A1', STYLED), /^<c r="A1" t="b" s="2">/);
        assert.match(cellXml('x', 'A1', 3), /^<c r="A1" t="inlineStr" s="3">/);
    });

    it('leaves the style attribute out for the default style', () => {
        assert.doesNotMatch(cellXml(1, 'A1', DEFAULT), / s=/);
    });
});


describe('cellXml: a formula', () => {
    it('writes the expression, and the value next to it as the cached result', () => {
        assert.equal(
            cellXml(45, 'C3', DEFAULT, 'SUM(A1:A9)'),
            '<c r="C3"><f>SUM(A1:A9)</f><v>45</v></c>',
        );
    });

    it('leaves the cell for the reader to work out when no value came with it', () => {
        assert.equal(cellXml(undefined, 'C3', DEFAULT, 'NOW()'), '<c r="C3"><f>NOW()</f></c>');
    });

    it('takes the expression with or without the = a sheet shows', () => {
        assert.equal(cellXml(null, 'A1', DEFAULT, '=B1+1'), cellXml(null, 'A1', DEFAULT, 'B1+1'));
    });

    it('escapes what would otherwise close the element', () => {
        assert.match(cellXml(null, 'A1', DEFAULT, 'IF(A1<2,"a & b","")'), /IF\(A1&lt;2,&quot;a &amp; b/);
    });

    it('carries the style and the type it is given', () => {
        assert.equal(
            cellXml('ok', 'A1', 3, 'B1', 'str'),
            '<c r="A1" t="str" s="3"><f>B1</f><v>ok</v></c>',
        );
    });

    it('types a cached result that is text, which a `<v>` is read as a number without', () => {
        assert.equal(cellXml('ok', 'A1', DEFAULT, 'B1'), '<c r="A1" t="str"><f>B1</f><v>ok</v></c>');
        assert.equal(cellXml(true, 'A1', DEFAULT, 'B1'), '<c r="A1" t="b"><f>B1</f><v>1</v></c>');
    });

    it('writes text next to a formula as `str`, never as an inline string', () => {
        // `inlineStr` puts the value in an `<is>`, and there is no `<is>` next
        // to an `<f>`: a formula's cached result is the `<v>` and nothing else.
        assert.equal(
            cellXml('ok', 'A1', DEFAULT, 'B1', 'inlineStr'),
            '<c r="A1" t="str"><f>B1</f><v>ok</v></c>',
        );
    });
});

describe('cellXml: the type said outright', () => {
    it('keeps a number that is really a code from being shown as one', () => {
        assert.equal(
            cellXml('007', 'A1', DEFAULT, undefined, 'inlineStr'),
            '<c r="A1" t="inlineStr"><is><t>007</t></is></c>',
        );
    });

    it('writes text asked for as `str` as the inline string it means', () => {
        // `str` is the cached result of a formula, and there is no formula
        // here: what the caller asked for is text, and text in a cell of its
        // own is an inline string.
        assert.equal(
            cellXml('007', 'A1', DEFAULT, undefined, 'str'),
            cellXml('007', 'A1', DEFAULT, undefined, 'inlineStr'),
        );
    });

    it('writes a string as the number the caller says it is', () => {
        assert.equal(cellXml('1.5', 'A1', DEFAULT, undefined, 'n'), '<c r="A1"><v>1.5</v></c>');
    });

    it('writes an error as the code a sheet shows', () => {
        assert.equal(
            cellXml('#N/A', 'A1', DEFAULT, undefined, 'e'),
            '<c r="A1" t="e"><v>#N/A</v></c>',
        );
    });

    it('is still an empty cell when there is nothing to type', () => {
        assert.equal(cellXml(null, 'A1', DEFAULT, undefined, 'n'), '');
        assert.equal(cellXml(null, 'A1', STYLED, undefined, 'n'), '<c r="A1" s="2"/>');
    });
});
