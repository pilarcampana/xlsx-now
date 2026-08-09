import assert from 'node:assert/strict';
import { WidthMeter, cellTextLength, columnWidth } from '../src/core/autoWidth.js';
import type { NativeValue } from '../src/core/valueTypes.js';

/** One column, measured through everything that was written in it. */
function widthOf(max: number, values: readonly NativeValue[]): number | undefined {
    const meter = new WidthMeter(max);
    for (const value of values) meter.see(0, value);
    return meter.columnWidths()[0];
}

describe('cellTextLength', () => {
    it('counts the characters of a string', () => {
        assert.equal(cellTextLength('Ana & Co'), 8);
    });

    it('counts the digits a number is written as', () => {
        assert.equal(cellTextLength(10.5), 4);
        assert.equal(cellTextLength(-1234), 5);
    });

    it('takes the width the value\'s own type gave, over the one it is written as', () => {
        // A date is a five-digit serial in the file and ten characters on the
        // screen; only the type that made it a serial knows that.
        assert.equal(cellTextLength(45306, 'dd/mm/yyyy'.length), 10);
    });

    it('measures an empty cell as nothing, whatever width came with it', () => {
        assert.equal(cellTextLength(null, 10), 0);
    });

    it('measures a boolean as Excel spells it', () => {
        assert.equal(cellTextLength(true), 'TRUE'.length);
        assert.equal(cellTextLength(false), 'FALSE'.length);
    });

    it('gives an empty cell no width at all', () => {
        assert.equal(cellTextLength(null), 0);
        assert.equal(cellTextLength(undefined), 0);
        assert.equal(cellTextLength(''), 0);
    });
});

describe('cellTextLength: the line breaks inside a cell', () => {
    const TEXT = 'uno\ndos y dos\ntres';

    it('counts the whole text where the cell does not wrap', () => {
        // Excel shows a line break as one only where the text wraps: without
        // it the value is on one line, break and all, and that is the width.
        assert.equal(cellTextLength(TEXT), TEXT.length);
        assert.equal(cellTextLength(TEXT, undefined, false), TEXT.length);
    });

    it('counts the longest line where it does', () => {
        assert.equal(cellTextLength(TEXT, undefined, true), 'dos y dos'.length);
    });

    it('leaves the carriage return of a CRLF out of the line it ends', () => {
        // It is not a character the cell shows, so it is not one the column
        // has to fit.
        assert.equal(cellTextLength('uno\r\ndos y dos', undefined, true), 'dos y dos'.length);
    });

    it('measures a wrapping cell with no break in it as it always did', () => {
        assert.equal(cellTextLength('Ana & Co', undefined, true), 8);
    });

    it('counts an empty line as the nothing it is', () => {
        assert.equal(cellTextLength('\n\n', undefined, true), 0);
        assert.equal(cellTextLength('uno\n\n', undefined, true), 3);
    });

    it('still takes the width the value\'s own type gave over any of this', () => {
        // A type that says how wide its value shows has already answered; the
        // serial underneath has no line breaks to look for.
        assert.equal(cellTextLength(45306, 10, true), 10);
    });
});

describe('columnWidth', () => {
    it('is the characters plus the padding a column carries, in 1/256ths', () => {
        // ECMA-376 §18.3.1.13: the count is not the width. Eight characters
        // of Calibri 11 is the 8.7109375 Excel writes when it autofits one.
        assert.equal(columnWidth(8), 8.7109375);
        assert.equal(columnWidth(10), 10.7109375);
    });

    it('is wider than the text it was measured from, which is the point', () => {
        // A column exactly as wide as its longest value clips it, and shows a
        // date or a number under it as `####`.
        for (const characters of [1, 3, 12, 50]) {
            assert.ok(columnWidth(characters) > characters, `${characters}`);
        }
    });
});

describe('WidthMeter', () => {
    it('keeps the longest cell of the column', () => {
        assert.equal(widthOf(50, ['ab', 'abcdef', 'abcd']), columnWidth(6));
    });

    it('stops the column at the maximum it was opened with', () => {
        assert.equal(widthOf(4, ['ab', 'a very long line of text']), columnWidth(4));
    });

    it('leaves a column nobody wrote in without a width', () => {
        assert.equal(widthOf(20, [null, undefined, '']), undefined);
        assert.deepEqual([...new WidthMeter(20).columnWidths()], []);
    });

    it('measures each column on its own, however far apart they are', () => {
        const meter = new WidthMeter(20);
        meter.see(0, 'ab');
        meter.see(5, 'abcdef');
        const widths = meter.columnWidths();
        assert.equal(widths[0], columnWidth(2));
        assert.equal(widths[5], columnWidth(6));
        assert.equal(widths[1], undefined);
    });

    it('measures nothing when it was given no maximum', () => {
        // Every sheet has a meter; a sheet with no `autoWidthMax` gets one
        // that measures nothing, so nobody downstream has to ask whether it
        // has one at all.
        const meter = new WidthMeter(undefined);
        assert.equal(meter.measures, false);
        meter.see(0, 'a line of text');
        assert.deepEqual([...meter.columnWidths()], []);
    });

    it('says it measures when it was given one', () => {
        assert.equal(new WidthMeter(10).measures, true);
    });

    it('measures a value by the width its own type gave', () => {
        const meter = new WidthMeter(50);
        meter.see(0, 45306, 'dddd d "de" mmmm'.length);
        assert.equal(meter.columnWidths()[0], columnWidth('dddd d "de" mmmm'.length));
    });

    it('refuses a maximum no column can be sized by', () => {
        for (const max of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
            assert.throws(() => new WidthMeter(max), /autoWidthMax/);
        }
    });
});
