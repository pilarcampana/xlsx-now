import assert from 'node:assert/strict';
import { WidthMeter, cellTextLength } from '../src/core/autoWidth.js';
import { DATETIME_FORMAT, DATE_FORMAT } from '../src/core/styles.js';
import type { CellValue } from '../src/core/types.js';

/** One column, measured through everything that was written in it. */
function widthOf(max: number, values: readonly CellValue[]): number | undefined {
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

    it('measures a date by the format it will be shown in', () => {
        assert.equal(cellTextLength(new Date(2024, 0, 15)), DATE_FORMAT.length);
        assert.equal(cellTextLength(new Date(2024, 0, 15, 12, 30)), DATETIME_FORMAT.length);
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

describe('WidthMeter', () => {
    it('keeps the longest cell of the column', () => {
        assert.equal(widthOf(50, ['ab', 'abcdef', 'abcd']), 6);
    });

    it('stops the column at the maximum it was opened with', () => {
        assert.equal(widthOf(4, ['ab', 'a very long line of text']), 4);
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
        assert.equal(widths[0], 2);
        assert.equal(widths[5], 6);
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

    it('refuses a maximum no column can be sized by', () => {
        for (const max of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
            assert.throws(() => new WidthMeter(max), /autoWidthMax/);
        }
    });
});
