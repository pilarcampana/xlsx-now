import assert from 'node:assert/strict';
import { isDateFormat, readNumberFormats } from '../../src/core/read/numberFormats.js';
import { stylesOf } from '../helpers/package.js';

describe('isDateFormat', () => {
    it('is what tells a date format from a number one', () => {
        for (const code of ['dd/mm/yyyy', 'yyyy-mm-dd', 'h:mm:ss', 'mmm-yy', 'd', '[$-409]h:mm AM/PM']) {
            assert.ok(isDateFormat(code), `${code} was not read as a date`);
        }
        for (const code of ['General', '0.00', '#,##0', '0.00%', '0.00E+00', '@']) {
            assert.ok(!isDateFormat(code), `${code} was read as a date`);
        }
    });

    it('does not read the literal text of a format as part of it', () => {
        // The `d` of `días` is text the format shows, not a day it writes.
        assert.ok(!isDateFormat('#,##0 "días"'));
        assert.ok(!isDateFormat('0 \\d'));
        assert.ok(!isDateFormat('[Red]-0.00;[Blue]0.00'));
    });

    it('reads elapsed time, which is the one value written in brackets', () => {
        assert.ok(isDateFormat('[h]:mm:ss'));
        assert.ok(isDateFormat('[mm]:ss'));
    });

    it('still finds the date after a literal it skipped', () => {
        assert.ok(isDateFormat('"al " dd/mm/yyyy'));
    });
});

describe('readNumberFormats', () => {
    it('gives back the code a workbook declared, ready to be written again', () => {
        const formats = readNumberFormats(stylesOf(['#,##0.00', 'dd/mm/yyyy']));
        assert.equal(formats.forStyle(0), '#,##0.00');
        assert.equal(formats.forStyle(1), 'dd/mm/yyyy');
    });

    it('gives back the id of a built-in one, which has no code to give', () => {
        const formats = readNumberFormats(stylesOf([14, 22]));
        assert.equal(formats.forStyle(0), 14);
        assert.equal(formats.forStyle(1), 22);
    });

    it('says nothing at all for the general format', () => {
        const formats = readNumberFormats(stylesOf([0]));
        assert.equal(formats.forStyle(0), undefined);
        assert.equal(formats.forStyle(undefined), undefined);
    });

    it('knows the built-in dates, and that a percentage is not one', () => {
        const formats = readNumberFormats(stylesOf([14, 22, 46, 9, 0, 30]));
        assert.deepEqual(
            [0, 1, 2, 3, 4, 5].map((style) => formats.isDate(style)),
            [true, true, true, false, false, true],
        );
    });

    it('works a declared format out from its code', () => {
        const formats = readNumberFormats(stylesOf(['yyyy-mm-dd', '#,##0.00']));
        assert.equal(formats.isDate(0), true);
        assert.equal(formats.isDate(1), false);
    });

    it('reads the cell formats and not the named ones next to them', () => {
        // `cellStyleXfs` comes first in the part and is a list of its own;
        // reading the two as one would shift every style by its length. The
        // helper puts a percentage there, so a reader that ran them together
        // would answer that one for style 0.
        const formats = readNumberFormats(stylesOf(['dd/mm/yyyy']));
        assert.equal(formats.forStyle(0), 'dd/mm/yyyy');
    });

    it('says nothing for a style the workbook does not have', () => {
        const formats = readNumberFormats(stylesOf([14]));
        assert.equal(formats.forStyle(7), undefined);
        assert.equal(formats.isDate(7), false);
    });
});
