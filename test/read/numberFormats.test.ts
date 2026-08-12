import assert from 'node:assert/strict';
import {
    dateFormatKind,
    isDateFormat,
    readNumberFormats,
} from '../../src/core/read/numberFormats.js';
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
            [0, 1, 2, 3, 4, 5].map((style) => formats.dateKind(style)),
            ['date', 'dateTime', 'time', undefined, undefined, 'date'],
        );
    });

    it('works a declared format out from its code', () => {
        const formats = readNumberFormats(stylesOf(['yyyy-mm-dd', '#,##0.00', 'hh:mm:ss']));
        assert.equal(formats.dateKind(0), 'date');
        assert.equal(formats.dateKind(1), undefined);
        assert.equal(formats.dateKind(2), 'time');
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
        assert.equal(formats.dateKind(7), undefined);
    });
});

describe('dateFormatKind', () => {
    it('tells a date from a time from both, by what the code writes', () => {
        assert.equal(dateFormatKind('yyyy-mm-dd'), 'date');
        assert.equal(dateFormatKind('hh:mm:ss'), 'time');
        assert.equal(dateFormatKind('yyyy-mm-dd hh:mm:ss'), 'dateTime');
        assert.equal(dateFormatKind('#,##0.00'), undefined);
    });

    it('reads an `m` as minutes next to an hour or a second, and as a month elsewhere', () => {
        // The one letter in a format code that means two things, and Excel's
        // own rule for it.
        assert.equal(dateFormatKind('h:mm'), 'time');
        assert.equal(dateFormatKind('mm:ss'), 'time');
        assert.equal(dateFormatKind('mm/dd/yy'), 'date');
        assert.equal(dateFormatKind('m/d/yy h:mm'), 'dateTime');
    });

    it('never reads three of them as minutes, since a clock has no use for it', () => {
        assert.equal(dateFormatKind('mmm-yy'), 'date');
        assert.equal(dateFormatKind('d-mmmm'), 'date');
    });

    it('reads elapsed time as the time it is', () => {
        assert.equal(dateFormatKind('[h]:mm:ss'), 'time');
        assert.equal(dateFormatKind('[mm]:ss'), 'time');
    });

    it('says nothing for the letters a code only seems to have', () => {
        // Quoted text, escapes, padding, and everything else in brackets.
        assert.equal(dateFormatKind('#,##0 "días"'), undefined);
        assert.equal(dateFormatKind('[Red]0.00'), undefined);
        assert.equal(dateFormatKind('0.00_);[Blue](0.00)'), undefined);
        assert.equal(dateFormatKind('0" hs"'), undefined);
    });

    it('stops at the end of a code that never closes what it opened', () => {
        // Malformed, and the point is that it ends rather than reading past
        // the string: a file can hold anything, and this is parsing, not
        // validation.
        assert.equal(dateFormatKind('yyyy "sin cerrar'), 'date');
        assert.equal(dateFormatKind('[sin cerrar'), undefined);
        // What was opened is still read for what it says: an unclosed `[h` is
        // as much elapsed hours as a closed one.
        assert.equal(dateFormatKind('[h'), 'time');
    });

    it('reads a date under a locale prefix, which is where a bracket is skipped', () => {
        assert.equal(dateFormatKind('[$-409]d-mmm-yy'), 'date');
        assert.equal(dateFormatKind('[$-404]e/m/d'), 'date');
    });

    it('agrees with isDateFormat, which is the same question asked shorter', () => {
        for (const code of ['yyyy-mm-dd', 'hh:mm', '#,##0.00', '[Red]0.00', 'm/d/yy h:mm']) {
            assert.equal(isDateFormat(code), dateFormatKind(code) !== undefined, code);
        }
    });
});
