import 'temporal-polyfill/global';
import assert from 'node:assert/strict';
import { Temporal } from 'temporal-polyfill';
import { columnWidth, WidthMeter } from '../src/core/autoWidth.js';
import { cellRowXml } from '../src/core/sheet.js';
import {
    DateFormats,
    DEFAULT_DATETIME_FORMAT,
    DEFAULT_DATE_FORMAT,
    DEFAULT_TIME_FORMAT,
    StyleTable,
} from '../src/core/styles.js';
import type { CellRow } from '../src/core/types.js';
import {
    bigintValue,
    dateValue,
    defaultTypes,
    plainDateTimeValue,
    plainDateValue,
    plainTimeValue,
    serialValue,
    shownWidth,
    urlValue,
    ValueTypes,
    withType,
    type ConvertedValue,
    type ConvertContext,
    type TypeMap,
} from '../src/core/valueTypes.js';

/** The context a workbook that said nothing about its dates hands out. */
const PLAIN: ConvertContext = { dates: new DateFormats(), clock: 'local' };
/** The same workbook, reading its `Date`s by the UTC clock instead. */
const UTC: ConvertContext = { dates: new DateFormats(), clock: 'utc' };

describe('dateValue', () => {
    it('writes a Date as an Excel serial number', () => {
        // 1970-01-01 is day 25569 of Excel's own epoch.
        assert.equal(dateValue(new Date(1970, 0, 1), PLAIN).v, 25569);
        assert.equal(dateValue(new Date(2024, 0, 15, 12, 0), PLAIN).v, 45306.5);
    });

    it('shows a date under the built-in short date, and adds the time only when there is one', () => {
        assert.equal(dateValue(new Date(2024, 0, 15), PLAIN).numFmt, DEFAULT_DATE_FORMAT);
        assert.equal(
            dateValue(new Date(2024, 0, 15, 12, 30), PLAIN).numFmt,
            DEFAULT_DATETIME_FORMAT,
        );
    });

    it('takes the formats the workbook it belongs to uses', () => {
        const context: ConvertContext = {
            dates: new DateFormats({ dateFormat: 'yyyy-mm-dd' }),
            clock: 'local',
        };
        assert.equal(dateValue(new Date(2024, 0, 15), context).numFmt, 'yyyy-mm-dd');
        assert.equal(
            dateValue(new Date(2024, 0, 15, 12, 30), context).numFmt,
            'yyyy-mm-dd hh:mm:ss',
        );
    });

    it('measures a built-in date as the widest one a locale writes', () => {
        // The reader spells the short date, not the file, so what is measured
        // is the longest it comes to anywhere.
        assert.equal(dateValue(new Date(2024, 0, 15), PLAIN).width, 'dd/mm/yyyy'.length);
        assert.equal(
            dateValue(new Date(2024, 0, 15, 12, 30), PLAIN).width,
            'dd/mm/yyyy hh:mm:ss'.length,
        );
    });

    it('measures a date the workbook spelled out by its own format code', () => {
        const context: ConvertContext = {
            dates: new DateFormats({ dateFormat: 'dd/mm/yy' }),
            clock: 'local',
        };
        assert.equal(dateValue(new Date(2024, 0, 15), context).width, 'dd/mm/yy'.length);
    });

    it('reads a Date by the UTC clock when the workbook asked for that one', () => {
        const noon = new Date(Date.UTC(2024, 0, 15, 12, 0));
        assert.equal(dateValue(noon, UTC).v, 45306.5);
        // The same instant read locally is the same serial moved by the zone,
        // which is the whole of what the option decides.
        assert.equal(
            dateValue(noon, PLAIN).v,
            45306.5 - (noon.getTimezoneOffset() * 60000) / 86400000,
        );
    });

});

describe('serialValue', () => {
    it('shows a whole day as a date and a day with an hour in it as both', () => {
        assert.equal(serialValue(45306, PLAIN).numFmt, DEFAULT_DATE_FORMAT);
        assert.equal(serialValue(45306.5, PLAIN).numFmt, DEFAULT_DATETIME_FORMAT);
    });

    it('shows a serial with no day left in it as a time of day', () => {
        // 31/12/1899 is the day serial 0 lands on: what is left is the time.
        assert.equal(serialValue(0.4375, PLAIN).numFmt, DEFAULT_TIME_FORMAT);
        assert.equal(serialValue(0.4375, PLAIN).width, 'hh:mm:ss'.length);
    });

    it('takes the kind from whoever knows it, rather than reading it off', () => {
        assert.equal(serialValue(45306, PLAIN, 'dateTime').numFmt, DEFAULT_DATETIME_FORMAT);
    });
});

describe('the Temporal values', () => {
    it('writes a PlainDate as the day it is, and shows it as a date', () => {
        const value = plainDateValue(Temporal.PlainDate.from('2024-01-15'), PLAIN);
        assert.equal(value.v, 45306);
        assert.equal(value.numFmt, DEFAULT_DATE_FORMAT);
    });

    it('writes a PlainDateTime as the day and the time in it', () => {
        const value = plainDateTimeValue(Temporal.PlainDateTime.from('2024-01-15T12:00'), PLAIN);
        assert.equal(value.v, 45306.5);
        assert.equal(value.numFmt, DEFAULT_DATETIME_FORMAT);
    });

    it('writes a PlainTime as the fraction of a day a sheet stores a time as', () => {
        const value = plainTimeValue(Temporal.PlainTime.from('10:30'), PLAIN);
        assert.equal(value.v, 0.4375);
        assert.equal(value.numFmt, DEFAULT_TIME_FORMAT);
    });

    it('does not move by a time zone, whatever clock the workbook reads Dates by', () => {
        const day = Temporal.PlainDate.from('2024-01-15');
        assert.equal(plainDateValue(day, PLAIN).v, plainDateValue(day, UTC).v);
    });

    it('is in the default types, since the environment has a Temporal', () => {
        const types = defaultTypes();
        assert.equal(types.get(Temporal.PlainDate)?.convert, plainDateValue);
        assert.equal(types.get(Temporal.PlainDateTime)?.convert, plainDateTimeValue);
        assert.equal(types.get(Temporal.PlainTime)?.convert, plainTimeValue);
    });

    it('is looked up by the class, the way every other type is', () => {
        const value = new ValueTypes(defaultTypes(), PLAIN).convert(
            Temporal.PlainDate.from('2024-01-15'),
        );
        assert.deepEqual(value, plainDateValue(Temporal.PlainDate.from('2024-01-15'), PLAIN));
    });
});

describe('bigintValue', () => {
    it('keeps a whole number a cell can hold as a number', () => {
        assert.deepEqual(bigintValue(0n), { v: 0 });
        assert.deepEqual(bigintValue(-42n), { v: -42 });
        assert.deepEqual(bigintValue(BigInt(Number.MAX_SAFE_INTEGER)), {
            v: Number.MAX_SAFE_INTEGER,
        });
    });

    it('writes one a cell cannot hold as text, in both directions', () => {
        // A cell holds a double: past 2^53 the number that comes back out of
        // the file is not the one that went in, and the digits are the whole
        // reason to have written a BigInt in the first place.
        const past = BigInt(Number.MAX_SAFE_INTEGER) + 2n;
        assert.deepEqual(bigintValue(past), { v: '9007199254740993', t: 'inlineStr' });
        assert.deepEqual(bigintValue(-past), { v: '-9007199254740993', t: 'inlineStr' });
    });
});

describe('urlValue', () => {
    it('writes a URL as its text', () => {
        assert.deepEqual(urlValue(new URL('https://example.com/a?b=1')), {
            v: 'https://example.com/a?b=1',
        });
    });
});

describe('shownWidth', () => {
    it('says nothing about a value that needed no conversion', () => {
        assert.equal(shownWidth(undefined), undefined);
    });

    it('takes the width the conversion gave, over anything it could work out', () => {
        assert.equal(shownWidth({ v: 1, numFmt: 'yyyy-mm-dd hh:mm:ss', width: 3 }), 3);
    });

    it('works the width out of a format code, which is a template of the output', () => {
        assert.equal(shownWidth({ v: 1, numFmt: '[h]:mm:ss' }), '[h]:mm:ss'.length);
    });

    it('has nothing to work out from a built-in format, which is an id', () => {
        assert.equal(shownWidth({ v: 1, numFmt: DEFAULT_DATE_FORMAT }), undefined);
    });

    it('has nothing to work out from a value that brought no format', () => {
        assert.equal(shownWidth({ v: 'x' }), undefined);
    });
});

describe('withType', () => {
    class Money {
        constructor(readonly cents: number) {}
    }
    const handler = { convert: (m: Money): ConvertedValue => ({ v: m.cents / 100 }) };

    it('adds the type to a map that has everything the one it is based on had', () => {
        const base = defaultTypes();
        const types = withType(base, Money, handler);
        assert.equal(types.get(Money), handler);
        assert.equal(types.get(Date), base.get(Date));
    });

    it('never touches what it was based on', () => {
        const before = defaultTypes().size;
        withType(defaultTypes(), Money, handler);
        assert.equal(defaultTypes().size, before);
        assert.equal(defaultTypes().get(Money), undefined);
    });

    it('composes, so a map can be built one type at a time', () => {
        class Weight {}
        const types = withType(withType(defaultTypes(), Money, handler), Weight, {
            convert: () => ({ v: 'heavy' }),
        });
        assert.equal(types.get(Money), handler);
        assert.ok(types.get(Weight));
    });
});

describe('ValueTypes: what claims a value', () => {
    class Money {
        constructor(readonly cents: number) {}
    }
    class Cents extends Money {}

    function typesFor(map: TypeMap): ValueTypes {
        return new ValueTypes(map, PLAIN);
    }

    it('leaves a value the writer already writes alone', () => {
        const types = typesFor(defaultTypes());
        for (const value of ['x', 1, true, null, undefined]) {
            assert.equal(types.convert(value), undefined, String(value));
        }
    });

    it('finds a registered class', () => {
        const types = typesFor(
            withType(defaultTypes(), Money, { convert: (m) => ({ v: m.cents / 100 }) }),
        );
        assert.deepEqual(types.convert(new Money(1250)), { v: 12.5 });
    });

    it('writes a subclass as whatever its base was registered as', () => {
        const types = typesFor(
            withType(defaultTypes(), Money, { convert: (m) => ({ v: m.cents / 100 }) }),
        );
        assert.deepEqual(types.convert(new Cents(300)), { v: 3 });
    });

    it('finds the same handler however many instances go through it', () => {
        // The lookup is remembered per prototype, so this is one walk and 99
        // hits — what it must not do is start answering something else.
        const types = typesFor(
            withType(defaultTypes(), Money, { convert: (m) => ({ v: m.cents }) }),
        );
        for (let k = 0; k < 100; k++) {
            assert.deepEqual(types.convert(new Money(k)), { v: k });
        }
    });

    it('finds a bigint by the class it would be written as', () => {
        assert.deepEqual(typesFor(defaultTypes()).convert(7n), { v: 7 });
    });

    it('claims nothing for an object that has no prototype at all', () => {
        assert.equal(typesFor(defaultTypes()).handlerFor(Object.create(null) as object), undefined);
    });

    it('knows nothing a map left out', () => {
        const types = typesFor(new Map());
        assert.equal(types.convert(new Date()), undefined);
        assert.equal(types.handlerFor(new Date()), undefined);
    });

    it('does not reach the Object entry by walking the chain', () => {
        // `handlerFor` stops before `Object.prototype`: a plain object has to
        // get past `isStyledCell` before the fallback can claim it.
        const types = typesFor(withType(defaultTypes(), Object, { convert: () => ({ v: 'any' }) }));
        assert.equal(types.handlerFor({}), undefined);
        assert.equal(types.handlerFor(new Money(1)), undefined);
        assert.ok(types.objectHandler);
    });

    it('converts through the Object entry once nothing else has claimed the value', () => {
        const types = typesFor(
            withType(defaultTypes(), Object, { convert: (o) => ({ v: JSON.stringify(o) }) }),
        );
        assert.deepEqual(types.convert({ a: 1 }), { v: '{"a":1}' });
        // A registered class still wins over it.
        assert.equal(types.convert(new Date(1970, 0, 1))?.v, 25569);
    });
});

describe('a value of a registered type, in a row', () => {
    class HourRange {
        constructor(
            readonly from: string,
            readonly to: string,
        ) {}
        toString(): string {
            return `${this.from} a ${this.to}`;
        }
    }

    const TYPES = withType(defaultTypes(), HourRange, {
        convert: (range) => ({ v: range.toString() }),
    });

    function rowXml(row: CellRow, types: TypeMap = TYPES): string {
        return cellRowXml(1, row, new StyleTable(), new ValueTypes(types, PLAIN));
    }

    it('is written as what its type made of it', () => {
        assert.equal(
            rowXml([new HourRange('8:00', '10:30')]),
            '<row r="1"><c r="A1" t="inlineStr"><is><t>8:00 a 10:30</t></is></c></row>',
        );
    });

    it('is a value and not a cell, so it keeps a style the cell put on it', () => {
        const styles = new StyleTable({ box: { bold: true } });
        const xml = cellRowXml(
            1,
            [{ v: new HourRange('8:00', '10:30'), s: 'box' }],
            styles,
            new ValueTypes(TYPES, PLAIN),
        );
        assert.match(xml, /<c r="A1" t="inlineStr" s="1">/);
        assert.match(xml, /<t>8:00 a 10:30<\/t>/);
    });

    it('lets a t written on the cell go over the one its type would have said', () => {
        // A BigInt too big for a cell is text; asking for a number is asking
        // for the rounded number.
        const big = BigInt(Number.MAX_SAFE_INTEGER) + 2n;
        assert.match(rowXml([big]), /t="inlineStr"/);
        assert.match(rowXml([{ v: big, t: 'n' }]), /^<row r="1"><c r="A1"><v>/);
    });

    it('refuses a class nobody registered, by name', () => {
        assert.throws(
            () => rowXml([new HourRange('8:00', '10:30')], defaultTypes()),
            /"HourRange" is not one of them/,
        );
    });

    it('still refuses an object that meant to be a cell and was spelled wrong', () => {
        assert.throws(() => rowXml([{ value: 1 } as never]), /"v", "s", "f", "t", "col", "colSpan" or "rowSpan"/);
    });

    it('refuses an instance with no class to name without pretending it has one', () => {
        // A prototype chain that never reaches `Object.prototype` has no
        // constructor to read a name off; there is still something to say.
        const instance = Object.create(Object.create(null) as object) as object;
        assert.throws(() => rowXml([instance as never]), /and this one is not one of them/);
    });

    it('measures a styled cell by the width its value\'s type gave', () => {
        const widths = new WidthMeter(50);
        cellRowXml(
            1,
            [{ v: new Date(2024, 0, 15), s: 'box' }],
            new StyleTable({ box: { bold: true } }),
            new ValueTypes(TYPES, PLAIN),
            undefined,
            widths,
        );
        // The serial is five digits; the date it stands for is ten characters.
        assert.equal(widths.columnWidths()[0], columnWidth('dd/mm/yyyy'.length));
    });

    it('sizes a column by the format a conversion asked for, not by the number under it', () => {
        // A duration as a fraction of a day: half an hour is
        // `0.020833333333333332`, twenty characters of a number nobody sees,
        // against the `0:30:00` the cell shows.
        class Interval {
            constructor(readonly ms: number) {}
        }
        const types = withType(defaultTypes(), Interval, {
            convert: (interval) => ({ v: interval.ms / 86400000, numFmt: '[h]:mm:ss' }),
        });
        const widths = new WidthMeter(50);
        cellRowXml(
            1,
            [new Interval(30 * 60 * 1000)],
            new StyleTable(),
            new ValueTypes(types, PLAIN),
            undefined,
            widths,
        );
        assert.equal(widths.columnWidths()[0], columnWidth('[h]:mm:ss'.length));
    });

    it('lets the Object entry take what would have been refused', () => {
        const types = withType(TYPES, Object, { convert: (o) => ({ v: Object.keys(o).join() }) });
        assert.match(rowXml([{ nope: 1 } as never], types), /<t>nope<\/t>/);
    });
});
