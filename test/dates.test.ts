import assert from 'node:assert/strict';
import {
    excelSerial,
    fromExcelSerial,
    hasTimeOfDay,
    kindOf,
    localClock,
    partsOfSerial,
    serialOfParts,
    utcClock,
    type DateClock,
    type DateParts,
} from '../src/core/dates.js';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
import { readXlsx } from '../src/core/read/readXlsx.js';
import { collect } from './helpers/streams.js';

/** The parts of a wall clock, with everything unsaid at zero. */
function parts(over: Partial<DateParts>): DateParts {
    return { year: 2024, month: 1, day: 15, hour: 0, minute: 0, second: 0, millisecond: 0, ...over };
}

describe('excelSerial', () => {
    it('is the day count from Excel\'s own epoch', () => {
        assert.equal(excelSerial(new Date(1970, 0, 1)), 25569);
        assert.equal(excelSerial(new Date(1970, 0, 2)), 25570);
    });

    it('is a whole number for a date with no time on it, wherever it is written', () => {
        // A sheet holds a wall clock, not an instant: taking `getTime()` for
        // the serial would move every date by the writer's own offset, which
        // west of Greenwich lands a midnight on the day before.
        for (const date of [new Date(2024, 0, 15), new Date(2024, 6, 15)]) {
            assert.ok(Number.isInteger(excelSerial(date)), `${date.toString()} is not a whole day`);
        }
    });

    it('keeps the time of day the caller reads off the date', () => {
        const noon = excelSerial(new Date(2024, 0, 15, 12, 0, 0));
        assert.equal(noon - Math.floor(noon), 0.5);
    });

    it('follows the offset of the date itself, daylight saving and all', () => {
        // Two dates six months apart are 182 whole days apart, even where
        // only one of them is in daylight saving time.
        const january = excelSerial(new Date(2024, 0, 15));
        const july = excelSerial(new Date(2024, 6, 15));
        assert.equal(july - january, 182);
    });

    it('numbers the days of 1900 the way a spreadsheet does', () => {
        // Excel's own numbering, which counts a 29/02/1900 that never was:
        // everything up to 28/02/1900 is one lower than the days since
        // 1899-12-30 that the arithmetic gives.
        assert.equal(excelSerial(new Date(1900, 0, 1)), 1);
        assert.equal(excelSerial(new Date(1900, 1, 28)), 59);
        assert.equal(excelSerial(new Date(1900, 2, 1)), 61);
    });

    it('keeps a time of day that carries no date, as serial zero does', () => {
        assert.equal(excelSerial(new Date(1899, 11, 31, 10, 30)), 0.4375);
    });

    it('refuses a date a spreadsheet has no number for', () => {
        assert.throws(() => excelSerial(new Date(1899, 11, 30)), RangeError);
        assert.throws(() => excelSerial(new Date(1885, 5, 20)), RangeError);
    });

    it('refuses an invalid date rather than writing NaN into a cell', () => {
        assert.throws(() => excelSerial(new Date('no es una fecha')), RangeError);
    });
});

describe('fromExcelSerial', () => {
    it('is excelSerial the other way round', () => {
        for (const date of [
            new Date(1900, 0, 1),
            new Date(1900, 1, 28),
            new Date(1900, 2, 1),
            new Date(1970, 0, 1),
            new Date(2024, 0, 15),
            new Date(2024, 0, 15, 12, 30),
            new Date(2024, 0, 15, 12, 30, 45),
            new Date(2024, 0, 15, 23, 59, 59, 999),
            new Date(2024, 6, 15),
            new Date(1899, 11, 31, 10, 30),
        ]) {
            assert.equal(
                fromExcelSerial(excelSerial(date)).getTime(),
                date.getTime(),
                `${date.toString()} did not survive the round trip`,
            );
        }
    });

    it('reads the wall clock the file was written with', () => {
        const date = fromExcelSerial(45306.5);
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
        assert.equal(date.getHours(), 12);
    });

    it('refuses the day that never was, and anything below zero', () => {
        assert.throws(() => fromExcelSerial(60), RangeError);
        assert.throws(() => fromExcelSerial(60.5), RangeError);
        assert.throws(() => fromExcelSerial(-1), RangeError);
    });
});

describe('hasTimeOfDay', () => {
    it('is what tells a date from a date and a time', () => {
        assert.equal(hasTimeOfDay(localClock.parts(new Date(2024, 0, 15))), false);
        assert.equal(hasTimeOfDay(localClock.parts(new Date(2024, 0, 15, 0, 0, 0, 1))), true);
        assert.equal(hasTimeOfDay(localClock.parts(new Date(2024, 0, 15, 12, 30))), true);
    });
});

describe('kindOf', () => {
    it('is the format a value falls back to, said as the value itself', () => {
        assert.equal(kindOf(localClock.parts(new Date(2024, 0, 15))), 'date');
        assert.equal(kindOf(localClock.parts(new Date(2024, 0, 15, 12, 30))), 'dateTime');
    });
});

describe('serialOfParts', () => {
    it('is a wall clock counted from Excel\'s day 0, with no zone anywhere in it', () => {
        assert.equal(serialOfParts(parts({ year: 1970, month: 1, day: 1 })), 25569);
        assert.equal(serialOfParts(parts({ hour: 12 })), 45306.5);
    });

    it('numbers a two-digit year as itself, not as the 1900s', () => {
        // `Date.UTC(50, ...)` is 1950, which would give the year 50 a serial
        // of its own instead of the refusal it has coming.
        assert.throws(() => serialOfParts(parts({ year: 50 })), RangeError);
        assert.throws(() => serialOfParts(parts({ year: 0 })), RangeError);
    });

    it('says which date it refused, with the time on it when there is one', () => {
        assert.throws(() => serialOfParts(parts({ year: 1885, month: 6, day: 20 })), /1885-06-20/);
        assert.throws(
            () => serialOfParts(parts({ year: 1885, month: 6, day: 20, hour: 9, minute: 5 })),
            /1885-06-20 09:05:00/,
        );
    });

    it('says as much as it can about a date that is not one at all', () => {
        const invalid = localClock.parts(new Date('no es una fecha'));
        assert.throws(() => serialOfParts(invalid), /An invalid date/);
    });
});

describe('partsOfSerial', () => {
    it('is serialOfParts the other way round, to the millisecond', () => {
        for (const value of [
            parts({ year: 1900, month: 1, day: 1 }),
            parts({ year: 1900, month: 3, day: 1 }),
            parts({ hour: 23, minute: 59, second: 59, millisecond: 999 }),
            parts({ year: 1899, month: 12, day: 31, hour: 10, minute: 30 }),
            parts({ year: 9999, month: 12, day: 31 }),
        ]) {
            assert.deepEqual(partsOfSerial(serialOfParts(value)), value);
        }
    });

    it('lands a serial under one on the day a bare time is stored on', () => {
        assert.deepEqual(
            partsOfSerial(0.4375),
            parts({ year: 1899, month: 12, day: 31, hour: 10, minute: 30 }),
        );
    });
});

describe('utcClock', () => {
    it('reads the wall clock toISOString shows, whatever the machine is set to', () => {
        const noon = new Date(Date.UTC(2024, 0, 15, 12, 30));
        assert.deepEqual(utcClock.parts(noon), parts({ hour: 12, minute: 30 }));
        assert.equal(excelSerial(noon, utcClock), 45306.52083333333);
    });

    it('is its own way back', () => {
        const written = parts({ month: 7, day: 4, hour: 6 });
        assert.deepEqual(utcClock.parts(utcClock.at(written)), written);
        assert.equal(utcClock.at(written).toISOString(), '2024-07-04T06:00:00.000Z');
    });

    it('and the local clock disagree by exactly the machine\'s offset', () => {
        const when = new Date(2024, 0, 15, 12, 0);
        const difference = excelSerial(when, utcClock) - excelSerial(when, localClock);
        assert.equal(Math.round(difference * 1440), when.getTimezoneOffset());
    });
});

describe('fromExcelSerial: the clock it is read with', () => {
    it('gives back the wall clock the serial says, under either one', () => {
        assert.equal(fromExcelSerial(45306.5, utcClock).toISOString(), '2024-01-15T12:00:00.000Z');
        assert.equal(fromExcelSerial(45306.5, localClock).getHours(), 12);
    });

    it('keeps a year a spreadsheet can hold but `new Date(y, …)` cannot', () => {
        // Serial 1 is 01/01/1900, and `new Date(1900, 0, 1)` is fine — the
        // trap is only under 100, which no serial reaches. What this pins is
        // that the way back does not go through the two-digit reading at all.
        assert.equal(fromExcelSerial(1, localClock).getFullYear(), 1900);
    });
});

describe('dateClock: the workbook option', () => {
    /** A one-cell workbook holding this date, written with that clock. */
    async function roundTrip(when: Date, dateClock?: 'local' | 'utc'): Promise<string> {
        const stream = createXlsxStream({ rows: [[when]], ...(dateClock ? { dateClock } : {}) });
        const bytes = new Uint8Array(await collect(stream));
        const [sheet] = await readXlsx(bytes, { dates: 'isoString' });
        return String(sheet?.cells[0]?.[0]);
    }

    const midnightUtc = new Date(Date.UTC(2024, 0, 15));

    it('writes the machine\'s own wall clock when nothing is said', async () => {
        const when = new Date(2024, 0, 15, 12, 30);
        assert.equal(await roundTrip(when), '2024-01-15T12:30:00');
        assert.equal(await roundTrip(when), await roundTrip(when, 'local'));
    });

    it('writes the UTC one when asked, so the file says the same thing anywhere', async () => {
        assert.equal(await roundTrip(midnightUtc, 'utc'), '2024-01-15');
    });

    it('takes a clock written out, which is how a third zone gets used', async () => {
        // Three hours behind UTC, fixed — the shape a caller's own clock has.
        const minus3: DateClock = {
            parts: (value) => utcClock.parts(new Date(value.getTime() - 3 * 3600000)),
            at: (parts) => new Date(utcClock.at(parts).getTime() + 3 * 3600000),
        };
        const stream = createXlsxStream({ rows: [[midnightUtc]], dateClock: minus3 });
        const bytes = new Uint8Array(await collect(stream));
        const [sheet] = await readXlsx(bytes, { dates: 'isoString' });
        assert.equal(sheet?.cells[0]?.[0], '2024-01-14T21:00:00');
    });
});
