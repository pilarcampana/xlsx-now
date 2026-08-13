// The four answers to one number.
//
// The environment these tests run in has no `Temporal` of its own, so the
// polyfill is installed for the whole suite by `.mocharc.json` — before any of
// the package loads, which is the one order this package counts on. What is
// imported here is the same classes the global carries, for their types: the
// polyfill installs the global but declares nothing for it, and these tests
// are the one place that needs to name `Temporal.PlainDate` in TypeScript.
import assert from 'node:assert/strict';
import { Temporal } from 'temporal-polyfill';
import { readDate, readDates } from '../../src/core/read/dates.js';

/** 15/01/2024, and the same day at half past twelve. */
const DAY = 45306;
const NOON_AND_A_HALF = 45306.5;
/** Half past ten in the morning, of no day at all. */
const TIME = 0.4375;

describe('readDates', () => {
    it('reads dates as Temporal values when nothing was asked for', () => {
        assert.equal(readDates(undefined), 'temporal');
    });

    it('says what the modes are when it is handed something else', () => {
        assert.throws(
            () => readDates('plainDate' as 'temporal'),
            /"plainDate" is not how a date can be read.*temporal, utcDate, localDate, isoString/s,
        );
    });

    it('says so up front when temporal was asked for and there is no Temporal', async () => {
        const global = globalThis as { Temporal?: unknown };
        const temporal = global.Temporal;
        delete global.Temporal;
        try {
            // A second copy of the module, loaded the way an environment with
            // no `Temporal` would load it — which is the only way to be
            // without one, since the global is read when the module loads and
            // never again.
            const path = '../../src/core/temporal.js';
            const fresh = (await import(`${path}?no-temporal`)) as typeof import(
                '../../src/core/temporal.js'
            );
            assert.equal(fresh.temporalApi, undefined);
            assert.throws(() => fresh.requireTemporal(), /Temporal\.PlainDate/);
            assert.throws(
                () => fresh.requireTemporal(),
                /dates: "utcDate", "localDate" or "isoString"/,
            );
        } finally {
            global.Temporal = temporal;
        }

        // And the three modes that need no Temporal never ask for one.
        assert.equal(readDate(DAY, 'isoString'), '2024-01-15');
        assert.ok(readDate(DAY, 'utcDate') instanceof Date);
    });
});

describe('readDate: temporal', () => {
    it('builds a PlainDate out of a whole day', () => {
        const date = readDate(DAY, 'temporal') as Temporal.PlainDate;
        assert.ok(date instanceof Temporal.PlainDate);
        assert.equal(date.toString(), '2024-01-15');
    });

    it('builds a PlainDateTime out of a day with an hour in it', () => {
        const when = readDate(NOON_AND_A_HALF, 'temporal') as Temporal.PlainDateTime;
        assert.ok(when instanceof Temporal.PlainDateTime);
        assert.equal(when.toString(), '2024-01-15T12:00:00');
    });

    it('builds a PlainTime out of a serial with no day left in it', () => {
        const time = readDate(TIME, 'temporal') as Temporal.PlainTime;
        assert.ok(time instanceof Temporal.PlainTime);
        assert.equal(time.toString(), '10:30:00');
    });

    it('keeps the seconds a serial carries', () => {
        // 12:34:56 of the same day, as the fraction of a day it is.
        const serial = DAY + (12 * 3600 + 34 * 60 + 56) / 86400;
        assert.equal(String(readDate(serial, 'temporal')), '2024-01-15T12:34:56');
    });
});

describe('readDate: isoString', () => {
    it('is the text the Temporal values are built from', () => {
        assert.equal(readDate(DAY, 'isoString'), '2024-01-15');
        assert.equal(readDate(NOON_AND_A_HALF, 'isoString'), '2024-01-15T12:00:00');
        assert.equal(readDate(TIME, 'isoString'), '10:30:00');
    });

    it('keeps a millisecond where there is one, and writes none where there is not', () => {
        assert.equal(readDate(DAY + 0.5 + 1 / 86400000, 'isoString'), '2024-01-15T12:00:00.001');
    });
});

describe('readDate: the two Dates', () => {
    it('reads the wall clock in UTC, whatever zone the reader is in', () => {
        const utc = readDate(NOON_AND_A_HALF, 'utcDate') as Date;
        assert.equal(utc.toISOString(), '2024-01-15T12:00:00.000Z');
    });

    it('reads the same wall clock locally', () => {
        const local = readDate(NOON_AND_A_HALF, 'localDate') as Date;
        assert.equal(local.getFullYear(), 2024);
        assert.equal(local.getMonth(), 0);
        assert.equal(local.getDate(), 15);
        assert.equal(local.getHours(), 12);
    });

    it('is the same clock said twice: the two differ by the zone and nothing else', () => {
        const local = readDate(NOON_AND_A_HALF, 'localDate') as Date;
        const utc = readDate(NOON_AND_A_HALF, 'utcDate') as Date;
        assert.equal(local.getTime() - local.getTimezoneOffset() * 60000, utc.getTime());
        assert.equal(local.getHours(), utc.getUTCHours());
    });

    it('refuses the day the calendar does not have, whichever mode asked', () => {
        for (const mode of ['temporal', 'utcDate', 'localDate', 'isoString'] as const) {
            assert.throws(() => readDate(60, mode), RangeError, mode);
            assert.throws(() => readDate(-1, mode), RangeError, mode);
        }
    });
});
