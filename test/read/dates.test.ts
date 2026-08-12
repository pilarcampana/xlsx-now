// What a date comes back as, once there is a choice about it.
//
// Every test here reads the same file — one cell, one serial, one format —
// and the only thing that changes is the `dates` option. That is the point of
// the option: nothing about the file decides this.
import assert from 'node:assert/strict';
// Imported for the global it installs *and* for the global it declares: with
// this in the program, `src/core/temporal.ts` resolves its conditional types
// against a real `Temporal` rather than against its own fallback shapes, so
// the branch a consumer on TypeScript 6 or 7 compiles is the branch the tests
// compile too.
import 'temporal-polyfill/global';
import 'temporal-polyfill/types/global';
import {
    dateReaderOf,
    dateReaders,
    isoDates,
    localDates,
    serialDates,
    temporalDates,
    utcDates,
    type DateOption,
    type DateReader,
} from '../../src/core/read/dates.js';
import { readXlsx, type ReadOptions } from '../../src/core/read/readXlsx.js';
import type { PlainDate, PlainDateTime, PlainTime } from '../../src/core/temporal.js';
import { stylesOf, xlsxPackage } from '../helpers/package.js';

/**
 * The one assertion here that is not a test: it is checked by the compiler
 * and never runs.
 *
 * With a `Temporal` declared in the program — which the import above puts
 * there — this package's `PlainDate` has to *be* `Temporal.PlainDate` and not
 * a look-alike of it, or the assignments below do not compile. It is the half
 * of the conditional in `src/core/temporal.ts` that a consumer on TypeScript
 * 6 or 7 gets, and the only way to check it from a repository that is on 5.9
 * is to declare the global the way their `lib` does.
 */
export const typedAsTheRealTemporal: [
    Temporal.PlainDate,
    Temporal.PlainDateTime,
    Temporal.PlainTime,
] = [] as unknown as [PlainDate, PlainDateTime, PlainTime];

/**
 * A one-cell workbook: the serial, under the format at `styles[style]`.
 *
 * Written by hand rather than by this package's writer, so what the reader is
 * being asked about is the file and not a round trip.
 */
function fileOf(serial: number, styles: readonly (string | number)[], style = 0): Uint8Array {
    return xlsxPackage({
        sheets: { H: `<row r="1"><c r="A1" s="${style}"><v>${serial}</v></c></row>` },
        styles: stylesOf(styles),
    });
}

/** That cell, read with whatever `dates` says. */
async function cell<D extends DateOption = 'temporal'>(
    bytes: Uint8Array,
    options: ReadOptions<'values', D> = {},
): Promise<unknown> {
    const [sheet] = await readXlsx(bytes, options);
    return sheet?.cells[0]?.[0];
}

/** 15/01/2024, as a whole day and as half past ten in the morning. */
const DAY = 45306;
const MOMENT = 45306.4375;
/** Half past ten with no day at all: Excel's day 0. */
const TIME_ONLY = 0.4375;

describe('dates: temporal', () => {
    it('is what a sheet is read as when nothing was asked for', async () => {
        assert.equal(dateReaderOf(undefined), temporalDates);
        assert.equal(dateReaders.temporal, temporalDates);
    });

    it('reads a whole serial as a PlainDate, with no midnight invented on it', async () => {
        const value = await cell(fileOf(DAY, [14]));
        assert.equal(String(value), '2024-01-15');
        assert.equal(Object.getPrototypeOf(value), Temporal.PlainDate.prototype);
    });

    it('reads a serial with a time on it as a PlainDateTime', async () => {
        const value = await cell(fileOf(MOMENT, [14]));
        assert.equal(String(value), '2024-01-15T10:30:00');
        assert.equal(Object.getPrototypeOf(value), Temporal.PlainDateTime.prototype);
    });

    it('reads a time under a time format as a PlainTime, and not as day zero', async () => {
        const value = await cell(fileOf(TIME_ONLY, [21]));
        assert.equal(String(value), '10:30:00');
        assert.equal(Object.getPrototypeOf(value), Temporal.PlainTime.prototype);
    });

    it('follows the value and not the format when the two disagree', async () => {
        // A whole serial under the date-and-time format is a day: the format
        // says what the cell shows, and `00:00:00` is not something the value
        // says. The other way round is the same rule — a fraction under a
        // plain date format keeps its time rather than dropping it.
        assert.equal(String(await cell(fileOf(DAY, [22]))), '2024-01-15');
        assert.equal(String(await cell(fileOf(MOMENT, [14]))), '2024-01-15T10:30:00');
    });

    it('keeps an elapsed time past a day, which no PlainTime holds', async () => {
        // `[h]:mm:ss` over a serial of 1.5 is 36 hours. There is no
        // `PlainTime` for it, and dropping the day would be losing half of
        // the value, so it comes back counted from Excel's own day 0.
        const value = await cell(fileOf(1.5, ['[h]:mm:ss']));
        assert.equal(String(value), '1900-01-01T12:00:00');
    });

    it('reads a bare time under a date format as the day a sheet puts it on', async () => {
        assert.equal(String(await cell(fileOf(TIME_ONLY, [14]))), '1899-12-31T10:30:00');
    });
});

describe('dates: localDate and utcDate', () => {
    it('read the same wall clock, on two different clocks', async () => {
        const local = (await cell(fileOf(MOMENT, [14]), { dates: 'localDate' })) as Date;
        const utc = (await cell(fileOf(MOMENT, [14]), { dates: 'utcDate' })) as Date;
        assert.ok(local instanceof Date);
        assert.ok(utc instanceof Date);
        assert.equal(local.getHours(), 10);
        assert.equal(utc.getUTCHours(), 10);
        assert.equal(utc.toISOString(), '2024-01-15T10:30:00.000Z');
    });

    it('are the same instant only where the machine is on UTC', async () => {
        const local = (await cell(fileOf(MOMENT, [14]), { dates: 'localDate' })) as Date;
        const utc = (await cell(fileOf(MOMENT, [14]), { dates: 'utcDate' })) as Date;
        assert.equal((local.getTime() - utc.getTime()) / 60000, local.getTimezoneOffset());
    });

    it('ignore the format, since a Date is a whole timestamp whatever it says', async () => {
        const time = (await cell(fileOf(TIME_ONLY, [21]), { dates: 'utcDate' })) as Date;
        assert.equal(time.toISOString(), '1899-12-31T10:30:00.000Z');
    });
});

describe('dates: isoString', () => {
    it('writes the three shapes the temporal reader gives the three types', async () => {
        assert.equal(await cell(fileOf(DAY, [14]), { dates: 'isoString' }), '2024-01-15');
        assert.equal(
            await cell(fileOf(MOMENT, [14]), { dates: 'isoString' }),
            '2024-01-15T10:30:00',
        );
        assert.equal(await cell(fileOf(TIME_ONLY, [21]), { dates: 'isoString' }), '10:30:00');
    });

    it('says what Temporal says, which is the point of it', async () => {
        for (const [serial, style] of [[DAY, 14], [MOMENT, 14], [TIME_ONLY, 21]] as const) {
            const bytes = fileOf(serial, [style]);
            assert.equal(
                await cell(bytes, { dates: 'isoString' }),
                String(await cell(bytes, { dates: 'temporal' })),
            );
        }
    });

    it('puts no zone on the end, since the file has none to put there', async () => {
        assert.doesNotMatch(
            String(await cell(fileOf(MOMENT, [14]), { dates: 'isoString' })),
            /Z|[+-]\d\d:\d\d$/,
        );
    });

    it('keeps the milliseconds when there are any, and leaves them off when not', async () => {
        const withMs = DAY + 1 / 86400000;
        assert.equal(
            await cell(fileOf(withMs, [22]), { dates: 'isoString' }),
            '2024-01-15T00:00:00.001',
        );
    });
});

describe('dates: serial', () => {
    it('is the number as the file has it, with the format noted and nothing done', async () => {
        assert.equal(await cell(fileOf(MOMENT, [14]), { dates: 'serial' }), MOMENT);
    });

    it('is the only way to see a serial no calendar has', async () => {
        // 60 is the 29/02/1900 Excel counts and the calendar does not, and
        // every other reader refuses it rather than answering with one of the
        // two days beside it.
        assert.equal(await cell(fileOf(60, [14]), { dates: 'serial' }), 60);
        await assert.rejects(cell(fileOf(60, [14]), { dates: 'temporal' }), /29\/02\/1900/);
    });

    it('leaves a number that is not a date exactly where it was', async () => {
        assert.equal(await cell(fileOf(45306, [0]), { dates: 'serial' }), 45306);
    });
});

describe('dates: a reader of the caller\'s own', () => {
    it('is read with, and gives back whatever it gives back', async () => {
        const counted: number[] = [];
        const mine: DateReader<{ serial: number; kind: string }> = {
            read(serial, { kind }) {
                counted.push(serial);
                return { serial, kind };
            },
        };
        // The kind is the *format's* word — style 0 here is the built-in 14,
        // a date — and not the value's, which is what makes it worth handing
        // over: the value is already in the serial.
        assert.deepEqual(await cell(fileOf(MOMENT, [14]), { dates: mine }), {
            serial: MOMENT,
            kind: 'date',
        });
        assert.deepEqual(counted, [MOMENT]);
    });

    it('is told which of the three the format shows', async () => {
        const kinds: string[] = [];
        const mine: DateReader<null> = {
            read(_serial, { kind }) {
                kinds.push(kind);
                return null;
            },
        };
        for (const [serial, style] of [[DAY, 14], [MOMENT, 22], [TIME_ONLY, 21]] as const) {
            await cell(fileOf(serial, [style]), { dates: mine });
        }
        assert.deepEqual(kinds, ['date', 'dateTime', 'time']);
    });

    it('has its check run when the package is opened, before any row', async () => {
        let checked = 0;
        const mine: DateReader<null> = {
            check() {
                checked++;
                throw new Error('no thanks');
            },
            read: () => null,
        };
        // A file with no date in it at all: the check still runs, because
        // whether this works is a question about the runtime and not about
        // which file happened to arrive.
        await assert.rejects(cell(fileOf(42, [0]), { dates: mine }), /no thanks/);
        assert.equal(checked, 1);
    });
});

describe('the readers as values', () => {
    it('are the same objects the names stand for', () => {
        assert.deepEqual(Object.keys(dateReaders), [
            'temporal',
            'localDate',
            'utcDate',
            'isoString',
            'serial',
        ]);
        assert.equal(dateReaderOf('localDate'), localDates);
        assert.equal(dateReaderOf('utcDate'), utcDates);
        assert.equal(dateReaderOf('isoString'), isoDates);
        assert.equal(dateReaderOf('serial'), serialDates);
    });

    it('are the same thing whether named or handed over', async () => {
        assert.equal(
            await cell(fileOf(MOMENT, [14]), { dates: 'isoString' }),
            await cell(fileOf(MOMENT, [14]), { dates: isoDates }),
        );
    });
});
