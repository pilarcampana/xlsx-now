// What a date is, between a spreadsheet and a runtime.
//
// A sheet stores a date as a number of days since an epoch, and nothing else:
// no zone, no offset, no instant. What that number names is a *civil* date —
// a wall clock — and every date type this package reads or writes is that
// number said in some other way. So the arithmetic is here, once, and the
// types are handled where they belong: the reader turns a serial into
// whatever the caller asked for, and the writer turns whatever the caller had
// back into a serial. Both meet in `DateParts`, which is the wall clock with
// no type on it at all.
//
// The one thing that is *not* arithmetic is which wall clock a `Date` shows.
// A `Date` is an instant, and an instant reads differently in every zone, so
// turning one into civil parts is a choice rather than a calculation. That
// choice is a `DateClock`, and it is the only place in this file where a zone
// is so much as mentioned.

/** Days between 1899-12-30 and 1970-01-01, which is what a serial counts from. */
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;
const MS_PER_HOUR = 3600000;
const MS_PER_MINUTE = 60000;
const MS_PER_SECOND = 1000;

/**
 * The serial Excel gives to a day that never happened.
 *
 * Lotus 1-2-3 took 1900 for a leap year, Excel copied the bug on purpose to
 * stay compatible with it, and every spreadsheet since carries it: serial 60
 * is 29/02/1900, a date the Gregorian calendar does not have. Everything from
 * 01/03/1900 on is numbered one higher than a straight count of days would
 * make it, and everything before it is numbered as the count says.
 *
 * So the arithmetic below — days since 1899-12-30 — is right for one side of
 * that day and one short on the other, which is what these two constants are
 * for. The correction is a single comparison in each direction; the day
 * itself is the part with no answer, since no calendar has it.
 */
const PHANTOM_LEAP_DAY_SERIAL = 60;
/** First serial the plain day count already agrees with: 01/03/1900. */
const FIRST_UNSHIFTED_SERIAL = 61;
/**
 * The lowest serial a cell can hold. Zero is Excel's own "day 0", which is
 * how a time of day with no date is stored — `0.4375` is half past ten in the
 * morning and nothing else — so it is a value to keep, not one to refuse.
 * Below it there is nothing: a negative serial is a date Excel has no
 * numbering for and shows as `######`.
 */
const MIN_SERIAL = 0;

/**
 * A wall clock, with nothing standing in for it: the reading a spreadsheet
 * stores, taken apart. No zone, no offset, no instant — those are what a
 * `DateClock` adds and takes away at the edges.
 *
 * `month` counts from 1, the way a calendar does and the way `Temporal` does,
 * and not from 0 the way `Date` alone among them does.
 */
export interface DateParts {
    year: number;
    /** 1 to 12. */
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
    millisecond: number;
}

/**
 * What a value says: a day, a time of day, or both.
 *
 * It is the one piece of vocabulary the reader and the writer share about a
 * date. The reader works it out from the number format a cell is under — that
 * is the only thing in a file that tells a date from a number — and the
 * writer works it out from the value it was handed. Both end up choosing
 * between the same three things, which is why it is one type and not two.
 */
export type DateKind = 'date' | 'time' | 'dateTime';

/** Whether these parts say anything past the day. */
export function hasTimeOfDay(parts: DateParts): boolean {
    return (
        parts.hour !== 0 || parts.minute !== 0 || parts.second !== 0 || parts.millisecond !== 0
    );
}

/**
 * What a wall clock with a day in it amounts to: a date, or a date and a
 * time. Never a bare time — parts that came from a `Date` or from a serial
 * always carry a day, whether or not it means anything, and the caller that
 * knows the day is meaningless is the one that says `'time'` itself.
 */
export function kindOf(parts: DateParts): DateKind {
    return hasTimeOfDay(parts) ? 'dateTime' : 'date';
}

/**
 * Days from 1970-01-01 to a civil date, and back.
 *
 * [Howard Hinnant's `days_from_civil`](https://howardhinnant.github.io/date_algorithms.html),
 * which is exact for every year a spreadsheet can hold and, more to the point,
 * touches no `Date` at all. Going through `Date.UTC` would be shorter and
 * wrong in one place that matters: it reads a two-digit year as 19xx, so a
 * date in the year 50 would come out as 1950 and be given a serial instead of
 * being refused.
 */
function daysFromCivil(year: number, month: number, day: number): number {
    // March-based years, so a leap day is the last day rather than one in the
    // middle: that is the whole trick, and what makes the rest arithmetic.
    const shifted = year - (month <= 2 ? 1 : 0);
    const era = Math.floor(shifted / 400);
    const yearOfEra = shifted - era * 400;
    const dayOfYear = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
    const dayOfEra =
        yearOfEra * 365 +
        Math.floor(yearOfEra / 4) -
        Math.floor(yearOfEra / 100) +
        dayOfYear;
    return era * 146097 + dayOfEra - 719468;
}

/** `daysFromCivil` the other way round. */
function civilFromDays(days: number): { year: number; month: number; day: number } {
    const shifted = days + 719468;
    const era = Math.floor(shifted / 146097);
    const dayOfEra = shifted - era * 146097;
    const yearOfEra = Math.floor(
        (dayOfEra -
            Math.floor(dayOfEra / 1460) +
            Math.floor(dayOfEra / 36524) -
            Math.floor(dayOfEra / 146096)) /
            365,
    );
    const year = yearOfEra + era * 400;
    const dayOfYear =
        dayOfEra - (365 * yearOfEra + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100));
    const monthOfYear = Math.floor((5 * dayOfYear + 2) / 153);
    const day = dayOfYear - Math.floor((153 * monthOfYear + 2) / 5) + 1;
    const month = monthOfYear + (monthOfYear < 10 ? 3 : -9);
    return { year: year + (month <= 2 ? 1 : 0), month, day };
}

/** A wall clock as a message can name it, invalid ones included. */
function spell(parts: DateParts): string {
    if (!Number.isFinite(parts.year)) return 'An invalid date';
    const pad = (value: number, width = 2): string => String(value).padStart(width, '0');
    const day = `${pad(parts.year, 4)}-${pad(parts.month)}-${pad(parts.day)}`;
    if (!hasTimeOfDay(parts)) return day;
    return `${day} ${pad(parts.hour)}:${pad(parts.minute)}:${pad(parts.second)}`;
}

/**
 * A wall clock as the serial number a sheet stores.
 *
 * The whole part is the day and the fraction is the time — `0.5` is midday,
 * whatever day it is on — so a date with no time on it is a whole number, and
 * that is what makes a serial a wall clock rather than an instant.
 */
export function serialOfParts(parts: DateParts): number {
    const timeOfDay =
        parts.hour * MS_PER_HOUR +
        parts.minute * MS_PER_MINUTE +
        parts.second * MS_PER_SECOND +
        parts.millisecond;
    const fromEpoch = daysFromCivil(parts.year, parts.month, parts.day) * MS_PER_DAY + timeOfDay;
    const days = fromEpoch / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
    // Days counted from 1899-12-30, which is the numbering Excel uses from
    // 01/03/1900 on. Before that its own count is one lower, because of the
    // 29/02/1900 it has and the calendar does not.
    const serial = days < FIRST_UNSHIFTED_SERIAL ? days - 1 : days;
    if (!(serial >= MIN_SERIAL)) {
        // `!(x >= 0)` and not `x < 0`, so a `NaN` — an invalid `Date` — is
        // caught here too rather than written out as `NaN` in a `<v>`.
        throw new RangeError(
            `${spell(parts)} cannot be written to a sheet: a spreadsheet numbers its days ` +
                'from 31/12/1899, and there is no serial for anything before that.',
        );
    }
    return serial;
}

/**
 * The serial number a sheet stores, back as the wall clock it names —
 * `serialOfParts` the other way round, and the whole of what the reader knows
 * about dates before a type is chosen for them.
 *
 * A serial under `1` carries a time of day and no meaningful date, and it
 * lands on 31/12/1899 — the day `serialOfParts` sends it back from, so a time
 * survives the round trip.
 *
 * The one serial with no answer is 60, [the day Excel has and the calendar
 * does not](https://learn.microsoft.com/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year):
 * there is no 29/02/1900, and the two days on either side of it are dates of
 * their own that a file can hold separately. Giving one of them back would
 * make two different serials read as the same day, so it is refused instead.
 */
export function partsOfSerial(serial: number): DateParts {
    if (serial >= PHANTOM_LEAP_DAY_SERIAL && serial < FIRST_UNSHIFTED_SERIAL) {
        throw new RangeError(
            `Serial ${serial} is 29/02/1900, a day this spreadsheet format has and the calendar does not.`,
        );
    }
    if (!(serial >= MIN_SERIAL)) {
        throw new RangeError(`Serial ${serial} is not a date: a sheet numbers its days from 0.`);
    }
    const days = serial < PHANTOM_LEAP_DAY_SERIAL ? serial + 1 : serial;
    // Rounded to the millisecond, because that is as fine as this goes and a
    // serial is a fraction of a day: half past twelve is `45306.520833333336`,
    // and taken at face value it comes back a millisecond short of the half
    // hour it went in as.
    const fromEpoch = Math.round((days - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY);
    // Floor and not truncate: before 1970 the count is negative, and there
    // the two are a day apart.
    const dayNumber = Math.floor(fromEpoch / MS_PER_DAY);
    const timeOfDay = fromEpoch - dayNumber * MS_PER_DAY;
    return {
        ...civilFromDays(dayNumber),
        hour: Math.floor(timeOfDay / MS_PER_HOUR),
        minute: Math.floor(timeOfDay / MS_PER_MINUTE) % 60,
        second: Math.floor(timeOfDay / MS_PER_SECOND) % 60,
        millisecond: timeOfDay % MS_PER_SECOND,
    };
}

/**
 * Which wall clock a `Date` shows — the one question about dates that has no
 * arithmetic answer.
 *
 * A `Date` is an instant. A sheet holds a wall clock. Between the two there
 * is a zone, and picking one is a decision rather than a calculation: the
 * same instant is `2024-01-15 00:00` in Buenos Aires and `2024-01-15 03:00`
 * in UTC, and a spreadsheet has room for exactly one of them.
 *
 * Two are built in, and a third is a caller's to write: a clock fixed to a
 * zone that is neither the machine's nor UTC is these same two methods with
 * an `Intl.DateTimeFormat` inside them. That is the reason this is an
 * interface and not a `'local' | 'utc'` the writer switches on.
 */
export interface DateClock {
    /** The wall clock this shows for an instant. */
    parts(value: Date): DateParts;
    /** The instant that shows this wall clock — `parts` the other way round. */
    at(parts: DateParts): Date;
}

/**
 * The machine's own zone: the reading `getFullYear()` and `getHours()` give.
 *
 * It is the default because it is what the caller is looking at. A date built
 * as `new Date(2024, 0, 15)` reads midnight of the 15th to whoever wrote it,
 * and a sheet that stored the instant instead would show the 14th to everyone
 * west of Greenwich — a value silently off by a day, which is the one outcome
 * worth going out of the way to avoid.
 */
export const localClock: DateClock = {
    parts: (value) => ({
        year: value.getFullYear(),
        month: value.getMonth() + 1,
        day: value.getDate(),
        hour: value.getHours(),
        minute: value.getMinutes(),
        second: value.getSeconds(),
        millisecond: value.getMilliseconds(),
    }),
    at(parts) {
        // Built by setting the fields rather than through the constructor:
        // `new Date(50, 0, 1)` is the year 1950, and a spreadsheet can hold
        // the year 50. The three date fields go in together so a day that
        // does not exist in the month standing there at the time — the 31st
        // of whatever February the epoch left behind — cannot roll over.
        const date = new Date(0);
        date.setFullYear(parts.year, parts.month - 1, parts.day);
        date.setHours(parts.hour, parts.minute, parts.second, parts.millisecond);
        return date;
    },
};

/**
 * UTC: the reading `getUTCFullYear()` and `getUTCHours()` give.
 *
 * What it is for is data that was never local to anybody — a timestamp out of
 * a database, an instant off a wire — where the machine's zone is an accident
 * of where the file happens to be written and the same input would land on
 * two different days on two different laptops.
 */
export const utcClock: DateClock = {
    parts: (value) => ({
        year: value.getUTCFullYear(),
        month: value.getUTCMonth() + 1,
        day: value.getUTCDate(),
        hour: value.getUTCHours(),
        minute: value.getUTCMinutes(),
        second: value.getUTCSeconds(),
        millisecond: value.getUTCMilliseconds(),
    }),
    at(parts) {
        const date = new Date(0);
        date.setUTCFullYear(parts.year, parts.month - 1, parts.day);
        date.setUTCHours(parts.hour, parts.minute, parts.second, parts.millisecond);
        return date;
    },
};

/** The clocks that have a name, for the option that takes one. */
export const dateClocks = {
    local: localClock,
    utc: utcClock,
} as const satisfies Record<string, DateClock>;

/** A built-in clock, as an option names it. */
export type DateClockName = keyof typeof dateClocks;

/** What a workbook says about which wall clock its `Date`s are read as. */
export interface DateClockOptions {
    /**
     * Which reading of a `Date` goes into the file: `'local'` — the default —
     * writes the wall clock the machine shows, and `'utc'` writes the one
     * `toISOString()` shows. A `DateClock` written out is how a third zone
     * gets used.
     *
     * It has no effect on a `Temporal.PlainDate` or on anything else that is
     * already a wall clock, since there is nothing there to choose between.
     */
    dateClock?: DateClockName | DateClock;
}

/** The clock a workbook uses when it said nothing: the machine's own. */
export const DEFAULT_CLOCK = localClock;

/** A clock as an option gives it: one of the two by name, or one written out. */
export function clockOf(clock: DateClockName | DateClock | undefined): DateClock {
    if (clock === undefined) return DEFAULT_CLOCK;
    return typeof clock === 'string' ? dateClocks[clock] : clock;
}

/**
 * A `Date` as the serial number a sheet stores, read as whichever clock was
 * asked for.
 */
export function excelSerial(value: Date, clock: DateClock = DEFAULT_CLOCK): number {
    return serialOfParts(clock.parts(value));
}

/**
 * The serial number a sheet stores, back as a `Date` — `excelSerial` the
 * other way round.
 *
 * The wall clock is kept the way it was written: the serial says
 * `2024-01-15 00:00` and the `Date` that comes back reads `2024-01-15 00:00`
 * to whoever asks it with the same clock, wherever they are.
 */
export function fromExcelSerial(serial: number, clock: DateClock = DEFAULT_CLOCK): Date {
    return clock.at(partsOfSerial(serial));
}
