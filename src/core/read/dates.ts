// What a date comes back as, and the choice of it.
//
// A sheet stores a date as a number under a number format, and that is all it
// stores. Which runtime type that becomes is not the file's business — it is
// the caller's, and every answer to it is a lossless one: `Temporal.PlainDate`
// is the closest thing JavaScript has to what a sheet actually holds, a
// `Date` is what most code around it still speaks, an ISO string is what
// crosses a wire, and the serial itself is what was in the file.
//
// So it is an option and not a decision, and the option takes a *reader*
// rather than a name from a fixed list. The four below are the ones worth
// shipping; a fifth — a Luxon `DateTime`, a `dayjs`, a class of the caller's
// own — is the same interface implemented outside this package, with nothing
// here to change to allow it. Which is the reader's half of what `types` is
// on the writer's side: one open question, asked once.
import {
    hasTimeOfDay,
    partsOfSerial,
    utcClock,
    type DateKind,
    type DateParts,
    localClock,
} from '../dates.js';
import { requireTemporal, type TemporalDate } from '../temporal.js';

/** What a reader is told about the cell beyond the serial in it. */
export interface DateContext {
    /**
     * What the cell's number format shows: a day, a time of day, or both.
     *
     * It is the only thing in a file that says a number is a date at all, so
     * it is the only thing a reader has to go on beyond the number itself.
     * What a reader does with it is its own business — `localDate` ignores it
     * outright, since a `Date` is a full timestamp whatever the format says.
     */
    readonly kind: DateKind;
}

/**
 * How a serial becomes a value.
 *
 * An object with one method rather than a bare function, for the same reason
 * `TypeHandler` is one on the writer's side: `check` is already a second
 * member, and a reader that grows a third does not make every call site move.
 */
export interface DateReader<T> {
    /**
     * The value this serial comes back as. The serial is the sheet's own, with
     * a 1904 workbook's epoch already corrected for, so a reader never has to
     * know which of the two the file counts from.
     */
    read(serial: number, context: DateContext): T;
    /**
     * Whether this runtime can do it at all, asked once when a package is
     * opened. Throw from here and the failure lands before the first row
     * rather than somewhere in the middle of a million of them — and it lands
     * on a workbook with no dates in it too, so whether the code runs does not
     * depend on what happened to be in the file that arrived.
     */
    check?(): void;
}

/** The value a reader gives back. */
export type DateOfReader<R> = R extends DateReader<infer T> ? T : never;

/**
 * A serial as `Temporal`, which is the type a spreadsheet's own idea of a date
 * translates into without losing anything.
 *
 * Three types, and which one it is follows the value:
 *
 * - a whole serial is a `PlainDate` — a day, with no time pretended onto it;
 * - a serial with a fraction is a `PlainDateTime`;
 * - a serial under 1 whose format shows only a time is a `PlainTime`, since
 *   the day it lands on (31/12/1899) is Excel's placeholder and not a date
 *   anybody wrote.
 *
 * The value decides and not the format, which is the same rule the writer
 * follows in reverse: a `PlainDate` goes out as a whole serial under a date
 * format, and comes back a `PlainDate`. The cost is that one column can come
 * back as two types where one row of it carries a time — and the alternative,
 * taking the format's word for it, is a time of day silently dropped from a
 * cell whose format only mentions the day.
 *
 * A time format over a serial of 1 or more is elapsed time — `[h]:mm` can say
 * 36 hours — and there is no `PlainTime` that holds it, so it comes back as a
 * `PlainDateTime` counted from Excel's own day 0.
 */
export const temporalDates: DateReader<TemporalDate> = {
    check: requireTemporal,
    read(serial, { kind }) {
        const api = requireTemporal();
        const parts = partsOfSerial(serial);
        if (kind === 'time' && serial < 1) {
            return new api.PlainTime(parts.hour, parts.minute, parts.second, parts.millisecond);
        }
        if (!hasTimeOfDay(parts)) return new api.PlainDate(parts.year, parts.month, parts.day);
        return new api.PlainDateTime(
            parts.year,
            parts.month,
            parts.day,
            parts.hour,
            parts.minute,
            parts.second,
            parts.millisecond,
        );
    },
};

/**
 * A serial as a `Date` reading the same wall clock on the machine that reads
 * it: the serial says `2024-01-15 00:00` and `getFullYear()`/`getHours()` say
 * the same, wherever that machine is.
 *
 * It is what this reader gave back before there was anything to choose, and
 * it is what a `Date` written by this package's writer comes back as
 * unchanged. What it is not is an instant anybody agreed on — two machines in
 * two zones read one file as two different moments — which is not this
 * reader's doing but the file's: a sheet has no zone to have written down.
 */
export const localDates: DateReader<Date> = {
    read: (serial) => localClock.at(partsOfSerial(serial)),
};

/**
 * The same, read as UTC: the serial says `2024-01-15 00:00` and
 * `toISOString()` says `2024-01-15T00:00:00.000Z`.
 *
 * The one to use when the file is data rather than somebody's spreadsheet,
 * and the dates in it have to mean the same thing on every machine that opens
 * it.
 */
export const utcDates: DateReader<Date> = {
    read: (serial) => utcClock.at(partsOfSerial(serial)),
};

/** Two digits, or however many are asked for. */
function pad(value: number, width = 2): string {
    return String(value).padStart(width, '0');
}

/** The time of day as ISO writes it, with the milliseconds only if there are any. */
function isoTime(parts: DateParts): string {
    const base = `${pad(parts.hour)}:${pad(parts.minute)}:${pad(parts.second)}`;
    return parts.millisecond === 0 ? base : `${base}.${pad(parts.millisecond, 3)}`;
}

/**
 * A serial as the text ISO 8601 gives it: `2024-01-15`, `2024-01-15T10:30:00`,
 * or `10:30:00`.
 *
 * Which of the three follows exactly the rule `temporalDates` follows, and on
 * purpose — this is what those same three values print as. So it is the
 * answer for a runtime with no `Temporal` and for anything that was going to
 * serialize the date anyway, and the two agree.
 *
 * No zone and no `Z` on the end, because there is none in the file: a `Z`
 * would be this reader inventing the one thing a sheet does not store.
 */
export const isoDates: DateReader<string> = {
    read(serial, { kind }) {
        const parts = partsOfSerial(serial);
        if (kind === 'time' && serial < 1) return isoTime(parts);
        const day = `${pad(parts.year, 4)}-${pad(parts.month)}-${pad(parts.day)}`;
        return hasTimeOfDay(parts) ? `${day}T${isoTime(parts)}` : day;
    },
};

/**
 * The number as it is in the file, with the date format noted and nothing
 * made of it.
 *
 * For code that does its own arithmetic on serials, and for reading a file
 * whose dates are out of range for every other answer — a serial of 60 is the
 * 29/02/1900 no calendar has, and it is a value this is the only way to see.
 */
export const serialDates: DateReader<number> = {
    read: (serial) => serial,
};

/**
 * The readers that have a name, for the option that takes one.
 *
 * `temporal` is the default because it is the only one of them that does not
 * have to decide something the file never said — a zone, a text encoding, or
 * that a day is a midnight.
 */
export const dateReaders = {
    temporal: temporalDates,
    localDate: localDates,
    utcDate: utcDates,
    isoString: isoDates,
    serial: serialDates,
} as const satisfies Record<string, DateReader<unknown>>;

/** A built-in reader, as an option names it. */
export type DateReaderName = keyof typeof dateReaders;

/** What the reader gives back when nothing was asked for. */
export const DEFAULT_DATE_READER: DateReaderName = 'temporal';

/** How an option is spelled: one of the five by name, or one written out. */
export type DateOption = DateReaderName | DateReader<unknown>;

/**
 * The type a `dates` option gives back, worked out from how it was written —
 * so `{ dates: 'isoString' }` reads a sheet of strings and a reader of the
 * caller's own reads a sheet of whatever it returns.
 */
export type DateOf<D> = D extends DateReaderName
    ? DateOfReader<(typeof dateReaders)[D]>
    : DateOfReader<D>;

/** A reader as an option gives it. */
export function dateReaderOf(option: DateOption | undefined): DateReader<unknown> {
    if (option === undefined) return dateReaders[DEFAULT_DATE_READER];
    return typeof option === 'string' ? dateReaders[option] : option;
}
