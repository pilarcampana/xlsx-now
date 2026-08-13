import type { CellType } from './types.js';
import type { NativeValue } from './valueTypes.js';

// Days between 1900-01-01 and 1970-01-01 (Excel's epoch quirk on Windows).
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;
const MS_PER_MINUTE = 60000;

/**
 * The serial Excel gives to a day that never happened.
 *
 * Lotus 1-2-3 took 1900 for a leap year, Excel copied the bug on purpose to
 * stay compatible with it, and every spreadsheet since carries it: serial 60
 * is 29/02/1900, a date the Gregorian calendar does not have. Everything from
 * 01/03/1900 on is numbered one higher than a straight count of days would
 * make it, and everything before it is numbered as the count says.
 *
 * So the arithmetic above — days since 1899-12-30 — is right for one side of
 * that day and one short on the other, which is what these two constants are
 * for. The correction is a single comparison in each direction; the day
 * itself is the part with no answer, since there is no `Date` for it.
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
 * The characters XML 1.0 has no place for — §2.2 leaves them out of `Char`
 * altogether, so there is no spelling of them a parser will read back: not as
 * themselves, and not as the `&#0;` a numeric reference would be either.
 *
 * The three control characters that *are* allowed stay: tab, line feed and
 * carriage return. What is left is what a text field of a database ends up
 * carrying by accident — a truncated field, a stray byte of something binary
 * — and one of them is enough to make the whole file unreadable.
 */
const FORBIDDEN_IN_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/**
 * Text as XML can carry it: the five entities spelled out, and the characters
 * XML has no room for dropped.
 *
 * Dropping is the same answer `sheetName` gives to a character Excel forbids,
 * and for the same reason: a workbook with one of these in it is a file that
 * does not open at all, and by the time anyone finds that out the rows are
 * long gone. It is what `exceljs` does with them too. Keeping them by some
 * escape of Excel's own — `_x0000_` — would be the other way out, and it is a
 * way out of a different problem: this writer makes spreadsheets, not a
 * container for bytes that survive a round trip.
 *
 * A lone surrogate is not one of these. It is not a character either, but it
 * comes out of `TextEncoder` as the replacement character rather than as
 * anything a parser refuses, so the file still opens. Dropping it here would
 * mean taking apart the pairs that make up every emoji, which is a good deal
 * worse than what it would fix.
 */
export function sanitizeText(value: unknown): string {
    return String(value)
        .replace(FORBIDDEN_IN_XML, '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// 0-based column index -> spreadsheet column letters ("A", "B", ..., "AA", ...).
export function columnLetters(index: number): string {
    let n = index + 1;
    let letters = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

export function cellRef(colIndex: number, rowNumber: number): string {
    return `${columnLetters(colIndex)}${rowNumber}`;
}

/**
 * Column letters back to a 0-based index — `columnLetters` the other way
 * round, for the coordinates a cell is written by. Returns `undefined`
 * for anything that is not a column: the caller knows what to say about it.
 */
export function columnIndex(letters: string): number | undefined {
    if (!letters) return undefined;
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
        const code = letters.charCodeAt(i) & ~32; // upper-cases a letter, and only a letter
        if (code < 65 || code > 90) return undefined;
        n = n * 26 + (code - 64);
    }
    return n - 1;
}

/**
 * Which clock a `Date` is read by on the way into a sheet.
 *
 * A sheet has no time zone and a `Date` is an instant, so writing one down
 * means picking the clock it is read by, and there are only two answers:
 * `local` is the one the caller is looking at — `getFullYear()`,
 * `getHours()` — and `utc` is the one `getUTCHours()` gives. A `Date` built
 * from a local calendar is `local`; one that came from an ISO text with a `Z`
 * on it, or from a database that stores instants, is `utc`.
 */
export type WriteDates = 'local' | 'utc';

/** The `dates` a workbook was written with, checked. Defaults to `local`. */
export function writeDates(option: WriteDates | undefined): WriteDates {
    if (option === undefined) return 'local';
    if (option !== 'local' && option !== 'utc') {
        throw new Error(
            `dates: "${String(option)}" is not a clock a Date can be read by: say "local" or "utc".`,
        );
    }
    return option;
}

/** What a date is worth showing: the day, the day and the hour, or the hour. */
export type DateKind = 'date' | 'dateTime' | 'time';

/**
 * What a serial has to show. It is the number that says so and nothing else —
 * which is what makes the answer the same on both sides of the file: the
 * format the writer shows a value under, and the shape the reader builds one
 * into, are the same decision made twice.
 *
 * A whole number is a day. A fraction of one is the time of day next to it. And
 * under `1` there is no day left: serial 0 is Excel's own "day zero", where a
 * time with no date is stored, so `0.4375` is half past ten in the morning and
 * nothing else.
 */
export function serialKind(serial: number): DateKind {
    if (serial < 1) return 'time';
    return Number.isInteger(serial) ? 'date' : 'dateTime';
}

/**
 * A `Date` as the serial number a sheet stores, read by the clock `dates`
 * asks for.
 *
 * A sheet has no time zone: what it holds is a wall clock, so a date that
 * reads `2024-01-15 00:00` has to come out of the file reading the same.
 * `getTime()` is UTC, and taking that as the serial moves every date by the
 * writer's own offset — three hours in Buenos Aires, which is enough to land
 * a midnight on the day before. So under `local`, the default,
 * `getTimezoneOffset()` — what the date itself says its offset is, daylight
 * saving and all — is taken off first, and what gets written is the same
 * reading `getFullYear()` and `getHours()` give.
 *
 * Under `utc` that step is skipped and the instant goes in as it is, which is
 * the right answer for a `Date` that was never a local calendar to begin with.
 */
export function excelSerial(value: Date, dates: WriteDates = 'local'): number {
    const wall =
        dates === 'utc'
            ? value.getTime()
            : value.getTime() - value.getTimezoneOffset() * MS_PER_MINUTE;
    const days = wall / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
    // Days counted from 1899-12-30, which is the numbering Excel uses from
    // 01/03/1900 on. Before that its own count is one lower, because of the
    // 29/02/1900 it has and the calendar does not.
    const serial = days < FIRST_UNSHIFTED_SERIAL ? days - 1 : days;
    if (!(serial >= MIN_SERIAL)) {
        // `!(x >= 0)` and not `x < 0`, so a `NaN` — an invalid `Date` — is
        // caught here too rather than written out as `NaN` in a `<v>`.
        throw new RangeError(
            `${value.toISOString?.() ?? String(value)} cannot be written to a sheet: ` +
                'a spreadsheet numbers its days from 31/12/1899, and there is no serial for anything before that.',
        );
    }
    return serial;
}

/**
 * The serial number a sheet stores, back as the `Date` that *reads* as it in
 * UTC — `excelSerial(value, 'utc')` the other way round.
 *
 * The wall clock the serial holds goes in unmoved: serial `45306.5` comes back
 * as the instant `2024-01-15T12:00:00.000Z`, whatever zone the reader is in.
 * Which makes it the one form of a date that does not depend on where it is
 * read, and the one everything else is built from — the ISO text of a date, a
 * `Temporal.PlainDate`, and the local `Date` below.
 *
 * A serial under `1` carries a time of day and no meaningful date, and it
 * lands on 31/12/1899 — the day `excelSerial` sends it back from, so a time
 * survives the round trip.
 *
 * The one serial with no answer is 60, [the day Excel has and the calendar
 * does not](https://learn.microsoft.com/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year):
 * there is no `Date` for 29/02/1900, and the two candidates on either side of
 * it are dates of their own that a file can hold separately. Giving one of
 * them back would make two different serials read as the same day, so it is
 * refused instead.
 */
export function fromExcelSerialUtc(serial: number): Date {
    if (serial >= PHANTOM_LEAP_DAY_SERIAL && serial < FIRST_UNSHIFTED_SERIAL) {
        throw new RangeError(
            `Serial ${serial} is 29/02/1900, a day this spreadsheet format has and the calendar does not.`,
        );
    }
    if (!(serial >= MIN_SERIAL)) {
        throw new RangeError(`Serial ${serial} is not a date: a sheet numbers its days from 0.`);
    }
    const days = serial < PHANTOM_LEAP_DAY_SERIAL ? serial + 1 : serial;
    // Rounded to the millisecond, because that is as fine as a `Date` gets
    // and a serial is a fraction of a day: half past twelve is
    // `45306.520833333336`, and taken at face value it comes back a
    // millisecond short of the half hour it went in as.
    return new Date(Math.round((days - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY));
}

/**
 * The serial number a sheet stores, back as the `Date` that reads as it on the
 * caller's own clock — `excelSerial` the other way round.
 *
 * The serial says `2024-01-15 00:00` and the `Date` that comes back reads
 * `2024-01-15 00:00` to whoever asks it, wherever they are.
 */
export function fromExcelSerial(serial: number): Date {
    const wall = fromExcelSerialUtc(serial).getTime();
    // `wall` is the wall clock read as if it were UTC, and the instant that
    // shows that wall clock is it plus whatever the zone's offset is *at that
    // instant* — which is what makes this a fixed point rather than a sum:
    // the offset depends on the date it is being applied to.
    //
    // The first guess reads the offset up to fourteen hours away from the
    // right instant, which only matters near a daylight saving change; the
    // second reads it within an hour, which settles every case except the
    // hour a zone skips over — and that one is a wall clock that never
    // happened, so it has no exact answer to arrive at.
    const guess = new Date(wall + new Date(wall).getTimezoneOffset() * MS_PER_MINUTE);
    return new Date(wall + guess.getTimezoneOffset() * MS_PER_MINUTE);
}

/**
 * What a value is written as when the cell does not say: a number is a number,
 * a boolean is a boolean, and everything else — including a number XML has no
 * spelling for, like `NaN` — is text. An empty value never gets here: it is a
 * cell with no `<v>` at all, and `cellXml` has already answered for it.
 *
 * Only a value the writer already knew reaches this. Anything else — a date
 * among them — was turned into one of these by its type's own conversion, and
 * says its own type there when the answer is not the one below.
 */
function inferredType(value: NativeValue): CellType {
    if (typeof value === 'number') return Number.isFinite(value) ? 'n' : 'inlineStr';
    if (typeof value === 'boolean') return 'b';
    return 'inlineStr';
}

/**
 * The type attribute, left out when there is nothing to say: `n` is what a
 * `<c>` holds without a `t`, so writing it is six bytes per cell to repeat
 * the default — and numbers are most of what a sheet is made of.
 */
function typeAttribute(type: CellType): string {
    return type === 'n' ? '' : ` t="${type}"`;
}

/**
 * Whether the `<t>` has to ask for its whitespace to be kept.
 *
 * XML does not touch whitespace inside an element — `xml:space` is what the
 * *application* reading it is told, and what Excel is told there decides
 * whether `' 007 '` comes back with its spaces. Only the edges are ever
 * trimmed, so a string with none pays nothing for the attribute.
 */
function keepsSpaces(text: string): boolean {
    return /^\s|\s$/.test(text);
}

/** The `<is>` of an inline string: the value itself, in the cell, escaped. */
function inlineStringXml(value: NativeValue): string {
    // Asked of the sanitized text, which is the one being written: escaping
    // never touches the edges — none of the five entities is a space — but a
    // dropped character can leave a space at one, and then the `<t>` is a
    // `<t>` that starts with a space.
    const text = sanitizeText(value);
    return `<is><t${keepsSpaces(text) ? ' xml:space="preserve"' : ''}>${text}</t></is>`;
}

/** What goes inside the `<v>`, for every type that has one. */
function valueText(value: NativeValue): string {
    if (typeof value === 'boolean') return value ? '1' : '0';
    if (typeof value === 'number') return String(value);
    // A string here is a caller having said the type outright — a number that
    // has to stay as written, an error code — so it goes in as it came.
    return sanitizeText(value);
}

/** The formula as `<f>` carries it: without the `=`, and XML-safe. */
function formulaText(formula: string): string {
    return sanitizeText(formula.startsWith('=') ? formula.slice(1) : formula);
}

/**
 * Renders a single `<c>` element. `styleIndex` is a 0-based index into
 * styles.xml's `<cellXfs>`, or 0 (falsy) for the default style.
 *
 * A cell with nothing in it is nothing at all: an empty value under the
 * default style writes no element, which is what keeps the holes in a sheet
 * from costing anything. It is only once something else is asked for — a
 * style, a formula — that an empty cell has to be written down.
 */
export function cellXml(
    value: NativeValue,
    ref: string,
    styleIndex: number,
    formula?: string,
    type?: CellType,
): string {
    const s = styleIndex ? ` s="${styleIndex}"` : '';
    const empty = value === null || value === undefined || value === '';

    if (formula !== undefined) {
        // The value next to a formula is its cached result, not the cell's
        // own contents: without one the cell stays blank until a reader
        // recalculates it, which is the caller's call to make.
        //
        // A cached result lives in the `<v>`, whatever it holds, so text there
        // is `str` — the "formula string" of the spec — and never `inlineStr`,
        // which has no `<is>` to go to next to an `<f>`.
        if (empty) return `<c r="${ref}"${s}><f>${formulaText(formula)}</f></c>`;
        const cached = type ?? inferredType(value);
        const t = typeAttribute(cached === 'inlineStr' ? 'str' : cached);
        return `<c r="${ref}"${t}${s}><f>${formulaText(formula)}</f><v>${valueText(value)}</v></c>`;
    }

    // A type with nothing to type is still an empty cell.
    if (empty) return styleIndex ? `<c r="${ref}"${s}/>` : '';

    const t = type ?? inferredType(value);
    // Inline strings (not shared strings): keeps the writer stateless and
    // streamable. `str` asks for the same thing here — with no formula in the
    // cell there is no formula result to be the string of — so it is written
    // as the inline string it means.
    if (t === 'inlineStr' || t === 'str') {
        return `<c r="${ref}" t="inlineStr"${s}>${inlineStringXml(value)}</c>`;
    }
    return `<c r="${ref}"${typeAttribute(t)}${s}><v>${valueText(value)}</v></c>`;
}
