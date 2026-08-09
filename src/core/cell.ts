import type { CellType } from './types.js';
import type { NativeValue } from './valueTypes.js';

// Days between 1900-01-01 and 1970-01-01 (Excel's epoch quirk on Windows).
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;
const MS_PER_MINUTE = 60000;

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
 * A `Date` as the serial number a sheet stores, read as the clock the caller
 * is looking at.
 *
 * A sheet has no time zone: what it holds is a wall clock, so a date that
 * reads `2024-01-15 00:00` has to come out of the file reading the same.
 * `getTime()` is UTC, and taking that as the serial moves every date by the
 * writer's own offset — three hours in Buenos Aires, which is enough to land
 * a midnight on the day before. `getTimezoneOffset()` is what the date itself
 * says its offset is, daylight saving and all, so what gets written is the
 * same reading `getFullYear()` and `getHours()` give.
 */
export function excelSerial(value: Date): number {
    const local = value.getTime() - value.getTimezoneOffset() * MS_PER_MINUTE;
    return local / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
}

/** Whether a `Date` says anything past the day — what decides its format. */
export function hasTimeOfDay(value: Date): boolean {
    return (
        value.getHours() !== 0 ||
        value.getMinutes() !== 0 ||
        value.getSeconds() !== 0 ||
        value.getMilliseconds() !== 0
    );
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
