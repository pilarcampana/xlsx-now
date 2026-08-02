import type { CellType, CellValue } from './types.js';

// Days between 1900-01-01 and 1970-01-01 (Excel's epoch quirk on Windows).
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;
const MS_PER_MINUTE = 60000;

export function sanitizeText(value: unknown): string {
    return String(value)
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
 * What a value is written as when the cell does not say: a date and a number
 * are numbers, a boolean is a boolean, and everything else — including a
 * number XML has no spelling for, like `NaN` — is text. An empty value never
 * gets here: it is a cell with no `<v>` at all, and `cellXml` has already
 * answered for it.
 */
function inferredType(value: CellValue): CellType {
    if (value instanceof Date) return 'n';
    if (typeof value === 'number') return Number.isFinite(value) ? 'n' : 'inlineStr';
    if (typeof value === 'boolean') return 'b';
    return 'inlineStr';
}

/** What goes inside the `<v>`, for every type that has one. */
function valueText(value: CellValue): string {
    if (value instanceof Date) return String(excelSerial(value));
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
    value: CellValue,
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
        const t = type ? ` t="${type}"` : '';
        return `<c r="${ref}"${t}${s}><f>${formulaText(formula)}</f>${empty ? '' : `<v>${valueText(value)}</v>`}</c>`;
    }

    // A type with nothing to type is still an empty cell.
    if (empty) return styleIndex ? `<c r="${ref}"${s}/>` : '';

    const t = type ?? inferredType(value);
    if (t === 'inlineStr') {
        // Inline strings (not shared strings): keeps the writer stateless/streamable.
        return `<c r="${ref}" t="inlineStr"${s}><is><t xml:space="preserve">${sanitizeText(value)}</t></is></c>`;
    }
    return `<c r="${ref}" t="${t}"${s}><v>${valueText(value)}</v></c>`;
}
