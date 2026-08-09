// Column widths worked out from what the sheet holds. Nothing here writes XML
// and nothing here knows about zip entries: it counts characters as the cells
// go by and hands back one width per column, which is all `autoWidthMax` ever
// was.
import type { NativeValue } from './valueTypes.js';

/**
 * The longest line of a text that has more than one, in characters.
 *
 * Walked rather than split: this is the one path that allocates per cell if
 * it is written the obvious way, and the answer is a number. A `\r\n` counts
 * as the line break it is — the carriage return is not a character the cell
 * shows, so it is not one the column has to fit.
 */
function longestLine(text: string): number {
    let longest = 0;
    let from = 0;
    for (;;) {
        const at = text.indexOf('\n', from);
        const end = at < 0 ? text.length : at;
        const carriageReturn = end > from && text.charCodeAt(end - 1) === 13;
        const length = end - from - (carriageReturn ? 1 : 0);
        if (length > longest) longest = length;
        if (at < 0) return longest;
        from = at + 1;
    }
}

/**
 * How many characters a cell shows.
 *
 * A string is its own length and a number is the length of the digits it is
 * written as — not of the number format that may be shown over it, which is
 * the one thing this cannot know: a format is a style, and a style is a name
 * the workbook resolves at the end. A boolean is `TRUE` or `FALSE`, as Excel
 * spells it.
 *
 * `shown` is the answer a value's own type gave, for the values whose written
 * form says nothing about their shown one — a date is a serial here and ten
 * characters on the screen, and only the type that made it a serial knows
 * that.
 *
 * `wraps` is whether the cell's style wraps its text, and it is what decides
 * what a line break in the value means. A cell that wraps shows one line per
 * break, so what the column has to fit is the longest of them; a cell that
 * does not shows the text on one line, break and all, and there the whole
 * length is the answer — which is why this cannot be settled by looking at
 * the value alone. Excel is the one drawing the distinction: a `CHAR(10)` in
 * a cell without wrap text is not shown as a line break at all, which is why
 * Alt+Enter turns wrapping on as it inserts one.
 *
 * An empty cell measures 0: it takes part in no width, so a column of blanks
 * is a column nobody asked to resize.
 */
export function cellTextLength(value: NativeValue, shown?: number, wraps?: boolean): number {
    if (value === null || value === undefined) return 0;
    if (shown !== undefined) return shown;
    if (typeof value === 'boolean') return value ? 4 : 5;
    const text = String(value);
    // The scan is what a wrapping cell pays; every other cell — which is
    // nearly all of them — is out before the text is looked at.
    if (!wraps || text.indexOf('\n') < 0) return text.length;
    return longestLine(text);
}

/**
 * The widest digit of the normal font, in pixels — Calibri 11 at 96 dpi,
 * which is the font a workbook has until one of its styles says otherwise —
 * and the padding a column carries around its text: two pixels of margin on
 * each side, and one more for the gridline.
 */
const MAX_DIGIT_WIDTH = 7;
const COLUMN_PADDING_PIXELS = 5;

/**
 * A count of characters as the `width` a `<col>` carries.
 *
 * The two are not the same number, which is the whole of this function.
 * ECMA-376 §18.3.1.13 measures a column in multiples of the widest digit of
 * the normal font *plus the padding*, and stores it in 1/256ths:
 *
 * ```
 * width = Truncate([{characters} * {digit width} + {5px padding}] / {digit width} * 256) / 256
 * ```
 *
 * which is why Excel writes `8.7109375` for a column it autofitted to eight
 * characters, and not `8`. Writing the count itself leaves every column short
 * by that padding — the text ends up clipped, and a number or a date under it
 * comes out as `##########`, which is the visible half of the same bug.
 */
export function columnWidth(characters: number): number {
    const pixels = characters * MAX_DIGIT_WIDTH + COLUMN_PADDING_PIXELS;
    return Math.trunc((pixels / MAX_DIGIT_WIDTH) * 256) / 256;
}

/** Why a max nobody can size a column by was refused. */
function badMaxError(max: number): Error {
    return new Error(
        `"${max}" is not an autoWidthMax: it is the widest a column may get, in characters, ` +
            'so it has to be a number above 0.',
    );
}

/**
 * The widths of a sheet, as its cells go by: every cell is measured into the
 * column it lands in, and the column keeps the longest of them, up to the
 * maximum it was opened with.
 *
 * One meter per worksheet, and every worksheet has one — a sheet with no
 * `autoWidthMax` gets a meter that measures nothing, so nothing downstream
 * has to ask whether there is one. What it does have to ask is `measures`:
 * that is what says whether the sheet can go out as it is written or has to
 * wait for its last row.
 */
export class WidthMeter {
    /** The longest cell seen per 0-based column; a hole is a column with nothing in it. */
    private readonly widths: number[] = [];
    private readonly max: number;
    /** Whether this meter was given a maximum at all — a sheet's own answer. */
    readonly measures: boolean;

    constructor(max: number | undefined) {
        if (max !== undefined && !(Number.isFinite(max) && max > 0)) throw badMaxError(max);
        this.measures = max !== undefined;
        this.max = max ?? 0;
    }

    /** One cell, in the column it was written in. */
    see(column: number, value: NativeValue, shown?: number, wraps?: boolean): void {
        if (!this.measures) return;
        const length = cellTextLength(value, shown, wraps);
        if (!length) return;
        const width = length > this.max ? this.max : length;
        if (width > (this.widths[column] ?? 0)) this.widths[column] = width;
    }

    /**
     * What every column measured, as the `width` a `<col>` is written with —
     * the characters it counted, plus the padding Excel measures a column by.
     * A column nobody wrote anything in is a hole: it keeps whatever
     * `columnFormats` says about it, and Excel's default width when that says
     * nothing either.
     */
    columnWidths(): readonly number[] {
        // `map` keeps the holes as holes, which is what the sparse array is for.
        return this.widths.map(columnWidth);
    }
}
