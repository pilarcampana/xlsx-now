// Column widths worked out from what the sheet holds. Nothing here writes XML
// and nothing here knows about zip entries: it counts characters as the cells
// go by and hands back one width per column, which is all `autoWidthMax` ever
// was.
import { DEFAULT_DATE_FORMATS, type DateFormats } from './styles.js';
import type { CellValue } from './types.js';

/**
 * How many characters a cell shows.
 *
 * A string is its own length and a number is the length of the digits it is
 * written as — not of the number format that may be shown over it, which is
 * the one thing this cannot know: a format is a style, and a style is a name
 * the workbook resolves at the end. A date is measured by the format it falls
 * back to, which is the workbook's own — and, when that is one of Excel's
 * built-in ones, by the widest date a locale writes under it, since the
 * spelling is the reader's to choose. A boolean is `TRUE` or `FALSE`, as
 * Excel spells it.
 *
 * An empty cell measures 0: it takes part in no width, so a column of blanks
 * is a column nobody asked to resize.
 */
export function cellTextLength(
    value: CellValue,
    dates: DateFormats = DEFAULT_DATE_FORMATS,
): number {
    if (value === null || value === undefined) return 0;
    if (value instanceof Date) return dates.textLength(value);
    if (typeof value === 'boolean') return value ? 4 : 5;
    return String(value).length;
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
    /** What a date is measured as, which is the workbook's own business. */
    private readonly dates: DateFormats;
    /** Whether this meter was given a maximum at all — a sheet's own answer. */
    readonly measures: boolean;

    constructor(max: number | undefined, dates: DateFormats = DEFAULT_DATE_FORMATS) {
        if (max !== undefined && !(Number.isFinite(max) && max > 0)) throw badMaxError(max);
        this.measures = max !== undefined;
        this.max = max ?? 0;
        this.dates = dates;
    }

    /** One cell, in the column it was written in. */
    see(column: number, value: CellValue): void {
        if (!this.measures) return;
        const length = cellTextLength(value, this.dates);
        if (!length) return;
        const width = length > this.max ? this.max : length;
        if (width > (this.widths[column] ?? 0)) this.widths[column] = width;
    }

    /**
     * What every column measured, by 0-based column. A column nobody wrote
     * anything in is a hole — it keeps whatever `columnFormats` says about it,
     * and Excel's default width when that says nothing either.
     */
    columnWidths(): readonly number[] {
        return this.widths;
    }
}
