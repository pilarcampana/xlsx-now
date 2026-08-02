// Column widths worked out from what the sheet holds. Nothing here writes XML
// and nothing here knows about zip entries: it counts characters as the cells
// go by and hands back one width per column, which is all `autoWidthMax` ever
// was.
import { hasTimeOfDay } from './cell.js';
import { DATE_FORMAT, DATETIME_FORMAT } from './styles.js';
import type { CellValue } from './types.js';

/**
 * How many characters a cell shows.
 *
 * A string is its own length and a number is the length of the digits it is
 * written as — not of the number format that may be shown over it, which is
 * the one thing this cannot know: a format is a style, and a style is a name
 * the workbook resolves at the end. A date is measured by the format it gets
 * (`yyyy-mm-dd`, or with the time of day when it says one), which is exactly
 * what the cell will show. A boolean is `TRUE` or `FALSE`, as Excel spells it.
 *
 * An empty cell measures 0: it takes part in no width, so a column of blanks
 * is a column nobody asked to resize.
 */
export function cellTextLength(value: CellValue): number {
    if (value === null || value === undefined) return 0;
    if (value instanceof Date) return (hasTimeOfDay(value) ? DATETIME_FORMAT : DATE_FORMAT).length;
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
 * One meter per worksheet — the widths are the sheet's own, and so is the
 * `<cols>` they end up in.
 */
export class WidthMeter {
    /** The longest cell seen per 0-based column; a hole is a column with nothing in it. */
    private readonly widths: number[] = [];
    private readonly max: number;

    constructor(max: number) {
        if (!Number.isFinite(max) || max <= 0) throw badMaxError(max);
        this.max = max;
    }

    /** One cell, in the column it was written in. */
    see(column: number, value: CellValue): void {
        const length = cellTextLength(value);
        if (!length) return;
        const width = length > this.max ? this.max : length;
        if (width > (this.widths[column] ?? 0)) this.widths[column] = width;
    }

    /**
     * What every column measured, by 0-based column. A column nobody wrote
     * anything in is left out — it keeps whatever `columnFormats` says about
     * it, and Excel's default width when that says nothing either.
     */
    columnWidths(): readonly (number | undefined)[] {
        return this.widths;
    }
}
