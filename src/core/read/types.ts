import type { PlainDate, PlainDateTime, PlainTime } from '../temporal.js';
import type { StyledCell } from '../types.js';

/**
 * A value as the reader gives it back: the four things a sheet holds, plus
 * whatever a number under a date format was asked to become.
 *
 * Deliberately narrower than the writer's `CellValue`, which is open to any
 * class the caller registered a type for. Nothing like that comes back out of
 * a file: what a sheet stores is a number, a string, a boolean or nothing,
 * and a date is the one of them the format gives a meaning to.
 *
 * Which of the date types it is, is the reader's `dates` option and nothing
 * else: they are all four in the union because the option is read at run time,
 * and a caller who picked one knows which one they picked. A `string` is in it
 * twice over, being both what a text cell holds and what `dates: 'isoString'`
 * gives back.
 *
 * `null` is a cell that is there and empty. A cell that is not there at all
 * is `undefined`, and the two are different on purpose — the same difference
 * the writer makes on the way in.
 */
export type ReadValue =
    | string
    | number
    | boolean
    | Date
    | PlainDate
    | PlainDateTime
    | PlainTime
    | null;

/**
 * What comes back for each cell, chosen by `mode`:
 *
 * - `values` — the value alone, which is what most callers want.
 * - `cells` — the `StyledCell` the writer takes, carrying the value in `v`,
 *   the formula in `f`, and the number format in `s`. Which means what a
 *   reader gives back can be handed straight to a writer.
 */
export type ReadMode = 'values' | 'cells';

/** What each mode gives back per cell. */
export interface ReadModes {
    values: ReadValue;
    cells: StyledCell;
}

/** One row of a sheet, as it is read. */
export interface ReadRow<C> {
    /** The row number the sheet gives it, counting from 1. */
    index: number;
    /**
     * The cells, by column index counting from 0. A position no cell was
     * written in is `undefined`, so a row with a hole in the middle has one
     * here too, and the array ends at the last cell the row actually has.
     */
    cells: (C | undefined)[];
}

/** One worksheet, read whole. */
export interface SheetData<C> {
    /** The name the workbook gives the sheet. */
    name: string;
    /**
     * The grid, by row and then by column, both counting from 0 — so `A1` is
     * `cells[0]?.[0]`.
     *
     * Dense in rows: there is an entry for every row up to the last one that
     * holds anything, and a row that holds nothing is an empty array. Ragged
     * in columns: each row ends at its own last cell, and `maxCol` is what
     * says how wide the sheet is as a whole.
     */
    cells: (C | undefined)[][];
    /** How many columns the widest row of the sheet reaches. */
    maxCol: number;
    /** How many rows the sheet reaches: the same as `cells.length`. */
    maxRow: number;
}
