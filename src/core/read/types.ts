import type { StyledCell } from '../types.js';
import type { TemporalDate } from '../temporal.js';

/**
 * A value as the reader gives it back: the four things a sheet holds, plus
 * whatever a number under a date format was asked to become.
 *
 * `D` is that last one, and it is the `dates` option said as a type — a
 * `Temporal.PlainDate` by default, a `Date` under `'localDate'`, a `string`
 * under `'isoString'`. The rest is not open to anything: what a sheet stores
 * is a number, a string, a boolean or nothing, and the date is the one of
 * them the format gives a second meaning to.
 *
 * `null` is a cell that is there and empty. A cell that is not there at all
 * is `undefined`, and the two are different on purpose — the same difference
 * the writer makes on the way in.
 */
export type ReadValue<D = TemporalDate> = string | number | boolean | D | null;

/**
 * The same with the date type left open: what a cell is read as before the
 * `dates` option's answer is put back on it.
 *
 * It exists because the value and the type it is known by are settled in two
 * different places. Reading a cell is one function for every option, so what
 * it gives back has to be wide enough for all of them; which of them it
 * actually is comes from `openXlsx`'s signature, and a value cannot carry
 * that. `object` is as narrow as the open end can be said, since a reader of
 * the caller's own returns a class this package has never heard of.
 */
export type RawReadValue = string | number | bigint | boolean | object | null;

/**
 * What comes back for each cell, chosen by `mode`:
 *
 * - `values` — the value alone, which is what most callers want.
 * - `cells` — the `StyledCell` the writer takes, carrying the value in `v`,
 *   the formula in `f`, and the number format in `s`. Which means what a
 *   reader gives back can be handed straight to a writer.
 */
export type ReadMode = 'values' | 'cells';

/**
 * What each mode gives back per cell.
 *
 * Only `values` carries the date type through. A `StyledCell` is the writer's
 * own shape and its `v` is a `CellValue`, which is already open to every class
 * a workbook was taught — so a cell read as one is a cell the writer takes
 * back whatever the dates in it were read as, which is the whole point of the
 * mode.
 */
export interface ReadModes<D = TemporalDate> {
    values: ReadValue<D>;
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
