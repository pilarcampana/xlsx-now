import type { StyleRef } from './styles.js';

/**
 * Anything a caller can put in a cell.
 *
 * A string, a number and a boolean go in as they are. Everything else is an
 * instance of a type the workbook knows how to write — a `Date`, a `URL`, one
 * of the caller's own classes — and what makes it known is the `types` map;
 * `object` is as narrow as this can be said, since the classes in that map are
 * the caller's to choose. An object of a type nobody registered is refused by
 * name rather than written out as whatever `String()` makes of it.
 */
export type CellValue = string | number | bigint | boolean | object | null | undefined;

/**
 * One incoming record: the shape a `columns[].key` is looked up on. A
 * property is a cell like any other, so a record can style one of its values
 * without giving up on being read by the sheet's columns.
 */
export type Row = Record<string, Cell>;

/**
 * What a `<c>` says it holds, as xlsx spells it: a number, a boolean, the
 * string result of a formula, a string written into the cell itself, or an
 * error. Left out, it is read off the value — which is what makes it worth
 * writing: `{ v: '007', t: 'inlineStr' }` keeps a code that looks like a
 * number from being shown as `7`.
 *
 * `str` and `inlineStr` are both text, and which of the two a cell is written
 * as is decided by the cell, not by the caller: `str` is the *formula string*
 * of the spec — the cached result of an `f` — and `inlineStr` is text the cell
 * holds itself, in an `<is>` there is no room for next to a formula. So text
 * asked for either way goes in as the one the cell has a place for.
 *
 * There is no `s` here — the shared string table — because this writer has
 * none: a string goes into the cell it belongs to and nowhere else, which is
 * what lets a sheet be written without holding on to anything.
 */
export type CellType = 'n' | 'b' | 'str' | 'inlineStr' | 'e';

/**
 * A cell that says more than its value.
 *
 * ```js
 * { v: 1234.5, s: 'money' }                    // a declared style, by name
 * { v: new Date(), s: { numFmt: 'dd/mm/yy' } } // a style written out
 * { v: 45, f: 'SUM(B2:B10)' }                  // a formula, and its result
 * { v: 'Total', col: 'J' }                     // in column J, not the next one
 * ```
 *
 * Every field is optional, the value included: `{ s: 'header' }` is an empty
 * cell that still carries its style, and `{ f: 'NOW()' }` is a formula whose
 * result the reader works out for itself.
 */
export interface StyledCell {
    /** The value. */
    v?: CellValue;
    /** The style: the name of a declared one, or one written out. */
    s?: StyleRef;
    /**
     * The formula, with or without a leading `=`. With `v` next to it, that
     * value is written as the cached result — which is what a reader shows
     * until it recalculates, and what one that never does shows for good.
     */
    f?: string;
    /** What the cell holds. Read off `v` when it is not said. */
    t?: CellType;
    /**
     * The column to write this cell in: `'J'`, or `10` for the same one —
     * columns are numbered from 1, as the sheet shows them. Whatever follows
     * carries on from there, so it is how a line reaches a far column without
     * counting out the holes in between. A line only moves forward: a column
     * it has already gone past is an error, not a cell written twice.
     */
    col?: string | number;
    /**
     * How many columns the cell takes, its own included: `3` merges it with
     * the two to its right, and `1` — the default — merges nothing.
     *
     * A merged range shows the value of its first cell and nothing else, so
     * the columns it covers are not the row's to write in: leave a hole
     * there, or reach past the merge with the next cell's `col`.
     *
     * ```js
     * [{ v: 'Ventas 2024', colSpan: 3 }, undefined, undefined, 'Total']
     * [{ v: 'Ventas 2024', colSpan: 3 }, { v: 'Total', col: 'D' }]
     * ```
     */
    colSpan?: number;
    /**
     * How many rows the cell takes, its own included: `3` merges it with the
     * two below it, and `1` — the default — merges nothing. Together with
     * `colSpan` it is a block, as Excel merges one.
     *
     * The rows underneath have to leave the covered columns empty, the same
     * way the rest of the row does — the sheet is a grid, and a merge takes
     * its place in it rather than pushing anything along.
     */
    rowSpan?: number;
}

/** One position of a row array: a bare value, or a cell that says more. */
export type Cell = CellValue | StyledCell;

/**
 * One row as an array: the position *is* the column, unless a cell says
 * otherwise with its `col`. An `undefined` position writes no cell at all —
 * as opposed to a `null` or an empty string, which are an empty cell.
 */
export type CellRow = readonly Cell[];

export interface Column {
    /** Text written in the header row. */
    name: string;
    /** Property read from each record; defaults to `name`. */
    key?: string;
    /**
     * Marks the column as a primary key, so it gets the highlight fill — and,
     * while the pks are the sheet's first columns, so it stays frozen next to
     * the header row.
     */
    pk?: boolean;
}

/** Accepts sync and async sources alike, so `rows` can be either. */
export type ForAwaitable<T> = AsyncIterable<T> | Iterable<T>;
