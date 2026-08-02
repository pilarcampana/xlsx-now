// What goes into the writer. One stream carries the whole workbook: rows of
// cells, records read by column, and the commands that open a worksheet or
// spell a line out.
import type { ColumnFormats, RowOptions } from './sheet.js';
import type { CellRow, Column, Row } from './types.js';

/**
 * The keys that make a message a command instead of a row: the worksheet to
 * open, and the line to write.
 *
 * A leading `#` is what marks a command, so a record cannot have keys that
 * start with one — which costs nothing, since a column's `name` (what the
 * header row shows) is free of the restriction and only its `key` is not.
 * A key that starts with `#` and is not one of these is refused by name,
 * rather than going in as a row of blanks.
 */
export const WORKSHEET = '#worksheet';
export const LINE = '#line';

/**
 * Everything that is decided per worksheet. The writer options carry these
 * for the workbook, and a `#worksheet` command carries them again for the
 * sheet it opens — where they override the defaults, one field at a time.
 */
export interface SheetOptions {
    /**
     * Turns on the columns mode for the sheet: it gets a header row of column
     * names, records are read by key, and the freezes below default to the
     * header row and the leading pk columns. Left out — or empty — the sheet
     * has no header of its own and takes rows of cells alone.
     */
    columns?: readonly Column[];
    /**
     * How wide the sheet's columns are, whether they are shown, and what
     * their cells look like without a style of their own. This is the sheet's
     * layout and `columns` is how a record is read, so the two are declared
     * apart: a sheet written from arrays has no `columns` and can still say
     * that its column C is 30 characters wide.
     */
    columnFormats?: ColumnFormats;
    /** Rows fixed at the top of the sheet. Defaults to 0, or to 1 with `columns`. */
    freezeRows?: number;
    /**
     * Columns fixed at the left of the sheet. Defaults to 0, or with
     * `columns` to however many leading ones are pks.
     */
    freezeColumns?: number;
}

/**
 * Closes the worksheet being written and opens a new one under the given
 * name, configured by whatever else the command carries.
 *
 * ```js
 * { '#worksheet': 'Ventas 2024', columns, freezeColumns: 2 }
 * ```
 *
 * Sent before any row, it configures the first worksheet instead of adding a
 * second — which is why `sheetName` and `columns` are optional in the writer
 * options: a stream whose first message is a command declares everything it
 * needs on its way in.
 */
export type WorksheetCommand = SheetOptions & { [WORKSHEET]: string };

/**
 * One line, said outright instead of left to be recognized:
 *
 * ```js
 * { '#line': 'row', values: { id: 1, name: 'Ana' } }   // read by the columns
 * { '#line': 'array', values: [1, 'Ana'] }             // position is the column
 * { '#line': 'empty' }                                 // a row and nothing in it
 * ```
 *
 * The point of saying it outright is `RowOptions`: `height`, `hidden` and an
 * `s` for the whole row have nowhere to go on a bare array or record.
 *
 * ```js
 * { '#line': 'array', values: ['Total'], s: { bold: true }, height: 22 }
 * ```
 *
 * A line that only touches a few far-apart columns needs no form of its own:
 * a cell says which column it goes in.
 *
 * ```js
 * { '#line': 'array', values: [{ v: 'total', col: 'A' }, { v: 12, col: 'F' }] }
 * ```
 */
export type LineCommand = RowOptions &
    (
        | { [LINE]: 'row'; values: Row }
        | { [LINE]: 'array'; values: CellRow }
        | { [LINE]: 'empty'; values?: undefined }
    );

/** One message on the way in: a row of cells, a record, or a command. */
export type SheetInput = CellRow | Row | WorksheetCommand | LineCommand;

export function isWorksheetCommand(input: SheetInput): input is WorksheetCommand {
    return !Array.isArray(input) && WORKSHEET in input;
}

export function isLineCommand(input: SheetInput): input is LineCommand {
    return !Array.isArray(input) && LINE in input;
}

/**
 * A record is what an object message is when it claims no command — but a key
 * that starts with `#` is a command nobody knows, and a misspelled one
 * (`#worksheets`, `#lines`) would otherwise go in as a row of blanks. Reading
 * the keys is what the columns mode does with the record anyway.
 */
export function checkRecord(record: Row): void {
    for (const key in record) {
        if (key.charCodeAt(0) === 35) {
            throw new Error(
                `Unknown command "${key}": the commands are "${WORKSHEET}" and "${LINE}".`,
            );
        }
    }
}

/** Why a record could not be written on a sheet that has no columns. */
export function noColumnsError(): Error {
    return new Error(
        'A record needs columns to be read by: declare them in the writer options, ' +
            `or in a "${WORKSHEET}" command. Rows of cells need no columns.`,
    );
}

/**
 * The cells a `#line` command spells out. `values` is taken as the kind says,
 * and a kind the writer does not know is refused by name — with `row` left to
 * the caller, which is the one form that needs the sheet's columns.
 */
export function lineCells(command: LineCommand): CellRow | undefined {
    switch (command[LINE]) {
        case 'empty':
            return [];
        case 'array':
            // Nothing to convert: `values` is already the row it will be
            // written as. Absent, the line is as empty as `empty` says.
            return command.values ?? [];
        case 'row':
            // The record still has to be read by the sheet's columns, which
            // this module knows nothing about.
            return undefined;
        default:
            throw new Error(
                `Unknown line "${String((command as Record<string, unknown>)[LINE])}": ` +
                    'a line is "row", "array" or "empty".',
            );
    }
}

/** The record a `{ '#line': 'row' }` command carries, checked like a bare one. */
export function lineRecord(command: LineCommand): Row {
    const values = (command as { values?: Row }).values ?? {};
    checkRecord(values);
    return values;
}
