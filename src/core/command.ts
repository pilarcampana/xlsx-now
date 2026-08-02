// What goes into the writer, beyond the rows themselves. One stream carries
// the whole workbook: rows of cells, records read by column, and the commands
// that open a new worksheet along the way.
import type { CellRow, Column, Row } from './types.js';

/**
 * The key that makes a message a command instead of a row, and the only
 * command there is: its value is the name of the worksheet to open.
 *
 * A leading `#` is what marks a command, so a record cannot have keys that
 * start with one — which costs nothing, since a column's `name` (what the
 * header row shows) is free of the restriction and only its `key` is not.
 */
export const WORKSHEET = '#worksheet';

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

/** One message on the way in: a row of cells, a record, or a command. */
export type SheetInput = CellRow | Row | WorksheetCommand;

/**
 * A command is an object carrying the `#worksheet` key — which a row array
 * cannot be, and a record must not be.
 */
export function isWorksheetCommand(input: SheetInput): input is WorksheetCommand {
    return !Array.isArray(input) && WORKSHEET in input;
}

/**
 * Why a record could not be written: either the sheet has no columns to read
 * it by, or what looked like a record is a command nobody knows — a
 * misspelled `#worksheet` would otherwise go in as a blank row.
 */
export function recordError(record: Row): Error {
    for (const key in record) {
        if (key.startsWith('#')) {
            return new Error(`Unknown command "${key}": the only one is "${WORKSHEET}".`);
        }
    }
    return new Error(
        'A record needs columns to be read by: declare them in the writer options, ' +
            `or in a "${WORKSHEET}" command. Rows of cells need no columns.`,
    );
}
