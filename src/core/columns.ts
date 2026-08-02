// The columns mode, expressed in terms of the row mode. Nothing here writes
// XML: it produces the freeze the sheet asked for, the header row, and the
// function that turns one incoming record into one row of cells — which is
// all the columns mode ever was.
import type { Freeze } from './sheet.js';
import type { CellRow, Column, Row } from './types.js';

/**
 * How many columns are frozen along with the header row: the primary keys,
 * but only while they are the sheet's first columns. A pk sitting after an
 * ordinary column cannot be frozen on its own — a freeze is a split at one
 * position, so it would drag every column before it along — and there the
 * header row is fixed by itself.
 */
function frozenColumnCount(columns: readonly Column[]): number {
    let leading = 0;
    while (leading < columns.length && columns[leading]!.pk) leading++;
    // All of them pk: freezing every column would leave nothing to scroll.
    return leading === columns.length ? 0 : leading;
}

export interface ColumnsMode {
    /** What the columns imply, and what the explicit options override. */
    freeze: Freeze;
    /** Row 1: the column names, bold, and highlighted where they are pks. */
    headerRow: CellRow;
    /** One record, read by key, as the row array the writer takes. */
    toCellRow(record: Row): CellRow;
}

export function columnsMode(columns: readonly Column[]): ColumnsMode {
    return {
        freeze: { rows: 1, columns: frozenColumnCount(columns) },
        headerRow: columns.map((column) => ({
            value: column.name,
            style: { bold: true, highlight: Boolean(column.pk) },
        })),
        toCellRow: (record) =>
            columns.map((column) => {
                const value = record[column.key ?? column.name];
                return column.pk ? { value, style: { highlight: true } } : value;
            }),
    };
}
