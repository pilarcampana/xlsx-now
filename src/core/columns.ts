// The columns mode, expressed in terms of the row mode. Nothing here writes
// XML: it produces the freeze the sheet asked for, the header row, and the
// function that turns one incoming record into one row of cells — which is
// all the columns mode ever was.
import { isStyledCell, type Freeze } from './sheet.js';
import type { StyleSpec } from './styles.js';
import type { CellRow, Column, Row } from './types.js';
import type { ValueTypes } from './valueTypes.js';

/**
 * What the columns mode looks like. These are held as constants rather than
 * built per row on purpose: the style table recognizes a spec it has already
 * been handed by identity, so a whole sheet of pk cells costs one lookup each
 * and not one rendering each.
 */
const PK_FILL = '#FFE699';
const HEADER: StyleSpec = { bold: true };
const HEADER_PK: StyleSpec = { bold: true, bg: PK_FILL };
const PK: StyleSpec = { bg: PK_FILL };

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

export function columnsMode(columns: readonly Column[], types: ValueTypes): ColumnsMode {
    return {
        freeze: { rows: 1, columns: frozenColumnCount(columns) },
        headerRow: columns.map((column) => ({
            v: column.name,
            s: column.pk ? HEADER_PK : HEADER,
        })),
        toCellRow: (record) =>
            columns.map((column) => {
                const value = record[column.key ?? column.name];
                if (!column.pk) return value;
                // The pk fill is what the column asks for; a cell that says
                // how it looks has said it for itself.
                return isStyledCell(value, types)
                    ? { ...value, s: value.s ?? PK }
                    : { v: value, s: PK };
            }),
    };
}
