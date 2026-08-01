import { cellRef, cellXml } from './cell.js';
import { STYLE, type StyleIndex } from './styles.js';
import type { CellValue, Column, Row } from './types.js';

const SHEET_PROLOG =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
export const SHEET_FOOTER = '</sheetData></worksheet>';

/** The worksheet's first row is the header, so data records start here. */
export const FIRST_DATA_ROW = 2;

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

/**
 * Fixes the header row, and the leading pk columns when there are any, so
 * they stay visible while the sheet scrolls. `topLeftCell` is the first cell
 * of the scrolling area and `activePane` names the pane it belongs to:
 * `bottomRight` when both splits are in play, `bottomLeft` with only the row.
 */
function sheetViewsXml(columns: readonly Column[]): string {
    const xSplit = frozenColumnCount(columns);
    const scrollFrom = cellRef(xSplit, FIRST_DATA_ROW);
    const pane = xSplit ? 'bottomRight' : 'bottomLeft';
    return (
        '<sheetViews><sheetView workbookViewId="0">' +
        `<pane${xSplit ? ` xSplit="${xSplit}"` : ''} ySplit="1"` +
        ` topLeftCell="${scrollFrom}" activePane="${pane}" state="frozen"/>` +
        `<selection pane="${pane}" activeCell="${scrollFrom}" sqref="${scrollFrom}"/>` +
        '</sheetView></sheetViews>'
    );
}

/** Everything the worksheet carries before its first `<row>`. */
export function sheetHeaderXml(columns: readonly Column[]): string {
    return SHEET_PROLOG + sheetViewsXml(columns) + '<sheetData>';
}

function styleFor(isHeaderRow: boolean, isPkColumn: boolean): StyleIndex {
    if (isHeaderRow && isPkColumn) return STYLE.PK_HEADER;
    if (isHeaderRow) return STYLE.HEADER;
    if (isPkColumn) return STYLE.PK;
    return STYLE.DEFAULT;
}

function rowXml(
    rowNumber: number,
    values: readonly CellValue[],
    columns: readonly Column[],
    isHeaderRow: boolean,
): string {
    let cells = '';
    for (let i = 0; i < columns.length; i++) {
        const style = styleFor(isHeaderRow, Boolean(columns[i]!.pk));
        cells += cellXml(values[i], cellRef(i, rowNumber), style);
    }
    return `<row r="${rowNumber}">${cells}</row>`;
}

/** Row 1: the column names, in the header style. */
export function headerRowXml(columns: readonly Column[]): string {
    return rowXml(
        1,
        columns.map((c) => c.name),
        columns,
        true,
    );
}

/** One `<row>` for an incoming record, reading each column by its key. */
export function dataRowXml(
    rowNumber: number,
    record: Row,
    columns: readonly Column[],
): string {
    return rowXml(
        rowNumber,
        columns.map((c) => record[c.key ?? c.name]),
        columns,
        false,
    );
}
