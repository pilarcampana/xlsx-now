import { cellRef, cellXml } from './cell.js';
import { styleIndex } from './styles.js';
import type { Cell, CellRow, CellStyle, CellValue } from './types.js';

const SHEET_PROLOG =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
export const SHEET_FOOTER = '</sheetData></worksheet>';

/** How much of the sheet stays put while the rest of it scrolls. */
export interface Freeze {
    /** Rows fixed at the top. */
    rows: number;
    /** Columns fixed at the left. */
    columns: number;
}

/**
 * Fixes the first `rows` rows and the first `columns` columns, so they stay
 * visible while the sheet scrolls. `topLeftCell` is the first cell of the
 * scrolling area and `activePane` names the pane it belongs to: `bottomRight`
 * when both splits are in play, `bottomLeft` or `topRight` with only one.
 *
 * A freeze is a single split at one position, and it always takes everything
 * before it along — which is why this is two counts and not a choice of which
 * rows or columns to fix.
 */
function sheetViewsXml({ rows, columns }: Freeze): string {
    // The view itself is always written, frozen or not: Excel's own files
    // carry it, and a reader that goes looking for it finds nothing to read
    // when it is missing.
    if (!rows && !columns) return '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
    const scrollFrom = cellRef(columns, rows + 1);
    const pane = rows ? (columns ? 'bottomRight' : 'bottomLeft') : 'topRight';
    return (
        '<sheetViews><sheetView workbookViewId="0">' +
        `<pane${columns ? ` xSplit="${columns}"` : ''}${rows ? ` ySplit="${rows}"` : ''}` +
        ` topLeftCell="${scrollFrom}" activePane="${pane}" state="frozen"/>` +
        `<selection pane="${pane}" activeCell="${scrollFrom}" sqref="${scrollFrom}"/>` +
        '</sheetView></sheetViews>'
    );
}

/** Everything the worksheet carries before its first `<row>`. */
export function sheetHeaderXml(freeze: Freeze): string {
    return SHEET_PROLOG + sheetViewsXml(freeze) + '<sheetData>';
}

/**
 * What a line can ask for beyond its cells. There is nowhere to put these on
 * a bare row — an array is all cells, a record is all values — which is what
 * the `#line` command is for.
 */
export interface RowOptions {
    /** Height in points. Excel's own default is 15, and it is what applies without this. */
    height?: number;
    /** Keeps the row in the sheet but out of sight. */
    hidden?: boolean;
    /** Applies to the whole row, under whatever style its cells carry themselves. */
    style?: CellStyle;
}

/** The attributes of a `<row>`, past its number. */
function rowAttributes(options: RowOptions | undefined): string {
    if (!options) return '';
    let attributes = '';
    // `customFormat` and `customHeight` are what tell Excel the row means it:
    // without them the `s` and the `ht` next to them are ignored.
    if (options.style) attributes += ` s="${styleIndex(options.style)}" customFormat="1"`;
    if (options.height !== undefined) attributes += ` ht="${options.height}" customHeight="1"`;
    if (options.hidden) attributes += ' hidden="1"';
    return attributes;
}

/**
 * A `Date` is a value, not a wrapper: it is the one object a cell can be
 * without asking for a style.
 */
function isStyled(cell: Cell): cell is { value: CellValue; style?: CellStyle } {
    return typeof cell === 'object' && cell !== null && !(cell instanceof Date);
}

/**
 * One `<row>` out of one row array. An `undefined` position writes no cell at
 * all, which is how a row leaves a column untouched; a `null` or an empty
 * string is an empty cell, and is written whenever it carries a style. An
 * explicit `{ value: undefined, style }` writes the styled cell too — the
 * wrapper is the caller asking for it.
 */
export function cellRowXml(rowNumber: number, row: CellRow, options?: RowOptions): string {
    let cells = '';
    for (let i = 0; i < row.length; i++) {
        const cell = row[i];
        if (cell === undefined) continue;
        const ref = cellRef(i, rowNumber);
        cells += isStyled(cell)
            ? cellXml(cell.value, ref, styleIndex(cell.style))
            : cellXml(cell, ref, styleIndex(undefined));
    }
    return `<row r="${rowNumber}"${rowAttributes(options)}>${cells}</row>`;
}
