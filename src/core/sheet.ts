import type { WidthMeter } from './autoWidth.js';
import { cellRef, cellXml, columnIndex } from './cell.js';
import type { StyleRef, StyleTable } from './styles.js';
import type { Cell, CellRow, StyledCell } from './types.js';

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
 * A column, however it was named: `'J'`, `'j'`, or `10` — columns are
 * numbered from 1, as the sheet shows them. Comes back 0-based, which is how
 * a row array counts.
 */
function columnAt(col: string | number): number {
    const asNumber = typeof col === 'number' ? col : /^\d+$/.test(col) ? Number(col) : undefined;
    if (asNumber === undefined) {
        const index = columnIndex(col as string);
        if (index !== undefined) return index;
    } else if (Number.isInteger(asNumber) && asNumber >= 1) {
        return asNumber - 1;
    }
    throw new Error(
        `"${col}" is not a column: name it by letter ("A", "B", "AA") or by number from 1.`,
    );
}

/**
 * What a column is given beyond the cells written in it: how wide it is,
 * whether it is shown at all, and the style its cells fall back to.
 *
 * This is the sheet's own layout, and it has nothing to do with the `columns`
 * that read a record: a sheet written from arrays can still say how wide its
 * third column is. The style applies to every cell of the column that carries
 * none of its own, and the row's style — and then the cell's — goes over it.
 */
export interface ColumnFormat {
    /** Width in characters, as Excel measures it. */
    width?: number;
    /** Keeps the column in the sheet but out of sight. */
    hidden?: boolean;
    /** What the column's cells look like without a style of their own. */
    s?: StyleRef;
}

/**
 * The columns of a sheet, laid out: by position, or by the column each one is
 * for.
 *
 * ```js
 * [{ width: 8 }, undefined, { width: 30, s: 'money' }]  // A and C
 * { A: { width: 8 }, C: { width: 30, s: 'money' } }     // the same two
 * ```
 */
export type ColumnFormats =
    | readonly (ColumnFormat | undefined | null)[]
    | Readonly<Record<string, ColumnFormat>>;

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

/** The formats as pairs of `[0-based column, what it asks for]`, in order. */
function columnFormatPairs(formats: ColumnFormats): [number, ColumnFormat][] {
    const pairs: [number, ColumnFormat][] = [];
    if (Array.isArray(formats)) {
        (formats as readonly (ColumnFormat | undefined | null)[]).forEach((format, index) => {
            if (format) pairs.push([index, format]);
        });
    } else {
        for (const [key, format] of Object.entries(formats as Record<string, ColumnFormat>)) {
            if (format) pairs.push([columnAt(key), format]);
        }
    }
    // A `<cols>` runs left to right, and a record's keys are in whatever
    // order they happen to have been written in.
    return pairs.sort(([a], [b]) => a - b);
}

/**
 * The columns to write out: what the formats say about each one, plus a width
 * for every column that measured one and was not given a width outright — a
 * width said in `columnFormats` is the width, and the measuring is what fills
 * in for the columns that said nothing.
 */
function columnLayout(
    formats: ColumnFormats | undefined,
    autoWidths: readonly (number | undefined)[] | undefined,
): [number, ColumnFormat][] {
    const layout = new Map<number, ColumnFormat>(formats ? columnFormatPairs(formats) : []);
    autoWidths?.forEach((width, index) => {
        if (width === undefined) return;
        const format = layout.get(index);
        if (format?.width !== undefined) return;
        layout.set(index, format ? { ...format, width } : { width });
    });
    // A `<cols>` runs left to right, and the two sources are in no order
    // between them.
    return [...layout].sort(([a], [b]) => a - b);
}

/**
 * The `<cols>` of a worksheet, which is where a column's width, its style and
 * whether it is shown at all are kept — one `<col>` per column, spanning the
 * one it is for and no more. `customWidth` is what makes Excel apply the
 * width, the same way `customHeight` does for a row.
 */
function colsXml(
    formats: ColumnFormats | undefined,
    styles: StyleTable,
    autoWidths?: readonly (number | undefined)[],
): string {
    const cols = columnLayout(formats, autoWidths)
        .map(([index, format]) => {
            const at = index + 1;
            return (
                `<col min="${at}" max="${at}"` +
                (format.width === undefined ? '' : ` width="${format.width}" customWidth="1"`) +
                (format.s === undefined ? '' : ` style="${styles.index(format.s)}"`) +
                (format.hidden ? ' hidden="1"' : '') +
                '/>'
            );
        })
        .join('');
    // An empty `<cols/>` is not a sheet with no column formats: it is a sheet
    // Excel refuses to open.
    return cols ? `<cols>${cols}</cols>` : '';
}

/**
 * Everything the worksheet carries before its first `<row>`. `autoWidths` is
 * what the sheet's own cells measured, when it was written with an
 * `autoWidthMax` — which is why a header is not always something the writer
 * can hand over before the rows: `<cols>` goes ahead of `<sheetData>`, and
 * those widths are not known until the last row of the sheet is in.
 */
export function sheetHeaderXml(
    freeze: Freeze,
    styles: StyleTable,
    columnFormats?: ColumnFormats,
    autoWidths?: readonly (number | undefined)[],
): string {
    const cols = columnFormats || autoWidths ? colsXml(columnFormats, styles, autoWidths) : '';
    return SHEET_PROLOG + sheetViewsXml(freeze) + cols + '<sheetData>';
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
    s?: StyleRef;
}

/** The attributes of a `<row>`, past its number. */
function rowAttributes(options: RowOptions | undefined, styles: StyleTable): string {
    if (!options) return '';
    let attributes = '';
    // `customFormat` and `customHeight` are what tell Excel the row means it:
    // without them the `s` and the `ht` next to them are ignored.
    if (options.s !== undefined) attributes += ` s="${styles.index(options.s)}" customFormat="1"`;
    if (options.height !== undefined) attributes += ` ht="${options.height}" customHeight="1"`;
    if (options.hidden) attributes += ' hidden="1"';
    return attributes;
}

/**
 * A `Date` is a value, not a cell that says more about itself: it is the one
 * object a cell can be on its own. Anything else has to be recognizable as a
 * cell — an object with none of the fields is a caller who meant something
 * the writer cannot guess, and letting it through as a blank would hide it.
 */
export function isStyledCell(cell: Cell): cell is StyledCell {
    if (typeof cell !== 'object' || cell === null || cell instanceof Date) return false;
    const styled = cell as StyledCell;
    if ('v' in styled || 's' in styled || 'f' in styled || 't' in styled || 'col' in styled) {
        return true;
    }
    const keys = Object.keys(styled);
    throw new Error(
        'A cell is a value, or an object with "v", "s", "f", "t" or "col": ' +
            (keys.length
                ? `this one has ${keys.map((key) => `"${key}"`).join(', ')}.`
                : 'this one is empty.'),
    );
}

/**
 * One `<row>` out of one row array.
 *
 * The position in the array is the column, until a cell names one with its
 * `col`; from there whatever follows carries on. A line only moves forward —
 * a `col` pointing at a column already written, or already gone past, is
 * refused rather than written twice and left for Excel to sort out.
 *
 * An `undefined` position writes no cell at all, which is how a row leaves a
 * column untouched; a `null` or an empty string is an empty cell, and is
 * written whenever it carries a style. An explicit `{ s }` with no value
 * writes the styled cell too — asking for the style is asking for the cell.
 *
 * `widths` is the sheet's meter, when it is sizing its columns by what they
 * hold: every value goes through it on its way out, in the column it turned
 * out to be written in — which is why this is where the measuring happens and
 * not a pass of its own. Left out, nothing is measured.
 */
export function cellRowXml(
    rowNumber: number,
    row: CellRow,
    styles: StyleTable,
    options?: RowOptions,
    widths?: WidthMeter,
): string {
    let cells = '';
    let next = 0;
    for (const cell of row) {
        // A hole still takes up its column: the array is the sheet's layout
        // as much as it is the values.
        if (cell === undefined) {
            next++;
            continue;
        }
        if (!isStyledCell(cell)) {
            cells += cellXml(cell, cellRef(next, rowNumber), styles.forValue(cell, undefined));
            widths?.see(next, cell);
            next++;
            continue;
        }
        const at = cell.col === undefined ? next : columnAt(cell.col);
        if (at < next) {
            throw new Error(
                `Column "${cell.col}" comes before what row ${rowNumber} has already written: ` +
                    'a line fills its columns left to right, once each.',
            );
        }
        cells += cellXml(
            cell.v,
            cellRef(at, rowNumber),
            styles.forValue(cell.v, cell.s),
            cell.f,
            cell.t,
        );
        // A formula's cached result is what the column will have to show; a
        // formula with no result in hand shows nothing until it is recalculated
        // and so measures nothing.
        widths?.see(at, cell.v);
        next = at + 1;
    }
    return `<row r="${rowNumber}"${rowAttributes(options, styles)}>${cells}</row>`;
}
