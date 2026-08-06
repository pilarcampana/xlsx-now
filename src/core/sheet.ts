import type { WidthMeter } from './autoWidth.js';
import { cellRef, cellXml, columnIndex } from './cell.js';
import type { StyleRef, StyleTable } from './styles.js';
import type { Cell, CellRow, StyledCell } from './types.js';
import { shownWidth, type NativeValue, type ValueTypes } from './valueTypes.js';

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
    /**
     * Width as Excel measures it: the number its own column-width dialog
     * shows, which already carries the padding a column has around its text.
     * It goes into the `<col>` as it is given — it is a width, not a count of
     * characters to be worked into one.
     */
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

/**
 * The columns of the sheet as one array, by 0-based column — the layout, and
 * the only place it is held: the formats as they were given, whether by
 * position or by the column each one is for, with the width a column measured
 * for itself filled in wherever the sheet did not say one outright.
 *
 * The array *is* the order a `<cols>` runs in, which is why nothing here has
 * to be sorted afterwards: a record's keys arrive in whatever order they were
 * written in and land in the column they name. A hole is a column that asked
 * for nothing.
 */
function columnLayout(
    formats: ColumnFormats | undefined,
    autoWidths: readonly number[] | undefined,
): ColumnFormat[] {
    // Sparse on purpose: the holes are the columns that asked for nothing, and
    // both `forEach` and a `for...of` over the entries pass them by.
    const layout: ColumnFormat[] = [];
    if (Array.isArray(formats)) {
        (formats as readonly (ColumnFormat | undefined | null)[]).forEach((format, index) => {
            if (format) layout[index] = format;
        });
    } else if (formats) {
        for (const [key, format] of Object.entries(formats as Record<string, ColumnFormat>)) {
            if (format) layout[columnAt(key)] = format;
        }
    }
    // A measured width is what fills in for a column that said nothing about
    // how wide it is; one said outright is the width, and is not measured over.
    autoWidths?.forEach((width, index) => {
        const format = layout[index];
        if (format?.width !== undefined) return;
        layout[index] = format ? { ...format, width } : { width };
    });
    return layout;
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
    autoWidths?: readonly number[],
): string {
    let cols = '';
    columnLayout(formats, autoWidths).forEach((format, index) => {
        const at = index + 1;
        cols +=
            `<col min="${at}" max="${at}"` +
            (format.width === undefined ? '' : ` width="${format.width}" customWidth="1"`) +
            (format.s === undefined ? '' : ` style="${styles.index(format.s)}"`) +
            (format.hidden ? ' hidden="1"' : '') +
            '/>';
    });
    // An empty `<cols/>` is not a sheet with no column formats: it is a sheet
    // Excel refuses to open.
    return cols ? `<cols>${cols}</cols>` : '';
}

/**
 * Everything the worksheet carries before its first `<row>`. `autoWidths` is
 * what the sheet's own cells measured, when it was written with an
 * `autoWidthMax` — which is why this header is not something the writer can
 * hand over before the rows: `<cols>` goes ahead of `<sheetData>`, and those
 * widths are not known until the last row of the sheet is in.
 */
export function sheetHeaderXml(
    freeze: Freeze,
    styles: StyleTable,
    columnFormats?: ColumnFormats,
    autoWidths?: readonly number[],
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
 * Whether an object is a cell that says more about itself, or a value.
 *
 * The order is what this is: an object of a type the workbook knows is a
 * value, whatever it looks like — a `Date` has no `v` and never meant to. Only
 * past that does an object get read as a cell, and only past *that* does the
 * `Object` entry get to claim what is left. Anything still unaccounted for is
 * a caller who meant something the writer cannot guess, and letting it through
 * as a blank would hide it.
 */
export function isStyledCell(cell: Cell, types: ValueTypes): cell is StyledCell {
    if (typeof cell !== 'object' || cell === null) return false;
    if (types.handlerFor(cell) !== undefined) return false;
    const styled = cell as StyledCell;
    if ('v' in styled || 's' in styled || 'f' in styled || 't' in styled || 'col' in styled) {
        return true;
    }
    if (types.objectHandler !== undefined) return false;
    throw unknownCellError(cell);
}

/**
 * Why an object was neither a value nor a cell.
 *
 * The two cases read differently and are worth saying differently. A plain
 * object is a cell that was spelled wrong — the fields it should have had are
 * the thing to name. An instance of a class is a type nobody registered, and
 * what the caller needs to hear is which class and where to say so.
 */
function unknownCellError(cell: object): Error {
    const prototype: unknown = Object.getPrototypeOf(cell);
    if (prototype !== Object.prototype && prototype !== null) {
        const { constructor } = prototype as { constructor?: { name?: string } };
        const named = constructor?.name ? `"${constructor.name}"` : 'this one';
        return new Error(
            `A cell is a value of a type the workbook knows, and ${named} is not one of them: ` +
                'add it to the writer\'s "types", with withType(defaultTypes, ...).',
        );
    }
    const keys = Object.keys(cell);
    return new Error(
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
 *
 * `types` is what a value that is not already a native one becomes. It is
 * asked once per cell, here, and what it answers is what the three things
 * downstream are handed — the XML, the style, and the width — so no two of
 * them can disagree about what the cell holds.
 */
export function cellRowXml(
    rowNumber: number,
    row: CellRow,
    styles: StyleTable,
    types: ValueTypes,
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
        if (!isStyledCell(cell, types)) {
            const value = types.convert(cell);
            const v = value ? value.v : (cell as NativeValue);
            cells += cellXml(
                v,
                cellRef(next, rowNumber),
                styles.forValue(value?.numFmt, undefined),
                undefined,
                value?.t,
            );
            widths?.see(next, v, shownWidth(value));
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
        const value = types.convert(cell.v);
        const v = value ? value.v : (cell.v as NativeValue);
        cells += cellXml(
            v,
            cellRef(at, rowNumber),
            styles.forValue(value?.numFmt, cell.s),
            cell.f,
            // A `t` written on the cell is the caller asking for that one, so
            // it goes over whatever the value's type would have said.
            cell.t ?? value?.t,
        );
        // A formula's cached result is what the column will have to show; a
        // formula with no result in hand shows nothing until it is recalculated
        // and so measures nothing.
        widths?.see(at, v, shownWidth(value));
        next = at + 1;
    }
    return `<row r="${rowNumber}"${rowAttributes(options, styles)}>${cells}</row>`;
}
