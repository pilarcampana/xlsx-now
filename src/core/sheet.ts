import type { WidthMeter } from './autoWidth.js';
import { cellRef, cellXml, columnIndex, columnLetters } from './cell.js';
import { MergeTable } from './merges.js';
import type { StyleRef, StyleTable } from './styles.js';
import type { Cell, CellRow, CellValue, StyledCell } from './types.js';
import { shownWidth, type NativeValue, type ValueTypes } from './valueTypes.js';

const SHEET_PROLOG =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

/**
 * Everything the worksheet carries after its last `<row>`. `<mergeCells>` goes
 * between the two, which is the whole reason merges cost the writer nothing
 * to stream: ECMA-376 §18.3.1.99 puts it after `<sheetData>`, so the ranges
 * are still the sheet's to write once its last row has gone out.
 */
export function sheetFooterXml(merges: MergeTable): string {
    return `</sheetData>${merges.xml()}</worksheet>`;
}

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
 * The style each column gives the cells written in it, by 0-based column. A
 * hole is a column that gives them none.
 *
 * The `<col style>` this comes from is not what makes a cell look like it:
 * that reaches the cells that are *not* in the file, and every cell that is
 * has to carry the style itself. So the layout is read a second time, here,
 * for the rows to be written against — see `StyleTable.stack`.
 */
export function columnStyles(
    formats: ColumnFormats | undefined,
): readonly (StyleRef | undefined)[] {
    // `map` keeps the holes as holes, which is what the sparse array is for.
    return columnLayout(formats, undefined).map((format) => format.s);
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
    if (
        'v' in styled ||
        's' in styled ||
        'f' in styled ||
        't' in styled ||
        'col' in styled ||
        'colSpan' in styled ||
        'rowSpan' in styled
    ) {
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
                'add it to the writer\'s "types", with withType(defaultTypes(), ...).',
        );
    }
    const keys = Object.keys(cell);
    return new Error(
        'A cell is a value, or an object with "v", "s", "f", "t", "col", "colSpan" or "rowSpan": ' +
            (keys.length
                ? `this one has ${keys.map((key) => `"${key}"`).join(', ')}.`
                : 'this one is empty.'),
    );
}

/**
 * How many columns or rows a cell said it takes, as a count. Left out it is
 * `1` — the cell itself, merged with nothing — and anything that is not a
 * whole count of cells is a caller who meant something else.
 */
function spanCount(span: number | undefined, field: 'colSpan' | 'rowSpan'): number {
    if (span === undefined) return 1;
    if (!Number.isInteger(span) || span < 1) {
        throw new Error(
            `"${field}" is how many cells a merge takes, its own included, counted from 1: ` +
                `${JSON.stringify(span)} is not one of them.`,
        );
    }
    return span;
}

/**
 * Whether what the caller left in a column some merge covers can be there.
 *
 * A merged range shows the value of its first cell and nothing else — Excel
 * says as much when it merges, and drops the rest — so a value written under
 * one would go into the file and never be seen again. A hole, a `null`, an
 * empty string: those are the column being left alone, which is what is
 * being asked for. A style there is the columns mode's or the caller's, and
 * it is not what gets written: every cell of a merge carries the style of the
 * cell that declared it.
 */
function fitsUnderMerge(cell: Cell, styled: StyledCell | undefined): boolean {
    const value = styled ? styled.v : (cell as CellValue);
    if (value !== undefined && value !== null && value !== '') return false;
    return (
        styled === undefined ||
        (styled.f === undefined && styled.colSpan === undefined && styled.rowSpan === undefined)
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
 *
 * `columns` is the style each column gives its cells. A cell falls under it,
 * then under the row's own `s`, then under whatever it says for itself, and
 * what gets written on the `<c>` is the three of them stacked: xlsx has no
 * inheritance to lean on, so the cell carries the answer.
 *
 * `merges` is the sheet's own, and it is read as much as it is written to: a
 * `colSpan` or a `rowSpan` goes into it, and what earlier rows put into it is
 * what says which of this row's columns are already taken. The default is
 * there for a row written on its own, with no sheet around it to remember
 * anything.
 */
export function cellRowXml(
    rowNumber: number,
    row: CellRow,
    styles: StyleTable,
    types: ValueTypes,
    options?: RowOptions,
    widths?: WidthMeter,
    columns?: readonly (StyleRef | undefined)[],
    merges: MergeTable = new MergeTable(),
): string {
    let cells = '';
    let next = 0;
    const rowStyle = options?.s;
    // What a merge declared in an earlier row leaves in this one. Ascending
    // by column, and read in step with the row's own cells: the `<c>` of a
    // row go in column order, whoever they came from.
    const covered = merges.openAt(rowNumber);
    let pending = 0;
    // The columns a merge of *this* row took, past the cell that declared it,
    // and the range they belong to. One watermark is enough: a line only
    // moves forward, so an earlier merge of the same row is behind it.
    let takenUntil = 0;
    let takenBy = '';

    /**
     * The cells the merges from above put in this row, up to `column`. They
     * carry the style of the cell that declared the merge — an empty cell
     * under the default style is nothing at all, and writes nothing.
     */
    const coveredBefore = (column: number): string => {
        let xml = '';
        while (pending < covered.length && covered[pending]!.column < column) {
            const span = covered[pending]!;
            xml += cellXml(undefined, cellRef(span.column, rowNumber), span.style);
            pending++;
        }
        return xml;
    };

    /**
     * The merge that has already taken `column`, if one has: this row's own,
     * or one still coming down from an earlier row. Either way the column is
     * not the cell's to write in.
     */
    const takenAt = (column: number): string | undefined => {
        if (column < takenUntil) return takenBy;
        const span = covered[pending];
        return span?.column === column ? span.ref : undefined;
    };

    for (const cell of row) {
        // A hole still takes up its column: the array is the sheet's layout
        // as much as it is the values.
        if (cell === undefined) {
            next++;
            continue;
        }
        const styled = isStyledCell(cell, types) ? cell : undefined;
        const at = styled?.col === undefined ? next : columnAt(styled.col);
        if (styled?.col !== undefined && at < next) {
            throw new Error(
                `Column "${styled.col}" comes before what row ${rowNumber} has already written: ` +
                    'a line fills its columns left to right, once each.',
            );
        }
        cells += coveredBefore(at);
        const taken = takenAt(at);
        if (taken !== undefined) {
            if (!fitsUnderMerge(cell, styled)) {
                throw new Error(
                    `Column ${columnLetters(at)} of row ${rowNumber} is covered by the merge ` +
                        `"${taken}": a merged range shows the value of its first cell, so the ` +
                        'rest of it has to be left empty.',
                );
            }
            // The merge's own cell for this column, when it is one coming
            // from above: the one written here is the caller's hole.
            cells += coveredBefore(at + 1);
            next = at + 1;
            continue;
        }
        const raw = styled ? styled.v : (cell as CellValue);
        const value = types.convert(raw);
        const v = value ? value.v : (raw as NativeValue);
        const colSpan = spanCount(styled?.colSpan, 'colSpan');
        const rowSpan = spanCount(styled?.rowSpan, 'rowSpan');
        const style = styles.forValue(
            value?.numFmt,
            styles.stack(columns?.[at], rowStyle, styled?.s),
        );
        cells += cellXml(
            v,
            cellRef(at, rowNumber),
            style,
            styled?.f,
            // A `t` written on the cell is the caller asking for that one, so
            // it goes over whatever the value's type would have said.
            styled?.t ?? value?.t,
        );
        if (colSpan > 1 || rowSpan > 1) {
            // The next merge still coming down, if it starts inside this one.
            // Two ranges that overlap are a file Excel repairs rather than
            // opens, and it repairs it by dropping things.
            const over = covered[pending];
            if (over !== undefined && over.column < at + colSpan) {
                throw new Error(
                    `The merge starting at ${cellRef(at, rowNumber)} runs over "${over.ref}", ` +
                        'which is still open: merged ranges cannot overlap.',
                );
            }
            takenBy = merges.add(at, rowNumber, colSpan, rowSpan, style);
            takenUntil = at + colSpan;
            // The rest of the range, in this row: empty cells under the same
            // style, which is what draws a border around a merge instead of
            // around its first cell. The rows below get theirs from `covered`.
            for (let column = at + 1; column < takenUntil; column++) {
                cells += cellXml(undefined, cellRef(column, rowNumber), style);
            }
        }
        // A formula's cached result is what the column will have to show; a
        // formula with no result in hand shows nothing until it is recalculated
        // and so measures nothing. A value shown across several columns
        // measures none of them: Excel's own autofit passes merged cells by,
        // and a title stretched over three columns is not how wide the first
        // one has to be. Whether the style wraps is what says how much of a
        // value with line breaks in it the column has to fit.
        if (colSpan === 1) widths?.see(at, v, shownWidth(value), styles.wraps(style));
        next = at + 1;
    }
    // Whatever the merges above left past the row's own last cell.
    cells += coveredBefore(Infinity);
    return `<row r="${rowNumber}"${rowAttributes(options, styles)}>${cells}</row>`;
}
