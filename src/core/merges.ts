// The merged ranges of one sheet: what the `colSpan` and `rowSpan` of its
// cells add up to. Nothing here writes a `<c>` — it answers which columns of
// the row being written belong to a merge that started above it, and hands
// back the `<mergeCells>` the worksheet closes with.
//
// That element is why merges cost the writer nothing to stream: ECMA-376 puts
// `<mergeCells>` *after* `</sheetData>`, so the ranges are collected as the
// rows go by and written at the end, next to the footer. The one thing that
// grows is the list of ranges itself — one short string per merge, and a
// sheet that merges something on every row is not what merges are for.
import { cellRef } from './cell.js';

/**
 * A merge still open below the row that declared it: one column of it, and
 * the rows it goes on covering.
 *
 * A range is one entry per column, not one per range, because that is how the
 * rows underneath read it: what a row asks is whether *this* column is taken,
 * and by what.
 */
export interface MergeSpan {
    /** The 0-based column it covers. */
    column: number;
    /** Its last row, 1-based: the row after which it covers nothing. */
    through: number;
    /**
     * The style index the cell that declared it was written with. Every cell
     * of a merge carries it, which is what makes a border go around the whole
     * of the range instead of around its first cell.
     */
    style: number;
    /** The range as the file spells it, `"A3:A5"` — what an error names. */
    ref: string;
}

/** What a row with nothing above it gets, so the common case allocates none. */
const NO_SPANS: readonly MergeSpan[] = [];

/**
 * The merges of one worksheet, filled in as its cells declare them. One per
 * sheet: a `#worksheet` command starts another, and what the last one had
 * open is not the new one's business.
 */
export class MergeTable {
    /** Every range declared so far, in the order the cells declared them. */
    private readonly refs: string[] = [];
    /**
     * The vertical spans that have not run out yet, ascending by column —
     * kept sorted as they arrive, since that is the order a row reads them in
     * and the rows are what this is for.
     */
    private open: MergeSpan[] = [];

    /** Keeps `open` ascending by column, whatever order the rows opened them in. */
    private insert(span: MergeSpan): void {
        let at = this.open.length;
        while (at > 0 && this.open[at - 1]!.column > span.column) at--;
        this.open.splice(at, 0, span);
    }

    /**
     * The spans covering `rowNumber`, ascending by column, as a list of its
     * own: a cell of the row being written may open one more, and what the
     * row is reading cannot grow under it.
     *
     * Rows arrive in order, so this is also where a span that ran out above
     * is let go of.
     */
    openAt(rowNumber: number): readonly MergeSpan[] {
        if (!this.open.length) return NO_SPANS;
        this.open = this.open.filter((span) => span.through >= rowNumber);
        return this.open.length ? this.open.slice() : NO_SPANS;
    }

    /**
     * Declares the merge a cell asked for and hands back the range it became.
     * `column` and `rowNumber` are where the cell itself is: the range starts
     * there and reaches `colSpan` columns to the right and `rowSpan` rows
     * down, the cell's own included.
     *
     * Only a merge that outlives its row is remembered here past the range:
     * one that stays inside it is over by the time the next row starts.
     */
    add(column: number, rowNumber: number, colSpan: number, rowSpan: number, style: number): string {
        const through = rowNumber + rowSpan - 1;
        const ref = `${cellRef(column, rowNumber)}:${cellRef(column + colSpan - 1, through)}`;
        this.refs.push(ref);
        if (rowSpan > 1) {
            for (let at = column; at < column + colSpan; at++) {
                this.insert({ column: at, through, style, ref });
            }
        }
        return ref;
    }

    /**
     * A merge that still expects a row `rowNumber` — which, asked with the row
     * the sheet would have written next, is a merge running past the end of
     * it. `undefined` when every range the sheet declared is closed.
     */
    unfinishedAt(rowNumber: number): MergeSpan | undefined {
        return this.open.find((span) => span.through >= rowNumber);
    }

    /**
     * The `<mergeCells>` of the worksheet, or nothing at all when the sheet
     * merged nothing: an empty one is not a sheet without merges, it is a
     * sheet Excel refuses to open — the same way an empty `<cols/>` is.
     */
    xml(): string {
        if (!this.refs.length) return '';
        const merged = this.refs.map((ref) => `<mergeCell ref="${ref}"/>`).join('');
        return `<mergeCells count="${this.refs.length}">${merged}</mergeCells>`;
    }
}
