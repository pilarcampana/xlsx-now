/** Anything a caller can put in a cell. Everything else is rendered as text. */
export type CellValue = string | number | boolean | Date | null | undefined;

/** One incoming record: the shape a `columns[].key` is looked up on. */
export type Row = Record<string, CellValue>;

/**
 * What a cell can be styled with. A closed set on purpose: `styles.xml` is
 * written before the first row arrives, so every combination a cell may ask
 * for has to be in the sheet's style table from the start. Two flags means
 * four combinations, which are exactly the four entries it already carries.
 */
export interface CellStyle {
    /** Bold font — what the header row uses. */
    bold?: boolean;
    /** The highlight fill — what the pk columns use. */
    highlight?: boolean;
}

/** One position of a row array: a bare value, or a value with a style. */
export type Cell = CellValue | { value: CellValue; style?: CellStyle };

/**
 * One row as an array: the position *is* the column, so there is nothing to
 * declare. An `undefined` position writes no cell at all — as opposed to a
 * `null` or an empty string, which are an empty cell.
 */
export type CellRow = readonly Cell[];

export interface Column {
    /** Text written in the header row. */
    name: string;
    /** Property read from each record; defaults to `name`. */
    key?: string;
    /**
     * Marks the column as a primary key, so it gets the highlight fill — and,
     * while the pks are the sheet's first columns, so it stays frozen next to
     * the header row.
     */
    pk?: boolean;
}

/** Accepts sync and async sources alike, so `rows` can be either. */
export type ForAwaitable<T> = AsyncIterable<T> | Iterable<T>;
