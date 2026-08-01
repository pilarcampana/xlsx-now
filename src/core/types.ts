/** Anything a caller can put in a cell. Everything else is rendered as text. */
export type CellValue = string | number | boolean | Date | null | undefined;

/** One incoming record: the shape a `columns[].key` is looked up on. */
export type Row = Record<string, CellValue>;

export interface Column {
    /** Text written in the header row. */
    name: string;
    /** Property read from each record; defaults to `name`. */
    key?: string;
    /** Marks the column as a primary key, so it gets the highlight fill. */
    pk?: boolean;
}

/** Accepts sync and async sources alike, so `rows` can be either. */
export type ForAwaitable<T> = AsyncIterable<T> | Iterable<T>;
