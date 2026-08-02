import type { SheetInput } from './command.js';
import type { ForAwaitable } from './types.js';
import { XlsxWriter, type XlsxWriterOptions } from './xlsxWriter.js';

export type CreateXlsxStreamOptions = XlsxWriterOptions & {
    /** The messages of the workbook: rows of cells, records, and commands. */
    rows: ForAwaitable<SheetInput>;
};

/** One iterator for both kinds of source; `await` on a sync result is free. */
function iterate<T>(rows: ForAwaitable<T>): AsyncIterator<T> | Iterator<T> {
    return Symbol.asyncIterator in rows
        ? rows[Symbol.asyncIterator]()
        : rows[Symbol.iterator]();
}

/**
 * The source form of the writer, for records that are not already a stream:
 * an array, a generator, a database cursor. Returns the finished file as a
 * Web `ReadableStream<Uint8Array>`, ready to be piped at a file, an HTTP
 * response or a `Blob`.
 *
 * `rows` accepts anything iterable, sync or async, and carries the whole
 * workbook: rows of cells, records read by the sheet's columns, and the
 * `{ '#worksheet': name }` commands that open one sheet after another. The
 * stream pulls: messages are read only when the consumer asks for more bytes,
 * which is what keeps memory flat however many of them are coming.
 */
export function createXlsxStream(options: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    const iterator = iterate(options.rows);
    let writer!: XlsxWriter;
    let emitted = false;

    return new ReadableStream<Uint8Array>({
        start(controller) {
            // `rows` rides along in the options the writer gets; it reads the
            // ones it knows and this is the only place the extra one exists.
            writer = new XlsxWriter((bytes) => {
                emitted = true;
                controller.enqueue(bytes);
            }, options);
        },

        async pull(controller) {
            // Read records until the writer actually produces something, so
            // one `pull` always makes progress instead of spinning on the
            // rows that are still accumulating into the current batch.
            emitted = false;
            while (!emitted) {
                const next = await iterator.next();
                if (next.done) {
                    writer.finish();
                    controller.close();
                    return;
                }
                writer.writeRow(next.value);
            }
        },

        async cancel(reason) {
            await iterator.return?.(reason);
        },
    });
}
