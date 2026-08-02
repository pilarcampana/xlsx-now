import type { ForAwaitable } from './types.js';
import { XlsxWriter, type RowOf, type XlsxWriterOptions } from './xlsxWriter.js';

export type CreateXlsxStreamOptions<O extends XlsxWriterOptions = XlsxWriterOptions> = O & {
    rows: ForAwaitable<RowOf<O>>;
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
 * `rows` accepts anything iterable, sync or async. The stream pulls: records
 * are read only when the consumer asks for more bytes, which is what keeps
 * memory flat however many of them are coming.
 */
export function createXlsxStream<O extends XlsxWriterOptions>(
    options: CreateXlsxStreamOptions<O>,
): ReadableStream<Uint8Array> {
    const iterator = iterate(options.rows);
    let writer!: XlsxWriter<O>;
    let emitted = false;

    return new ReadableStream<Uint8Array>({
        start(controller) {
            // `rows` rides along in the options the writer gets; it reads the
            // ones it knows and this is the only place the extra one exists.
            writer = new XlsxWriter<O>((bytes) => {
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
