import type { ForAwaitable, Row } from './types.js';
import { XlsxWriter, type XlsxWriterOptions } from './xlsxWriter.js';

export interface CreateXlsxStreamOptions extends XlsxWriterOptions {
    rows: ForAwaitable<Row>;
}

/** One iterator for both kinds of source; `await` on a sync result is free. */
function iterate(rows: ForAwaitable<Row>): AsyncIterator<Row> | Iterator<Row> {
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
export function createXlsxStream({
    rows,
    ...options
}: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    const iterator = iterate(rows);
    let writer!: XlsxWriter;
    let emitted = false;

    return new ReadableStream<Uint8Array>({
        start(controller) {
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
