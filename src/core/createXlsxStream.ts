import type { ForAwaitable, Row } from './types.js';
import { XlsxStream, type XlsxStreamOptions } from './xlsxStream.js';

export interface CreateXlsxStreamOptions extends XlsxStreamOptions {
    rows: ForAwaitable<Row>;
}

/**
 * How many records are grouped into one write to the stream.
 *
 * `XlsxStream` takes batches, and awaiting one write per record costs about
 * 1.8 µs each — 1.8 s over a million rows, measured, and it does not depend
 * on the compression level. Grouping them recovers that. The batch is the
 * only thing held: a hundred-odd records, whatever the row count.
 */
const PUMP_BATCH_ROWS = 256;

/**
 * Feeds `rows` into `writable` in batches. Each `write` is awaited, so the
 * stream's backpressure reaches all the way back to the row source and memory
 * stays flat however many records are coming.
 */
async function pump(rows: ForAwaitable<Row>, writable: WritableStream<readonly Row[]>): Promise<void> {
    const writer = writable.getWriter();
    try {
        let batch: Row[] = [];
        for await (const record of rows) {
            batch.push(record);
            if (batch.length >= PUMP_BATCH_ROWS) {
                await writer.write(batch);
                batch = [];
            }
        }
        if (batch.length) await writer.write(batch);
    } catch (err) {
        // Errors the readable side too, so the consumer gets this failure
        // instead of a silently truncated file.
        await writer.abort(err);
        return;
    }
    await writer.close();
}

/**
 * The source form of `XlsxStream`, for callers whose rows are not already a
 * stream: an array, a generator, a database cursor. Returns the same bytes,
 * as a Web `ReadableStream<Uint8Array>` ready to be piped at a file, an HTTP
 * response or a `Blob`.
 *
 * `rows` accepts anything iterable, sync or async. Use `XlsxStream` directly
 * when the records already arrive as a stream and this can sit in the pipe.
 */
export function createXlsxStream({
    rows,
    ...options
}: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    const stream = new XlsxStream(options);
    void pump(rows, stream.writable).catch(() => {
        // There is no second place to report to: whatever went wrong has
        // already errored the readable side (`abort` above, or a `close` that
        // failed while finishing the file), and that is where the consumer
        // reads it.
    });
    return stream.readable;
}
