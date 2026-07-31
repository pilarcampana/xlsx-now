import { Zip, ZipDeflate, ZipPassThrough } from 'fflate';
import type { ForAwaitable, ZipEntry } from './types.js';

/**
 * How much worksheet XML is accumulated before it is handed to the zip.
 *
 * The worksheet arrives one `<row>` at a time, and deflating each row on its
 * own is both slower and slightly worse (~2% larger output) than working on
 * bigger blocks. Batching does not buffer the workbook: only this much text
 * is ever held, no matter how many rows go through.
 */
const PUSH_BATCH_BYTES = 64 * 1024;

const EMPTY = new Uint8Array(0);

/** Deflate effort, 0-9. `0` stores the entry uncompressed. */
export type CompressionLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/** Same default as most zip tools: the balanced point of the 0-9 scale. */
export const DEFAULT_COMPRESSION_LEVEL: CompressionLevel = 6;

/**
 * Yields the entry's text as UTF-8, in blocks of at least `PUSH_BATCH_BYTES`
 * (the last block is whatever is left over).
 */
async function* encodeBatched(input: ZipEntry['input']): AsyncGenerator<Uint8Array, void, undefined> {
    const encoder = new TextEncoder();

    if (typeof input === 'string') {
        yield encoder.encode(input);
        return;
    }

    let batch = '';
    for await (const piece of input) {
        batch += piece;
        if (batch.length >= PUSH_BATCH_BYTES) {
            yield encoder.encode(batch);
            batch = '';
        }
    }
    if (batch) yield encoder.encode(batch);
}

/**
 * Drives `fflate`'s push-based `Zip` and yields the archive bytes as they are
 * produced.
 *
 * `Zip` reports output through a callback, which cannot be awaited — so
 * everything it emits lands in `pending`, and this generator drains it after
 * each push. Because the caller decides when to ask for the next chunk (see
 * `readableFromAsyncIterable`), reading the archive slowly is what stops the
 * rows from being consumed faster than they can be written out.
 */
async function* zipChunks(
    entries: ForAwaitable<ZipEntry>,
    level: CompressionLevel,
): AsyncGenerator<Uint8Array, void, undefined> {
    const pending: Uint8Array[] = [];
    let failure: unknown;
    let finished = false;

    const zip = new Zip((err, chunk, final) => {
        if (err) {
            failure = err;
            return;
        }
        if (chunk.length) pending.push(chunk);
        if (final) finished = true;
    });

    function* drain(): Generator<Uint8Array, void, undefined> {
        if (failure) throw failure;
        while (pending.length) yield pending.shift()!;
    }

    try {
        for await (const entry of entries) {
            // `ZipPassThrough` stores the entry verbatim; level 0 is the one
            // case where running it through the deflater would only add
            // overhead for no gain.
            const file =
                level === 0 ? new ZipPassThrough(entry.name) : new ZipDeflate(entry.name, { level });
            zip.add(file);

            for await (const bytes of encodeBatched(entry.input)) {
                file.push(bytes);
                yield* drain();
            }
            file.push(EMPTY, true);
            yield* drain();
        }

        zip.end();
        yield* drain();

        if (!finished) throw new Error('The zip stream ended before its central directory was written.');
    } catch (err) {
        // Releases whatever `fflate` is holding; the archive is incomplete
        // and about to be discarded anyway.
        zip.terminate();
        throw err;
    }
}

/**
 * Wraps an async iterable of bytes in a Web `ReadableStream`, pulling exactly
 * one chunk per consumer request so backpressure reaches all the way back to
 * the row source.
 */
function readableFromAsyncIterable(source: AsyncIterable<Uint8Array>): ReadableStream<Uint8Array> {
    const iterator = source[Symbol.asyncIterator]();

    return new ReadableStream<Uint8Array>({
        async pull(controller) {
            const { done, value } = await iterator.next();
            if (done) controller.close();
            else controller.enqueue(value);
        },
        async cancel(reason) {
            await iterator.return?.(reason);
        },
    });
}

/**
 * Builds a ZIP archive as a Web `ReadableStream<Uint8Array>` from entries
 * whose sizes are not known in advance.
 *
 * Deliberately a plain ZIP 2.0 archive with deflate and data descriptors —
 * that combination is what OOXML consumers (Excel included) expect, and it is
 * why the total size does not have to be known before the first byte goes
 * out. ZIP64 is never emitted, which is fine here: a worksheet cannot get
 * anywhere near 4 GB within Excel's 1,048,576-row limit.
 */
export function createZipStream(
    entries: ForAwaitable<ZipEntry>,
    level: CompressionLevel = DEFAULT_COMPRESSION_LEVEL,
): ReadableStream<Uint8Array> {
    return readableFromAsyncIterable(zipChunks(entries, level));
}
