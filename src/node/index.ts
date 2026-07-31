// Node-only helpers. Everything here touches `node:` modules, which is
// exactly why it lives outside src/core — core has to keep loading unchanged
// in the browser.
import { createWriteStream } from 'node:fs';
import { Duplex, Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import type {
    ReadableStream as NodeWebReadableStream,
    ReadableWritablePair as NodeWebReadableWritablePair,
} from 'node:stream/web';
import { createXlsxStream, type CreateXlsxStreamOptions } from '../core/createXlsxStream.js';
import { XlsxStream, type XlsxStreamOptions } from '../core/xlsxStream.js';

/**
 * `XlsxStream` as a Node `Duplex`, so the writer can sit in a plain `.pipe()`
 * chain:
 *
 * ```js
 * rowBatches.pipe(createXlsxDuplex({ columns })).pipe(createWriteStream(path))
 * ```
 *
 * The writable side is in object mode and takes `Row[]`, same as the
 * transform it wraps; the readable side emits the file's bytes.
 */
export function createXlsxDuplex(options: XlsxStreamOptions): Duplex {
    // Same objects at runtime; the cast only bridges the DOM lib's stream
    // declarations and node:stream/web's.
    const pair = new XlsxStream(options) as unknown as NodeWebReadableWritablePair;
    return Duplex.fromWeb(pair, { objectMode: true });
}

/** A Web `ReadableStream<Uint8Array>` as a Node `Readable`. */
export function toNodeReadable(stream: ReadableStream<Uint8Array>): Readable {
    return Readable.fromWeb(stream as unknown as NodeWebReadableStream<Uint8Array>);
}

/**
 * Writes a whole `.xlsx` file from records that are not already a stream —
 * an array, a generator, a database cursor. Resolves once the file is closed.
 */
export async function writeXlsxFile(path: string, options: CreateXlsxStreamOptions): Promise<void> {
    await pipeline(toNodeReadable(createXlsxStream(options)), createWriteStream(path));
}
