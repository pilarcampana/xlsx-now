// Node-only helpers. Everything here touches `node:` modules, which is
// exactly why it lives outside src/core — core has to keep loading unchanged
// in the browser.
import { createWriteStream } from 'node:fs';
import { Readable, Transform, type TransformCallback } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';
import { createXlsxStream, type CreateXlsxStreamOptions } from '../core/createXlsxStream.js';
import type { Row } from '../core/types.js';
import { XlsxWriter, type XlsxWriterOptions } from '../core/xlsxWriter.js';

export type XlsxTransformOptions = XlsxWriterOptions;

/** Runs `step`, reporting either outcome through a Node stream callback. */
function report(step: () => void, callback: TransformCallback): void {
    try {
        step();
    } catch (err) {
        callback(err instanceof Error ? err : new Error(String(err)));
        return;
    }
    callback();
}

/**
 * A styled `.xlsx` as a native Node `Transform`, so the writer can sit in a
 * plain `.pipe()` chain between whatever produces the records and wherever
 * the file has to go:
 *
 * ```js
 * createReadStream('input.ndjson')
 *     .pipe(new LineSplitter({}))
 *     .pipe(new JsonParserTransformer())
 *     .pipe(new XlsxTransform({ columns, sheetName: 'Widgets' }))
 *     .pipe(createWriteStream('widgets.xlsx'));
 * ```
 *
 * The writable side is in object mode and takes one record per chunk; the
 * readable side emits the file's bytes. Nothing is buffered: Node's own
 * backpressure — the readable side full, so `_transform` is not called again
 * — is what stops records from being consumed faster than they can be
 * written out.
 */
export class XlsxTransform extends Transform {
    readonly #writer: XlsxWriter;

    constructor(options: XlsxTransformOptions) {
        super({ writableObjectMode: true });
        this.#writer = new XlsxWriter((bytes) => {
            this.push(bytes);
        }, options);
    }

    override _transform(record: Row, _encoding: BufferEncoding, callback: TransformCallback): void {
        report(() => this.#writer.writeRow(record), callback);
    }

    override _flush(callback: TransformCallback): void {
        report(() => this.#writer.finish(), callback);
    }
}

/** A Web `ReadableStream<Uint8Array>` as a Node `Readable`. */
export function toNodeReadable(stream: ReadableStream<Uint8Array>): Readable {
    // Same object at runtime; the cast only bridges the DOM lib's
    // ReadableStream declaration and node:stream/web's.
    return Readable.fromWeb(stream as unknown as NodeWebReadableStream<Uint8Array>);
}

/**
 * Writes a whole `.xlsx` file from records that are not already a stream —
 * an array, a generator, a database cursor. Resolves once the file is closed.
 */
export async function writeXlsxFile(path: string, options: CreateXlsxStreamOptions): Promise<void> {
    await pipeline(toNodeReadable(createXlsxStream(options)), createWriteStream(path));
}
