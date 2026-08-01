import type { Row } from './types.js';
import { XlsxWriter, type XlsxWriterOptions } from './xlsxWriter.js';

export type XlsxStreamOptions = XlsxWriterOptions;

function xlsxTransformer(options: XlsxStreamOptions): Transformer<Row, Uint8Array> {
    // Assigned by `start`, which the stream always runs before `transform`
    // and `flush`.
    let writer!: XlsxWriter;

    return {
        start(controller) {
            writer = new XlsxWriter((bytes) => controller.enqueue(bytes), options);
        },
        transform(record) {
            writer.writeRow(record);
        },
        flush() {
            writer.finish();
        },
    };
}

/**
 * A styled `.xlsx` as a Web `TransformStream`: records go in the writable
 * side, one per chunk, and the bytes of the file come out the readable side.
 *
 * ```js
 * rows.pipeThrough(new XlsxStream({ columns })).pipeTo(destination)
 * ```
 *
 * This is the browser-side face of `XlsxWriter`; in Node, `XlsxTransform`
 * from `xlsx-now/node` is the same writer as a native `stream.Transform`, for
 * a plain `.pipe()` chain.
 *
 * Nothing is buffered: the file is emitted as the rows arrive, and the
 * standard's own backpressure — readable side full, writable side not ready —
 * is what stops rows from being consumed faster than they can be written out.
 */
export class XlsxStream extends TransformStream<Row, Uint8Array> {
    constructor(options: XlsxStreamOptions) {
        super(xlsxTransformer(options));
    }
}
