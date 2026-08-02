// Node speaks the Web Streams standard natively, so the writer needs no
// Node-specific face: `XlsxStream` goes straight into a `pipeThrough` chain
// here exactly as it does in the browser. What is left in this module is the
// one thing the standard has no answer for in Node — a file as a destination
// — and the convenience wrapper built on it.
import { createWriteStream } from 'node:fs';
import { createXlsxStream, type CreateXlsxStreamOptions } from '../core/createXlsxStream.js';
import type { XlsxWriterOptions } from '../core/xlsxWriter.js';

/**
 * A file as a Web `WritableStream<Uint8Array>`, to close a `pipeTo`.
 *
 * Node has no native web-stream writer to a file, so this bridges its own
 * `fs.WriteStream` — the only Node stream left anywhere in the package.
 *
 * Node ships a ready-made bridge, `Writable.toWeb`, and this does not use it:
 * it does not carry the file's backpressure tightly enough. Writing a million
 * rows uncompressed (a 247 MB file, where the writer outruns the disk by the
 * widest margin) peaks at 348 MB of RSS through `Writable.toWeb` against
 * 111 MB straight into a web sink. Waiting on `drain` here, which is all the
 * difference amounts to, brings it back down.
 */
export function createFileWritable(path: string): WritableStream<Uint8Array> {
    const file = createWriteStream(path);

    // The file can fail before a single byte is offered to it — a bad path
    // fails on open — and an `error` nobody is listening for becomes an
    // uncaught exception. So it is always listened for, kept, and handed to
    // whichever step is waiting, or to the next one to start.
    let failure: Error | undefined;
    const waiting = new Set<(err: Error) => void>();
    file.on('error', (err: Error) => {
        failure = err;
        for (const reject of waiting) reject(err);
        waiting.clear();
    });

    /**
     * Resolves once `step` reports done, and rejects if the file fails —
     * either through the `error` event above or through the error a Node
     * callback is handed as its first argument.
     */
    function until(step: (done: (err?: Error | null) => void) => void): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            if (failure) {
                reject(failure);
                return;
            }
            waiting.add(reject);
            step((err) => {
                waiting.delete(reject);
                if (err) reject(err);
                else resolve();
            });
        });
    }

    return new WritableStream<Uint8Array>({
        write(chunk) {
            // `write` returns false once the file's own buffer is full, and
            // only then is there anything to wait for. Not resolving until
            // `drain` is what carries the backpressure back up the chain.
            return until((done) => {
                if (file.write(chunk)) done();
                else file.once('drain', done);
            });
        },
        close() {
            return until((done) => file.end(done));
        },
        abort(reason: unknown) {
            file.destroy(reason instanceof Error ? reason : new Error(String(reason)));
        },
    });
}

/**
 * Writes a whole `.xlsx` file from records that are not already a stream —
 * an array, a generator, a database cursor. Resolves once the file is closed.
 */
export async function writeXlsxFile<O extends XlsxWriterOptions>(
    path: string,
    options: CreateXlsxStreamOptions<O>,
): Promise<void> {
    await createXlsxStream(options).pipeTo(createFileWritable(path));
}
