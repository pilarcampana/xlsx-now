// Node speaks the Web Streams standard natively, so the writer needs no
// Node-specific face: `XlsxStream` goes straight into a `pipeThrough` chain
// here exactly as it does in the browser. What is left in this module is the
// one thing the standard has no answer for in Node — a file as a destination
// — and the convenience wrapper built on it.
import { createWriteStream } from 'node:fs';
import { open, type FileHandle } from 'node:fs/promises';
import { createXlsxStream, type CreateXlsxStreamOptions } from '../core/createXlsxStream.js';
import type { RandomAccess } from '../core/read/randomAccess.js';
import { openXlsx, type ReadOptions, type XlsxReader } from '../core/read/readXlsx.js';
import type { ReadMode, ReadModes, SheetData } from '../core/read/types.js';

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
export async function writeXlsxFile(
    path: string,
    options: CreateXlsxStreamOptions,
): Promise<void> {
    await createXlsxStream(options).pipeTo(createFileWritable(path));
}

/**
 * An open file as something the reader can seek in.
 *
 * This is the whole of what Node has to add to the reader, and it is what
 * makes the reading side keep the promise the writing side makes: a file goes
 * through without ever being in memory whole. The reader asks for the parts
 * it needs, where they are, and a worksheet goes past 64 KB at a time.
 */
function fileAccess(file: FileHandle, size: number): RandomAccess {
    return {
        size,
        async read(offset: number, length: number): Promise<Uint8Array> {
            const bytes = new Uint8Array(length);
            const { bytesRead } = await file.read(bytes, 0, length, offset);
            if (bytesRead !== length) {
                // A short read here is a file that ended where the archive
                // says it does not — truncated, or being written into while
                // it is read — and carrying on with the zeroes that are left
                // in the buffer would read that as data.
                throw new Error(
                    `The file ended after ${bytesRead} of the ${length} bytes the archive points at from ${offset}.`,
                );
            }
            return bytes;
        },
    };
}

/** An open workbook, and the file it is being read out of. */
export interface XlsxFileReader<C> extends XlsxReader<C> {
    /** Closes the file. The sheets cannot be read after it. */
    close(): Promise<void>;
}

/**
 * Opens a workbook file: reads everything except the rows.
 *
 * The one to reach for when the rows are not all wanted at once — a sheet
 * bigger than memory, or a run that stops at the first row that matters.
 * Whoever opens it closes it.
 *
 * ```js
 * const workbook = await openXlsxFile('ventas.xlsx');
 * try {
 *     for await (const row of workbook.sheets[0].rows()) console.log(row.cells);
 * } finally {
 *     await workbook.close();
 * }
 * ```
 */
export async function openXlsxFile<M extends ReadMode = 'values'>(
    path: string,
    options: ReadOptions<M> = {},
): Promise<XlsxFileReader<ReadModes[M]>> {
    const file = await open(path);
    try {
        const { size } = await file.stat();
        const reader = await openXlsx(fileAccess(file, size), options);
        return { sheets: reader.sheets, close: () => file.close() };
    } catch (err) {
        // The handle is this function's until it is handed over, and a
        // package that fails to open would leave it behind.
        await file.close();
        throw err;
    }
}

/** Every sheet of a workbook file, read whole. */
export async function readXlsxFile<M extends ReadMode = 'values'>(
    path: string,
    options: ReadOptions<M> = {},
): Promise<SheetData<ReadModes[M]>[]> {
    const workbook = await openXlsxFile(path, options);
    try {
        const sheets: SheetData<ReadModes[M]>[] = [];
        for (const sheet of workbook.sheets) sheets.push(await sheet.read());
        return sheets;
    } finally {
        await workbook.close();
    }
}
