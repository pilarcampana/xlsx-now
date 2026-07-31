import { Zip, ZipDeflate, ZipPassThrough } from 'fflate';

const EMPTY = new Uint8Array(0);

const encoder = new TextEncoder();

/** Deflate effort, 0-9. `0` stores the entry uncompressed. */
export type CompressionLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/** Same default as most zip tools: the balanced point of the 0-9 scale. */
export const DEFAULT_COMPRESSION_LEVEL: CompressionLevel = 6;

/**
 * Push-based writer for a ZIP archive whose entry sizes are not known in
 * advance: bytes go in through `push`, and every byte `fflate` produces comes
 * straight back out through the `onChunk` sink, as soon as it is produced.
 *
 * Deliberately a plain ZIP 2.0 archive with deflate and data descriptors —
 * that combination is what OOXML consumers (Excel included) expect, and it is
 * why the total size does not have to be known before the first byte goes
 * out. ZIP64 is never emitted, which is fine here: a worksheet cannot get
 * anywhere near 4 GB within Excel's 1,048,576-row limit.
 */
export class ZipWriter {
    private readonly zip: Zip;
    private readonly level: CompressionLevel;
    private entry: ZipDeflate | ZipPassThrough | undefined;
    private failure: unknown;
    private closed = false;

    constructor(
        onChunk: (bytes: Uint8Array) => void,
        level: CompressionLevel = DEFAULT_COMPRESSION_LEVEL,
    ) {
        this.level = level;
        this.zip = new Zip((err, chunk, final) => {
            // `fflate` reports through this callback, which has no way to
            // throw back at whoever pushed; the error is kept here and raised
            // by `checkFailure` on that same operation, before it returns.
            if (err) {
                this.failure = err;
                return;
            }
            if (chunk.length) onChunk(chunk);
            if (final) this.closed = true;
        });
    }

    private checkFailure(): void {
        if (this.failure !== undefined) throw this.failure;
    }

    /** Opens an entry; every `push` from here on belongs to it. */
    startEntry(name: string): void {
        if (this.entry) throw new Error(`Cannot open "${name}": another zip entry is still open.`);
        // `ZipPassThrough` stores the entry verbatim; level 0 is the one case
        // where running it through the deflater would only add overhead for
        // no gain.
        this.entry =
            this.level === 0
                ? new ZipPassThrough(name)
                : new ZipDeflate(name, { level: this.level });
        this.zip.add(this.entry);
        this.checkFailure();
    }

    push(bytes: Uint8Array): void {
        if (!this.entry) throw new Error('Cannot push zip bytes: no entry is open.');
        this.entry.push(bytes);
        this.checkFailure();
    }

    endEntry(): void {
        if (!this.entry) throw new Error('Cannot end a zip entry: no entry is open.');
        this.entry.push(EMPTY, true);
        this.entry = undefined;
        this.checkFailure();
    }

    /** Writes a whole entry whose content is already in hand. */
    writeEntry(name: string, text: string): void {
        this.startEntry(name);
        this.push(encoder.encode(text));
        this.endEntry();
    }

    /** Closes the archive: central directory and end-of-central-directory. */
    end(): void {
        this.zip.end();
        this.checkFailure();
        if (!this.closed) {
            throw new Error('The zip stream ended before its central directory was written.');
        }
    }

    /**
     * Releases whatever `fflate` is holding. Only for an archive that is being
     * abandoned — it does not produce a readable file.
     */
    terminate(): void {
        this.zip.terminate();
    }
}
