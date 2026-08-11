/**
 * How the reader reaches the bytes of an archive.
 *
 * A one-method interface, and the reason it exists at all is the format: a
 * zip says where its parts are in a directory at the *end* of the file, and
 * an xlsx does not promise any useful order before that — a worksheet can
 * sit ahead of the shared strings its cells refer to by index, and the part
 * that names the sheets can be the last entry of all. So a reader has to be
 * able to go back, which a forward-only stream cannot do.
 *
 * Being able to go back is not the same as holding everything: `Blob.slice`
 * in a browser and a positional read in Node are both this interface, and
 * neither loads the file to answer. `bytesAccess` is the one that does, and
 * it is here because a `Uint8Array` in hand is the common case.
 */
export interface RandomAccess {
    /** Total bytes, which is what makes the end of the file findable. */
    readonly size: number;
    /**
     * The bytes at `[offset, offset + length)`. A read that would go past
     * the end is a corrupt archive pointing outside itself, not a short
     * read to be worked around, so it fails rather than returning less.
     */
    read(offset: number, length: number): Promise<Uint8Array>;
}

/** An archive already in memory, as something the reader can seek in. */
export function bytesAccess(bytes: Uint8Array): RandomAccess {
    return {
        size: bytes.length,
        async read(offset: number, length: number): Promise<Uint8Array> {
            if (offset < 0 || length < 0 || offset + length > bytes.length) {
                throw new RangeError(
                    `The archive points to bytes ${offset}..${offset + length}, and it is ${bytes.length} bytes long.`,
                );
            }
            return bytes.subarray(offset, offset + length);
        },
    };
}
