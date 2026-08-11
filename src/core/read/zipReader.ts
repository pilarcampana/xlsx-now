// Reading the container back: the other half of `zip.ts`.
//
// The writer streams because it decides its own order — it opens an entry,
// pushes bytes, and says how big it was afterwards. Reading has no such
// freedom. A zip is addressed through the central directory at the end of the
// file, and an OOXML package addresses its parts by name through the
// relationships, so the order the entries happen to be in means nothing. This
// module is therefore the one place in the reader that seeks, and everything
// above it works in terms of named parts.
//
// What it does keep from the writer's world is that an entry is not read all
// at once unless someone asks for it whole: `entryChunks` inflates as it goes,
// which is what lets a worksheet of a million rows go through the parser
// without ever being a string.
import { Inflate } from 'fflate';
import type { RandomAccess } from './randomAccess.js';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_HEADER_SIGNATURE = 0x04034b50;

const EOCD_SIZE = 22;
const CENTRAL_HEADER_SIZE = 46;
const LOCAL_HEADER_SIZE = 30;

/** The comment at the end of a zip is a 16-bit length, so this is its most. */
const MAX_COMMENT_SIZE = 0xffff;

/** The two methods an OOXML package uses; anything else is refused by name. */
const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

/**
 * The value a 32-bit field carries when the real one is in a ZIP64 record,
 * and the 16-bit one for a count. Reading them means the archive is past
 * what this reader handles, which is worth saying rather than working out a
 * wrong offset from a placeholder.
 */
const ZIP64_MARKER_32 = 0xffffffff;
const ZIP64_MARKER_16 = 0xffff;

/** How much is inflated at a time. Big enough to matter, small enough to hold. */
const CHUNK_SIZE = 65536;

const decoder = new TextDecoder();

/** One entry of the archive, as its central directory describes it. */
export interface ZipEntry {
    name: string;
    method: number;
    compressedSize: number;
    uncompressedSize: number;
    /** Where the entry's own header is; the data follows it. */
    localHeaderOffset: number;
}

function view(bytes: Uint8Array): DataView {
    return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/**
 * Where the end-of-central-directory record starts.
 *
 * It is the last thing in the file, but its own last field is a comment of
 * any length up to 64 KB, so its position is not fixed and has to be searched
 * for backwards from the end. Backwards and not forwards because the
 * signature can also occur inside compressed data, and the last one is the
 * real one.
 */
function findEndOfCentralDirectory(tail: Uint8Array): number {
    const data = view(tail);
    for (let offset = tail.length - EOCD_SIZE; offset >= 0; offset--) {
        if (data.getUint32(offset, true) === EOCD_SIGNATURE) return offset;
    }
    throw new Error('This is not a zip archive: it has no end-of-central-directory record.');
}

/**
 * Every entry of the archive, by name.
 *
 * By name and not in order, because that is how a package is addressed: the
 * relationships name the parts and the reader looks them up. A duplicate name
 * would be an archive with two parts of the same name, which the last one
 * silently winning would hide — so the first is kept and the clash is raised.
 *
 * The sizes here are the ones to trust even for an entry written the way
 * `ZipWriter` writes them, with the sizes in a data descriptor after the
 * data: whatever the local header left at zero, the central directory
 * carries filled in.
 */
export async function readCentralDirectory(
    access: RandomAccess,
): Promise<ReadonlyMap<string, ZipEntry>> {
    const tailSize = Math.min(access.size, EOCD_SIZE + MAX_COMMENT_SIZE);
    const tail = await access.read(access.size - tailSize, tailSize);
    const eocd = findEndOfCentralDirectory(tail);
    const header = view(tail);

    const count = header.getUint16(eocd + 10, true);
    const directorySize = header.getUint32(eocd + 12, true);
    const directoryOffset = header.getUint32(eocd + 16, true);
    if (
        count === ZIP64_MARKER_16 ||
        directorySize === ZIP64_MARKER_32 ||
        directoryOffset === ZIP64_MARKER_32
    ) {
        throw new Error('This archive is ZIP64, which this reader does not handle.');
    }

    const directory = await access.read(directoryOffset, directorySize);
    const data = view(directory);
    const entries = new Map<string, ZipEntry>();
    let offset = 0;
    for (let index = 0; index < count; index++) {
        if (data.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) {
            throw new Error(`The central directory entry ${index} does not start with its signature.`);
        }
        const nameLength = data.getUint16(offset + 28, true);
        const extraLength = data.getUint16(offset + 30, true);
        const commentLength = data.getUint16(offset + 32, true);
        const name = decoder.decode(
            directory.subarray(offset + CENTRAL_HEADER_SIZE, offset + CENTRAL_HEADER_SIZE + nameLength),
        );
        const entry: ZipEntry = {
            name,
            method: data.getUint16(offset + 10, true),
            compressedSize: data.getUint32(offset + 20, true),
            uncompressedSize: data.getUint32(offset + 24, true),
            localHeaderOffset: data.getUint32(offset + 42, true),
        };
        if (
            entry.compressedSize === ZIP64_MARKER_32 ||
            entry.uncompressedSize === ZIP64_MARKER_32 ||
            entry.localHeaderOffset === ZIP64_MARKER_32
        ) {
            throw new Error(`The entry "${name}" is stored as ZIP64, which this reader does not handle.`);
        }
        if (!entries.has(name)) entries.set(name, entry);
        else throw new Error(`The archive carries two entries named "${name}".`);
        offset += CENTRAL_HEADER_SIZE + nameLength + extraLength + commentLength;
    }
    return entries;
}

/**
 * Where an entry's bytes begin.
 *
 * The central directory says where the *header* is, not the data, and the
 * header's own extra field is not required to be the same length as the one
 * the directory carries — so the local header has to be read to know where it
 * ends. A zip writer is allowed to differ there and some do.
 */
async function dataOffset(access: RandomAccess, entry: ZipEntry): Promise<number> {
    const header = view(await access.read(entry.localHeaderOffset, LOCAL_HEADER_SIZE));
    if (header.getUint32(0, true) !== LOCAL_HEADER_SIGNATURE) {
        throw new Error(`The entry "${entry.name}" does not start with a local file header.`);
    }
    const nameLength = header.getUint16(26, true);
    const extraLength = header.getUint16(28, true);
    return entry.localHeaderOffset + LOCAL_HEADER_SIZE + nameLength + extraLength;
}

/**
 * The bytes of an entry, inflated, as they come out.
 *
 * Chunk by chunk on both sides: `CHUNK_SIZE` of the file goes in, and
 * whatever `fflate` makes of it comes out. Nothing here holds the whole
 * entry, which is the point — the worksheet is read through this.
 */
export async function* entryChunks(
    access: RandomAccess,
    entry: ZipEntry,
): AsyncIterable<Uint8Array> {
    if (entry.method !== METHOD_STORE && entry.method !== METHOD_DEFLATE) {
        throw new Error(
            `The entry "${entry.name}" is compressed with method ${entry.method}, and this reader reads stored and deflated entries.`,
        );
    }
    const start = await dataOffset(access, entry);
    if (entry.method === METHOD_STORE) {
        for (let read = 0; read < entry.compressedSize; read += CHUNK_SIZE) {
            yield await access.read(start + read, Math.min(CHUNK_SIZE, entry.compressedSize - read));
        }
        return;
    }

    // `Inflate` hands its output to a callback as it produces it, and a
    // generator cannot yield from inside one — so a push fills this and the
    // loop drains it right after, which is the same order with the chunks in
    // hand instead of on the stack. A failure in there throws out of `push`,
    // so nothing is swallowed.
    const inflated: Uint8Array[] = [];
    const inflate = new Inflate((chunk) => {
        if (chunk.length) inflated.push(chunk);
    });
    for (let read = 0; read < entry.compressedSize; read += CHUNK_SIZE) {
        const length = Math.min(CHUNK_SIZE, entry.compressedSize - read);
        inflate.push(await access.read(start + read, length), read + length >= entry.compressedSize);
        yield* inflated.splice(0);
    }
}

/** An entry read whole, for the parts small enough that it is the simplest way. */
export async function readEntry(access: RandomAccess, entry: ZipEntry): Promise<Uint8Array> {
    const chunks: Uint8Array[] = [];
    let total = 0;
    for await (const chunk of entryChunks(access, entry)) {
        chunks.push(chunk);
        total += chunk.length;
    }
    const whole = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) {
        whole.set(chunk, at);
        at += chunk.length;
    }
    return whole;
}

/** An entry read whole, as the text of an XML part. */
export async function readEntryText(access: RandomAccess, entry: ZipEntry): Promise<string> {
    return decoder.decode(await readEntry(access, entry));
}
