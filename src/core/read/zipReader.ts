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

/**
 * The first bytes of an OLE2 compound file, which is what a real `.xls` is.
 *
 * Nothing is ever read from one, and this is not a check on the way in: it is
 * looked at only after a file has already failed to open, to say which file
 * it was. An `.xls` arrives here often — it is a spreadsheet, and the
 * extension is one letter away — and "save it as .xlsx" is something the
 * person holding it can do, where "invalid file" is not.
 */
const OLE2_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

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

function startsWith(bytes: Uint8Array, signature: readonly number[]): boolean {
    return bytes.length >= signature.length && signature.every((byte, at) => bytes[at] === byte);
}

/**
 * The failure for a file that would not open, put in the terms of whoever
 * handed it over.
 *
 * By the time this is called the file has already been given every chance to
 * be read, so it changes nothing about what opens — only what is said about
 * what did not. And what is said is about the file, not about the container:
 * "no end-of-central-directory record" is a true sentence about zips and no
 * answer at all to someone whose spreadsheet will not open.
 */
async function unreadableFormat(access: RandomAccess): Promise<Error> {
    const head = await access.read(0, Math.min(access.size, OLE2_SIGNATURE.length));
    if (startsWith(head, OLE2_SIGNATURE)) {
        return new Error('Old Excel 97-2003 .xls file format detected. Format .xlsx expected.');
    }
    return new Error('Invalid file format. Expecting an .xlsx file.');
}

/** The three fields of the end record that say where the directory is. */
interface EndOfCentralDirectory {
    count: number;
    directorySize: number;
    directoryOffset: number;
}

/**
 * Whether the directory a candidate record points at is really there.
 *
 * The four bytes of the signature are not enough to have found the record:
 * they can fall inside an archive comment as easily as inside compressed
 * data. So the candidate is asked to agree with the file — a directory that
 * fits inside it, and a central header where it says one starts. Without
 * this, a wrong candidate is only noticed further down, in the entry loop,
 * which then reports a malformed directory for an archive whose directory is
 * fine and was never the one being read.
 *
 * Nothing that used to open stops opening for this: a candidate turned down
 * here is one that had no directory to read, and the search goes on to the
 * record that has one instead of stopping at the first four bytes that
 * looked like it.
 */
async function pointsAtCentralDirectory(
    access: RandomAccess,
    { count, directorySize, directoryOffset }: EndOfCentralDirectory,
): Promise<boolean> {
    // The ZIP64 placeholders point nowhere by design, and they are what a
    // real ZIP64 archive carries here: taking the candidate means the refusal
    // names ZIP64, which is the true reason, instead of the search walking
    // past the record and ending at "no end-of-central-directory".
    if (
        count === ZIP64_MARKER_16 ||
        directorySize === ZIP64_MARKER_32 ||
        directoryOffset === ZIP64_MARKER_32
    ) {
        return true;
    }
    if (directoryOffset + directorySize > access.size) return false;
    // An archive of no entries has no header to look at, and a directory that
    // fits in the file is all there is to ask of it. Asking for more — that
    // an empty directory also be zero bytes long — would turn one such file
    // away, and turning anything away is not what this is for.
    if (count === 0) return true;
    if (directorySize < CENTRAL_HEADER_SIZE) return false;
    return view(await access.read(directoryOffset, 4)).getUint32(0, true) === CENTRAL_HEADER_SIGNATURE;
}

/**
 * The end-of-central-directory record: where the entries are, and how many.
 *
 * It is the last thing in the file, but its own last field is a comment of
 * any length up to 64 KB, so its position is not fixed and has to be searched
 * for backwards from the end. Backwards and not forwards because the
 * signature can also occur inside compressed data, and the last one is the
 * real one — and every candidate is checked against the file before it is
 * taken, because being last is not the same as being right.
 *
 * Running out of candidates is where a file that is not a package ends up,
 * whatever it is, so that is where it gets told what it was holding.
 */
async function findEndOfCentralDirectory(
    access: RandomAccess,
    tail: Uint8Array,
): Promise<EndOfCentralDirectory> {
    const data = view(tail);
    for (let offset = tail.length - EOCD_SIZE; offset >= 0; offset--) {
        if (data.getUint32(offset, true) !== EOCD_SIGNATURE) continue;
        const record: EndOfCentralDirectory = {
            count: data.getUint16(offset + 10, true),
            directorySize: data.getUint32(offset + 12, true),
            directoryOffset: data.getUint32(offset + 16, true),
        };
        if (await pointsAtCentralDirectory(access, record)) return record;
    }
    throw await unreadableFormat(access);
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
    const { count, directorySize, directoryOffset } = await findEndOfCentralDirectory(access, tail);
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
