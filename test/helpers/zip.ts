// Reads a generated archive back with implementations that had nothing to do
// with writing it — `yauzl` plus Node's own `zlib` — for the same reason
// `scripts/validate-xlsx.ts` does: checking with `fflate` would only prove
// that the writer agrees with itself.
import { crc32 } from 'node:zlib';
import { fromBufferPromise, type Entry } from 'yauzl';

/** General-purpose flag bit 3: sizes and CRC follow the data, not precede it. */
export const FLAG_DATA_DESCRIPTOR = 0x0008;
export const METHOD_STORE = 0;
export const METHOD_DEFLATE = 8;
/** ZIP 2.0, the highest version Office accepts. */
export const MAX_VERSION = 20;
/** "Version needed to extract" value that means ZIP64. */
export const ZIP64_VERSION = 45;

export const SHEET_PART = 'xl/worksheets/sheet1.xml';

export interface ZipEntry {
    name: string;
    /** The inflated bytes, as text. */
    text: string;
    data: Buffer;
    /** 0 stored, 8 deflated. */
    method: number;
    compressedSize: number;
    uncompressedSize: number;
    /** Low byte of `versionMadeBy`: the ZIP version, times ten. */
    madeByVersion: number;
    versionNeededToExtract: number;
    /** The archive's own CRC agrees with one recomputed from the bytes. */
    crcMatches: boolean;
    /** Local header flagged bit 3: the sizes were not known upfront. */
    streamed: boolean;
    /** ...and the local header does carry the zeroes bit 3 promises. */
    localSizesZeroed: boolean;
}

async function readAll(stream: AsyncIterable<unknown>): Promise<Buffer> {
    const chunks: Buffer[] = [];
    for await (const chunk of stream) chunks.push(chunk as Buffer);
    return Buffer.concat(chunks);
}

/**
 * Every entry of `bytes`, in archive order, inflated and with the container
 * properties an OOXML consumer looks at already read off the headers.
 *
 * `yauzl` verifies the declared sizes while inflating (`validateEntrySizes`),
 * so a truncated or corrupt entry fails here rather than returning short.
 */
export async function readZipEntries(bytes: Uint8Array): Promise<ZipEntry[]> {
    const zip = await fromBufferPromise(Buffer.from(bytes), { lazyEntries: true });
    const entries: ZipEntry[] = [];
    try {
        for await (const entry of zip.eachEntry() as AsyncIterable<Entry>) {
            const data = await readAll(await zip.openReadStreamPromise(entry));
            const header = await zip.readLocalFileHeaderPromise(entry);
            entries.push({
                name: entry.fileName,
                text: data.toString('utf8'),
                data,
                method: entry.compressionMethod,
                compressedSize: entry.compressedSize,
                uncompressedSize: entry.uncompressedSize,
                madeByVersion: entry.versionMadeBy & 0xff,
                versionNeededToExtract: entry.versionNeededToExtract,
                crcMatches: crc32(data) === entry.crc32,
                streamed: (header.generalPurposeBitFlag & FLAG_DATA_DESCRIPTOR) !== 0,
                localSizesZeroed:
                    header.crc32 === 0 &&
                    header.compressedSize === 0 &&
                    header.uncompressedSize === 0,
            });
        }
    } finally {
        zip.close();
    }
    return entries;
}

export interface XlsxParts {
    /** Entry names, in the order the archive carries them. */
    names: string[];
    byName: Map<string, ZipEntry>;
    /** `xl/worksheets/sheet1.xml`, the part every test looks at. */
    sheet: string;
}

export async function readXlsx(bytes: Uint8Array): Promise<XlsxParts> {
    const entries = await readZipEntries(bytes);
    const byName = new Map(entries.map((entry) => [entry.name, entry]));
    return {
        names: entries.map((entry) => entry.name),
        byName,
        sheet: byName.get(SHEET_PART)?.text ?? '',
    };
}

/** The `<row>` elements of a worksheet, tags and all. */
export function sheetRows(sheetXml: string): string[] {
    return sheetXml.match(/<row .*?<\/row>/g) ?? [];
}
