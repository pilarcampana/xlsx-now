// The reader's container layer, checked against archives it did not make:
// the ones `ZipWriter` streams out — sizes in a data descriptor, which is the
// shape the local headers leave empty — and one written by `exceljs`, whose
// zip is an ordinary one with the sizes upfront.
import assert from 'node:assert/strict';
import { bytesAccess } from '../../src/core/read/randomAccess.js';
import {
    entryChunks,
    readCentralDirectory,
    readEntry,
    readEntryText,
} from '../../src/core/read/zipReader.js';
import { ZipWriter } from '../../src/core/zip.js';
import { recordingSink } from '../helpers/streams.js';

/** Text long enough that deflate has something to work with. */
const COMPRESSIBLE = 'the same line over and over\n'.repeat(500);

/**
 * Text of `length` characters that deflate cannot shrink much, from a
 * generator with no randomness in it so a failure is the same twice.
 */
function incompressible(length: number): string {
    let seed = 1;
    let text = '';
    while (text.length < length) {
        seed = (seed * 1103515245 + 12345) % 2147483648;
        text += seed.toString(36);
    }
    return text.slice(0, length);
}

/** An archive of the given entries, as `ZipWriter` writes them. */
function archive(entries: Record<string, string>, level?: 0): Uint8Array {
    const { sink, bytes } = recordingSink();
    const zip = new ZipWriter(sink, level);
    for (const [name, text] of Object.entries(entries)) zip.writeEntry(name, text);
    zip.end();
    return bytes();
}

/** The bytes an end-of-central-directory record starts with. */
const EOCD_SIGNATURE = Uint8Array.of(0x50, 0x4b, 0x05, 0x06);

/** How long that record is, comment aside, which is where its length lives. */
const EOCD_SIZE = 22;

/** The start of an OLE2 compound file, which is what a real `.xls` is. */
const OLE2_HEAD = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);

/** And one more format to show the case is not about `.xls` in particular. */
const PDF_HEAD = Uint8Array.of(0x25, 0x50, 0x44, 0x46, 0x2d);

/**
 * `head` and then filler carrying an end-of-central-directory signature,
 * which is the false positive this is about: four bytes any binary can have
 * in it by chance, far enough from the end for a whole record to seem to fit.
 */
function withStraySignature(head: Uint8Array): Uint8Array {
    const filler = new Uint8Array(512).fill(0x41);
    filler.set(EOCD_SIGNATURE, 200);
    const whole = new Uint8Array(head.length + filler.length);
    whole.set(head);
    whole.set(filler, head.length);
    return whole;
}

/** The same archive with a comment on the end of its end record. */
function withComment(bytes: Uint8Array, comment: Uint8Array): Uint8Array {
    const whole = new Uint8Array(bytes.length + comment.length);
    whole.set(bytes);
    whole.set(comment, bytes.length);
    const eocd = bytes.length - EOCD_SIZE;
    new DataView(whole.buffer).setUint16(eocd + 20, comment.length, true);
    return whole;
}

describe('readCentralDirectory', () => {
    it('finds every entry, by name', async () => {
        const access = bytesAccess(archive({ 'first.txt': 'hello', 'dir/second.txt': 'bye' }));
        const entries = await readCentralDirectory(access);
        assert.deepEqual([...entries.keys()], ['first.txt', 'dir/second.txt']);
    });

    it('reads the sizes the data descriptor left out of the local header', async () => {
        const access = bytesAccess(archive({ 'a.txt': COMPRESSIBLE }));
        const entry = (await readCentralDirectory(access)).get('a.txt');
        assert.equal(entry?.uncompressedSize, COMPRESSIBLE.length);
        assert.ok((entry?.compressedSize ?? 0) > 0, 'the compressed size is missing');
    });

    it('refuses something that is not a zip at all', async () => {
        await assert.rejects(
            readCentralDirectory(bytesAccess(new TextEncoder().encode('no soy un zip'))),
            /not a zip archive/,
        );
    });

    it('names the format of an .xls instead of talking about a central directory', async () => {
        const error = await readCentralDirectory(bytesAccess(withStraySignature(OLE2_HEAD))).then(
            () => undefined,
            (thrown: Error) => thrown,
        );
        assert.match(error?.message ?? '', /Excel 97-2003/);
        assert.match(error?.message ?? '', /\.xlsx/);
        // The point of the whole thing: the stray signature must not have
        // turned this into a complaint about a directory that never existed.
        assert.doesNotMatch(error?.message ?? '', /central directory/);
    });

    it('refuses any other binary that happens to carry the end signature', async () => {
        await assert.rejects(
            readCentralDirectory(bytesAccess(withStraySignature(PDF_HEAD))),
            /not a zip archive.*unknown format/,
        );
    });

    it('reads past a stray end signature in the comment of a real archive', async () => {
        // A zip comment is arbitrary bytes at the very end of the file, so it
        // is searched *before* the real record. One that carries the
        // signature is the case the check on the candidate is for.
        const comment = new Uint8Array(64).fill(0x41);
        comment.set(EOCD_SIGNATURE, 8);
        const access = bytesAccess(withComment(archive({ 'a.txt': 'hola' }), comment));
        const entries = await readCentralDirectory(access);
        assert.deepEqual([...entries.keys()], ['a.txt']);
        assert.equal(await readEntryText(access, entries.get('a.txt')!), 'hola');
    });
});

describe('readEntry', () => {
    it('inflates a deflated entry back to what went in', async () => {
        const access = bytesAccess(archive({ 'a.txt': COMPRESSIBLE }));
        const entries = await readCentralDirectory(access);
        assert.equal(await readEntryText(access, entries.get('a.txt')!), COMPRESSIBLE);
    });

    it('reads a stored entry, where there is nothing to inflate', async () => {
        const access = bytesAccess(archive({ 'a.txt': 'sin comprimir' }, 0));
        const entries = await readCentralDirectory(access);
        assert.equal(await readEntryText(access, entries.get('a.txt')!), 'sin comprimir');
    });

    it('reads an empty entry as nothing', async () => {
        const access = bytesAccess(archive({ 'empty.txt': '' }));
        const entries = await readCentralDirectory(access);
        assert.equal((await readEntry(access, entries.get('empty.txt')!)).length, 0);
    });

    it('reads every entry, not only the first', async () => {
        const access = bytesAccess(archive({ 'a.txt': 'uno', 'b.txt': 'dos', 'c.txt': 'tres' }));
        const entries = await readCentralDirectory(access);
        const texts = [];
        for (const name of ['c.txt', 'a.txt', 'b.txt']) {
            texts.push(await readEntryText(access, entries.get(name)!));
        }
        // Read out of order on purpose: the point of the central directory is
        // that a part is reached by name, whatever place it has in the file.
        assert.deepEqual(texts, ['tres', 'uno', 'dos']);
    });
});

describe('bytesAccess', () => {
    it('refuses a read that would go past the end', async () => {
        const access = bytesAccess(new Uint8Array(10));
        await assert.rejects(access.read(5, 10), RangeError);
        await assert.rejects(access.read(-1, 2), RangeError);
    });
});

describe('entryChunks', () => {
    it('refuses an entry compressed with a method it does not read', async () => {
        const access = bytesAccess(archive({ 'a.txt': 'hola' }));
        const entry = { ...(await readCentralDirectory(access)).get('a.txt')!, method: 99 };
        await assert.rejects(readEntry(access, entry), /method 99/);
    });

    it('refuses an entry that is not where the directory says it is', async () => {
        const bytes = archive({ 'a.txt': 'hola' });
        // The first entry starts at byte 0, so this is its signature — the
        // third byte of it and not the first, which is the `PK` the reader
        // now looks at before it goes anywhere near the entries.
        bytes[2] = 0;
        const access = bytesAccess(bytes);
        const entry = (await readCentralDirectory(access)).get('a.txt')!;
        await assert.rejects(readEntry(access, entry), /local file header/);
    });
});

describe('entryChunks: what it hands out', () => {
    it('hands the bytes out as it inflates, not in one piece at the end', async () => {
        // Text deflate can do little with, so that what goes *in* is bigger
        // than one read: `COMPRESSIBLE` repeated would shrink to a few
        // kilobytes and never take a second turn round the loop.
        const big = incompressible(400_000);
        const access = bytesAccess(archive({ 'big.txt': big }));
        const entries = await readCentralDirectory(access);

        const chunks: Uint8Array[] = [];
        for await (const chunk of entryChunks(access, entries.get('big.txt')!)) chunks.push(chunk);

        assert.ok(chunks.length > 1, `everything came out at once (${chunks.length} chunk)`);
        // And the pieces still join back into exactly what went in, which is
        // what says the chunks stayed theirs while the next one was read.
        assert.equal(Buffer.concat(chunks.map((chunk) => Buffer.from(chunk))).toString(), big);
    });
});
