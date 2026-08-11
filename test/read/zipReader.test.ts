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

describe('entryChunks', () => {
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
