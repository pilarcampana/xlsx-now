import assert from 'node:assert/strict';
import { DEFAULT_COMPRESSION_LEVEL, ZipWriter } from '../src/core/zip.js';
import { recordingSink } from './helpers/streams.js';
import { METHOD_DEFLATE, METHOD_STORE, MAX_VERSION, ZIP64_VERSION, readZipEntries } from './helpers/zip.js';

/** Text long enough that deflate has something to work with. */
const COMPRESSIBLE = 'the same line over and over\n'.repeat(500);

describe('ZipWriter', () => {
    it('defaults to the balanced point of the 0-9 scale', () => {
        assert.equal(DEFAULT_COMPRESSION_LEVEL, 6);
    });

    it('writes an archive yauzl reads back byte for byte', async () => {
        const { sink, bytes } = recordingSink();
        const zip = new ZipWriter(sink);
        zip.writeEntry('first.txt', 'hello');
        zip.writeEntry('dir/second.txt', COMPRESSIBLE);
        zip.end();

        const entries = await readZipEntries(bytes());
        assert.deepEqual(entries.map((entry) => entry.name), ['first.txt', 'dir/second.txt']);
        assert.equal(entries[0]?.text, 'hello');
        assert.equal(entries[1]?.text, COMPRESSIBLE);
        assert.ok(entries.every((entry) => entry.crcMatches), 'a CRC does not match');
    });

    it('joins into one entry everything pushed between start and end', async () => {
        const { sink, bytes } = recordingSink();
        const zip = new ZipWriter(sink);
        zip.startEntry('parts.txt');
        for (const part of ['a', 'b', 'c']) zip.push(new TextEncoder().encode(part));
        zip.endEntry();
        zip.end();

        const [entry] = await readZipEntries(bytes());
        assert.equal(entry?.text, 'abc');
    });

    it('deflates by default, and stores at level 0', async () => {
        for (const [level, method] of [[undefined, METHOD_DEFLATE], [0, METHOD_STORE]] as const) {
            const { sink, bytes } = recordingSink();
            const zip = level === undefined ? new ZipWriter(sink) : new ZipWriter(sink, level);
            zip.writeEntry('text.txt', COMPRESSIBLE);
            zip.end();

            const [entry] = await readZipEntries(bytes());
            assert.equal(entry?.method, method, `level ${level}`);
            assert.equal(entry?.uncompressedSize, COMPRESSIBLE.length);
            if (method === METHOD_STORE) {
                assert.equal(entry?.compressedSize, COMPRESSIBLE.length);
            } else {
                assert.ok((entry?.compressedSize ?? 0) < COMPRESSIBLE.length, 'not smaller');
            }
        }
    });

    it('compresses at least as hard at a higher level', () => {
        // Repetitive, but not the same line every time: level 1 and level 9
        // agree on text with nothing to look for.
        const mixed = Array.from(
            { length: 2000 },
            (_, i) => `row ${i}: ${'a value '.repeat(i % 9)}\n`,
        ).join('');

        const size = (level: 1 | 9): number => {
            const { sink, bytes } = recordingSink();
            const zip = new ZipWriter(sink, level);
            zip.writeEntry('text.txt', mixed);
            zip.end();
            return bytes().length;
        };
        assert.ok(size(9) <= size(1), 'level 9 came out bigger than level 1');
    });

    it('writes an archive Office can open: ZIP 2.0, no ZIP64, sizes after the data', async () => {
        const { sink, bytes } = recordingSink();
        const zip = new ZipWriter(sink);
        zip.writeEntry('text.txt', COMPRESSIBLE);
        zip.end();

        const [entry] = await readZipEntries(bytes());
        assert.ok(entry);
        assert.ok(entry.madeByVersion <= MAX_VERSION, `made by ZIP ${entry.madeByVersion / 10}`);
        assert.ok(entry.versionNeededToExtract < ZIP64_VERSION, 'ZIP64 entry');
        assert.ok(entry.streamed, 'sizes were known upfront');
        assert.ok(entry.localSizesZeroed, 'local header carries sizes despite bit 3');

        const tail = bytes().subarray(-128);
        assert.ok(!tail.includes(Buffer.from('PK\x06\x06')), 'ZIP64 end of central directory');
        assert.ok(!tail.includes(Buffer.from('PK\x06\x07')), 'ZIP64 locator');
    });

    it('hands the bytes out as they are produced, not at the end', () => {
        const { sink, chunks } = recordingSink();
        const zip = new ZipWriter(sink, 0);
        zip.writeEntry('text.txt', COMPRESSIBLE);
        assert.ok(chunks.length > 0, 'nothing came out before end()');
        zip.end();
    });

    it('refuses to open an entry while another one is open', () => {
        const zip = new ZipWriter(recordingSink().sink);
        zip.startEntry('first.txt');
        assert.throws(() => zip.startEntry('second.txt'), /another zip entry is still open/);
    });

    it('refuses to push or to close with no entry open', () => {
        const zip = new ZipWriter(recordingSink().sink);
        assert.throws(() => zip.push(new Uint8Array(1)), /no entry is open/);
        assert.throws(() => zip.endEntry(), /no entry is open/);
        zip.writeEntry('text.txt', 'x');
        assert.throws(() => zip.endEntry(), /no entry is open/);
        zip.end();
    });

    it('will not call an archive with an unclosed entry finished', () => {
        const zip = new ZipWriter(recordingSink().sink);
        zip.startEntry('never-ended.txt');
        assert.throws(() => zip.end(), /ended before its central directory was written/);
    });

    it('lets go of what fflate is holding on terminate', () => {
        const zip = new ZipWriter(recordingSink().sink);
        zip.startEntry('abandoned.txt');
        zip.push(new TextEncoder().encode('half a file'));
        // Nothing to assert about the bytes: a terminated archive is not a
        // readable file. What matters is that it does not throw and the
        // deflater is released.
        assert.doesNotThrow(() => zip.terminate());
    });

    it('raises what fflate reports on the operation that caused it', () => {
        const { sink } = recordingSink();
        const zip = new ZipWriter(sink);
        zip.writeEntry('text.txt', 'x');
        zip.end();
        // The archive is closed; fflate reports the extra entry through its
        // callback, and `startEntry` is where it surfaces.
        assert.throws(() => zip.startEntry('too-late.txt'));
    });

    it('lets a failing sink through to the caller', () => {
        const zip = new ZipWriter(() => {
            throw new Error('sink is full');
        }, 0);
        assert.throws(() => zip.writeEntry('text.txt', 'x'), /sink is full/);
    });
});
