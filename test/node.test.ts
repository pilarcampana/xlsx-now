import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';
import { createFileWritable, writeXlsxFile } from '../src/node/index.js';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
import type { CellRow, Column } from '../src/core/types.js';
import { readXlsx, sheetRows } from './helpers/zip.js';

const COLUMNS: readonly Column[] = [{ name: 'id', pk: true }, { name: 'name' }];

describe('the Node face', () => {
    let dir: string;

    beforeEach(async () => {
        dir = await mkdtemp(join(tmpdir(), 'xlsx-now-'));
    });

    afterEach(async () => {
        await rm(dir, { recursive: true, force: true });
    });

    describe('writeXlsxFile', () => {
        it('leaves a readable .xlsx on disk', async () => {
            const path = join(dir, 'people.xlsx');
            await writeXlsxFile(path, {
                columns: COLUMNS,
                rows: [
                    { id: 1, name: 'Ana' },
                    { id: 2, name: 'Beto' },
                ],
            });

            const { names, sheet } = await readXlsx(await readFile(path));
            assert.equal(names.length, 6);
            assert.equal(sheetRows(sheet).length, 3); // header + 2
            assert.ok(sheet.includes('>Beto<'));
        });

        it('resolves only once the file is closed', async () => {
            const path = join(dir, 'closed.xlsx');
            await writeXlsxFile(path, { rows: [['a']] });
            // Readable straight after the promise settles: nothing is still
            // on its way to the file.
            assert.ok((await readFile(path)).length > 0);
        });

        it('reports a path it cannot write to', async () => {
            const path = join(dir, 'no', 'such', 'directory', 'out.xlsx');
            await assert.rejects(writeXlsxFile(path, { rows: [['a']] }), /ENOENT/);
        });

        it('carries the file\'s backpressure back up the chain', async () => {
            // Big enough (uncompressed) that the file's own buffer fills and
            // the writer has to wait on `drain` rather than write straight
            // through.
            const path = join(dir, 'big.xlsx');
            const rows: CellRow[] = Array.from({ length: 20000 }, (_, i) => [
                i,
                'a fairly long value, repeated so the file gets past a few buffers',
            ]);
            await writeXlsxFile(path, { compressionLevel: 0, rows });

            const bytes = await readFile(path);
            assert.ok(bytes.length > 1024 * 1024, `only ${bytes.length} bytes`);
            assert.equal(sheetRows((await readXlsx(bytes)).sheet).length, 20000);
        });
    });

    describe('createFileWritable', () => {
        it('is a WritableStream a stream can be piped at', async () => {
            const path = join(dir, 'piped.xlsx');
            await createXlsxStream({ rows: [['a']] }).pipeTo(createFileWritable(path));
            assert.equal(sheetRows((await readXlsx(await readFile(path))).sheet).length, 1);
        });

        it('hands the failure to the write that is waiting on the file', async () => {
            const writer = createFileWritable(join(dir, 'missing', 'out.xlsx')).getWriter();
            // Past the file's own buffer in one chunk, so the write is not
            // done until `drain` — which never comes, because the open
            // failed. Nothing but the `error` event resolves this one.
            await assert.rejects(writer.write(new Uint8Array(200_000)), /ENOENT/);
        });

        it('hands a failure it already saw to the next write', async () => {
            const writer = createFileWritable(join(dir, 'missing', 'out.xlsx')).getWriter();
            // The open fails on the threadpool, so the first write is likely
            // to be buffered before the failure lands; from there on every
            // write is refused with the failure already in hand.
            for (let attempt = 0; attempt < 20; attempt++) {
                try {
                    await writer.write(new Uint8Array([1, 2, 3]));
                } catch (err) {
                    assert.match(String(err), /ENOENT/);
                    return;
                }
                await sleep(25);
            }
            assert.fail('the failed open was never reported');
        });

        it('rejects a close on a file that failed to open', async () => {
            const writable = createFileWritable(join(dir, 'missing', 'out.xlsx'));
            const writer = writable.getWriter();
            // Nothing was ever offered to it, so the failure is only waiting
            // to be handed to whichever step starts next.
            await assert.rejects(writer.close(), /ENOENT/);
        });

        it('destroys the file when the stream is aborted', async () => {
            const path = join(dir, 'aborted.xlsx');
            const writable = createFileWritable(path);
            const writer = writable.getWriter();
            await writer.write(new Uint8Array([1, 2, 3]));
            await writer.abort(new Error('changed my mind'));
            await assert.rejects(writer.write(new Uint8Array([4])));
        });

        it('takes an abort reason that is not an Error', async () => {
            const writable = createFileWritable(join(dir, 'aborted-plain.xlsx'));
            await writable.abort('no reason in particular');
        });
    });
});
