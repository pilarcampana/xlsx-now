// The browser face, as far as Node can take it: `Blob`, `Response` and the
// Web Streams are all native here, so only the DOM around them is faked. What
// this cannot answer — that a real browser saves the file — is the second
// stage, and `npm run example:browser:test` already drives the example page
// through Chromium.
import assert from 'node:assert/strict';
import { createXlsxBlob, downloadXlsx } from '../src/browser/index.js';
import type { Column } from '../src/core/types.js';
import { stubDom, type DomStub } from './helpers/dom.js';
import { readXlsx, sheetRows } from './helpers/zip.js';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const COLUMNS: readonly Column[] = [{ name: 'id', pk: true }, { name: 'name' }];
const RECORDS = [
    { id: 1, name: 'Ana' },
    { id: 2, name: 'Beto' },
];

async function bytesOf(blob: Blob): Promise<Buffer> {
    return Buffer.from(await blob.arrayBuffer());
}

describe('createXlsxBlob', () => {
    it('is the same file, as a Blob', async () => {
        const blob = await createXlsxBlob({ columns: COLUMNS, rows: RECORDS });
        const { names, sheet } = await readXlsx(await bytesOf(blob));
        assert.equal(names.length, 6);
        assert.equal(sheetRows(sheet).length, 3); // header + 2
    });

    it('is stamped with the .xlsx MIME type', async () => {
        const blob = await createXlsxBlob({ rows: [['a']] });
        assert.equal(blob.type, XLSX_MIME);
        assert.ok(blob.size > 0);
    });
});

describe('downloadXlsx', () => {
    let dom: DomStub;

    afterEach(() => dom.restore());

    it('streams straight to disk where the File System Access API exists', async () => {
        dom = stubDom({ showSaveFilePicker: true });

        const route = await downloadXlsx('people.xlsx', { columns: COLUMNS, rows: RECORDS });

        assert.equal(route, 'file-system-access');
        assert.equal(dom.saved?.suggestedName, 'people.xlsx');
        assert.equal(sheetRows((await readXlsx(dom.saved?.bytes ?? Buffer.alloc(0))).sheet).length, 3);
        assert.equal(dom.objectUrls.length, 0, 'a Blob was materialized anyway');
    });

    it('falls back to a Blob download where it does not', async () => {
        dom = stubDom({ showSaveFilePicker: false });

        const route = await downloadXlsx('people.xlsx', { columns: COLUMNS, rows: RECORDS });

        assert.equal(route, 'blob');
        assert.equal(dom.anchors.length, 1);
        assert.equal(dom.anchors[0]?.download, 'people.xlsx');
        assert.equal(dom.anchors[0]?.clicks, 1);
        assert.equal(dom.anchors[0]?.href, dom.objectUrls[0]?.url);
        assert.equal(dom.objectUrls[0]?.blob.type, XLSX_MIME);
        assert.equal(sheetRows((await readXlsx(await bytesOf(dom.objectUrls[0]!.blob))).sheet).length, 3);
    });

    it('revokes the object URL it created', async () => {
        dom = stubDom({ showSaveFilePicker: false });
        await downloadXlsx('people.xlsx', { rows: [['a']] });
        assert.deepEqual(dom.revokedUrls, [dom.objectUrls[0]?.url]);
    });

    it('revokes it even when the download itself fails', async () => {
        dom = stubDom({ showSaveFilePicker: false });
        const document = Reflect.get(globalThis, 'document') as { createElement: unknown };
        document.createElement = () => {
            throw new Error('detached document');
        };

        await assert.rejects(downloadXlsx('people.xlsx', { rows: [['a']] }), /detached document/);
        assert.equal(dom.revokedUrls.length, 1);
    });
});
