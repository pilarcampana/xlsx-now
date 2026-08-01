// Browser-side proof: byte-for-byte the same core module as the Node example,
// reached here through the browser-only helpers in src/browser. The bare
// "fflate" specifier that the core imports is resolved in the browser by the
// import map in index.html.
import { createXlsxBlob, downloadXlsx } from '../../src/browser/index.js';
import type { Column, Row } from '../../src/core/types.js';

declare global {
    interface Window {
        /** Exposed for automated testing (see run-in-chromium-test.ts). */
        __generateXlsxBlob?: (rowCount?: number) => Promise<Blob>;
    }
}

const columns: Column[] = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
];

// Simulates rows trickling in from a backend NDJSON stream (one record at a
// time, with a delay) instead of a full array being ready up front.
async function* simulateUpstreamRows(count: number): AsyncGenerator<Row> {
    for (let i = 1; i <= count; i++) {
        yield { id: i, name: `Widget ${i}`, price: Math.round(i * 3.33 * 100) / 100, inStock: i % 5 !== 0 };
        if (i % 25 === 0) await new Promise((r) => setTimeout(r, 0));
    }
}

function generateXlsxBlob(rowCount = 200): Promise<Blob> {
    return createXlsxBlob({
        columns,
        rows: simulateUpstreamRows(rowCount),
        sheetName: 'Widgets',
    });
}

async function generateAndDownload(): Promise<void> {
    const status = document.getElementById('status')!;
    status.textContent = 'Generating...';

    // `downloadXlsx` streams straight to disk through the File System Access
    // API when the browser has it (nothing held in memory as one big Blob),
    // and falls back to a Blob download where it does not.
    const route = await downloadXlsx('example-browser.xlsx', {
        columns,
        rows: simulateUpstreamRows(200),
        sheetName: 'Widgets',
    });

    status.textContent =
        route === 'file-system-access' ? 'Done (streamed to disk).' : 'Done (Blob fallback).';
}

document.getElementById('generate')!.addEventListener('click', () => {
    generateAndDownload().catch((err: unknown) => {
        console.error(err);
        document.getElementById('status')!.textContent =
            `Error: ${err instanceof Error ? err.message : String(err)}`;
    });
});

// Exposed for automated testing (Playwright) without needing a real user
// gesture / native save dialog.
window.__generateXlsxBlob = generateXlsxBlob;
