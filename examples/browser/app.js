// Browser-side proof: the same package entry point as the Node example,
// imported here as a path because this demo runs without a bundler.
import { createXlsxStream } from '/src/index.js';

const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
    { name: 'in_stock', key: 'inStock' },
];

// Simulates rows trickling in from a backend NDJSON stream (one record at a
// time, with a delay) instead of a full array being ready up front.
async function* simulateUpstreamRows(count) {
    for (let i = 1; i <= count; i++) {
        yield { id: i, name: `Widget ${i}`, price: Math.round(i * 3.33 * 100) / 100, inStock: i % 5 !== 0 };
        if (i % 25 === 0) await new Promise((r) => setTimeout(r, 0));
    }
}

async function generateXlsxBlob(rowCount = 200) {
    const webStream = createXlsxStream({
        columns,
        rows: simulateUpstreamRows(rowCount),
        sheetName: 'Widgets',
    });
    // Response is a convenient built-in ReadableStream -> Blob adapter.
    return new Response(webStream).blob();
}

async function generateAndDownload() {
    const status = document.getElementById('status');

    // Best case: File System Access API streams straight to disk, so the
    // download genuinely starts before generation finishes and nothing is
    // held in memory as one big Blob.
    if (window.showSaveFilePicker) {
        status.textContent = 'Streaming directly to disk via File System Access API...';
        const handle = await window.showSaveFilePicker({ suggestedName: 'example-browser.xlsx' });
        const writable = await handle.createWritable();
        const webStream = createXlsxStream({
            columns,
            rows: simulateUpstreamRows(200),
            sheetName: 'Widgets',
        });
        await webStream.pipeTo(writable);
        status.textContent = 'Done (streamed to disk).';
        return;
    }

    // Fallback for browsers without the File System Access API (Firefox,
    // Safari): still generated incrementally, but materializes as a Blob
    // before the browser's normal download flow can start.
    status.textContent = 'File System Access API unavailable, falling back to Blob download...';
    const blob = await generateXlsxBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'example-browser.xlsx';
    a.click();
    URL.revokeObjectURL(url);
    status.textContent = 'Done (Blob fallback).';
}

document.getElementById('generate').addEventListener('click', () => {
    generateAndDownload().catch((err) => {
        console.error(err);
        document.getElementById('status').textContent = `Error: ${err.message}`;
    });
});

// Exposed for automated testing (Playwright) without needing a real user
// gesture / native save dialog.
window.__generateXlsxBlob = generateXlsxBlob;
