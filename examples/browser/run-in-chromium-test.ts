// Ad-hoc verification: launches the pre-installed Chromium, runs the exact
// same compiled core module in a real browser JS engine, and saves the
// resulting Blob to disk for the same openpyxl validation used on the Node
// output.
import { chromium } from 'playwright-core';
import { writeFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { startExampleServer, EXAMPLE_PATH } from './serve.js';

const CHROMIUM_PATH = process.env['CHROMIUM_PATH'] ?? '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
const PORT = 8091;

// Both browser examples expose the same `window.__generateXlsxBlob` hook, so
// this harness drives either one: the ESM page (default) or the UMD page.
const PAGE_PATH = process.env['PAGE_PATH'] ?? EXAMPLE_PATH;
const OUT_FILE = process.env['OUT_FILE'] ?? 'out/example-browser.xlsx';

// Paths are resolved from the repo root — run via `npm run example:browser:test`.
// No child process: `start()` resolves when the port is really listening, so
// there is nothing to sleep for and nothing to kill afterwards.
const server = await startExampleServer(PORT);

try {
    const browser = await chromium.launch({ executablePath: CHROMIUM_PATH, args: ['--no-sandbox'] });
    const page = await browser.newPage();
    page.on('console', (msg) => console.log('[browser]', msg.text()));
    page.on('pageerror', (err) => console.error('[browser error]', err));

    await page.goto(`http://localhost:${PORT}${PAGE_PATH}`);
    const base64 = await page.evaluate(async () => {
        const blob = await window.__generateXlsxBlob!(200);
        const buf = await blob.arrayBuffer();
        let binary = '';
        const bytes = new Uint8Array(buf);
        for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]!);
        return btoa(binary);
    });

    await mkdir(resolve('out'), { recursive: true });
    const outPath = resolve(OUT_FILE);
    await writeFile(outPath, Buffer.from(base64, 'base64'));
    console.log(`Wrote ${outPath}`);

    await browser.close();
} finally {
    await server.closeServer();
}
