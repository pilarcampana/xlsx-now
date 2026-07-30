// Ad-hoc verification: launches the pre-installed Chromium, runs the exact
// same compiled core module in a real browser JS engine, and saves the
// resulting Blob to disk for the same openpyxl validation used on the Node
// output.
import { chromium } from 'playwright-core';
import { spawn } from 'node:child_process';
import { writeFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const CHROMIUM_PATH = process.env['CHROMIUM_PATH'] ?? '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';

// Paths are resolved from the repo root — run via `npm run example:browser:test`.
const server = spawn('node', ['dist/examples/browser/serve.js'], {
    stdio: 'inherit',
    env: { ...process.env, PORT: '8091' },
});
await new Promise((r) => setTimeout(r, 500));

try {
    const browser = await chromium.launch({ executablePath: CHROMIUM_PATH, args: ['--no-sandbox'] });
    const page = await browser.newPage();
    page.on('console', (msg) => console.log('[browser]', msg.text()));
    page.on('pageerror', (err) => console.error('[browser error]', err));

    await page.goto('http://localhost:8091/');
    const base64 = await page.evaluate(async () => {
        const blob = await window.__generateXlsxBlob!(200);
        const buf = await blob.arrayBuffer();
        let binary = '';
        const bytes = new Uint8Array(buf);
        for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]!);
        return btoa(binary);
    });

    await mkdir(resolve('out'), { recursive: true });
    const outPath = resolve('out/example-browser.xlsx');
    await writeFile(outPath, Buffer.from(base64, 'base64'));
    console.log(`Wrote ${outPath}`);

    await browser.close();
} finally {
    server.kill();
}
