// Minimal static file server for local testing of the browser example —
// no bundler needed since tsc emits plain ESM with ".js" extensions.
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize, resolve } from 'node:path';

// Served from the repo root — run via `npm run example:browser`.
const ROOT = resolve('.');
const PORT = process.env['PORT'] ?? 8080;

const TYPES: Record<string, string> = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.mjs': 'text/javascript',
    '.json': 'application/json',
    '.map': 'application/json',
    // Sources are fetched by devtools when following source maps.
    '.ts': 'text/plain',
};

createServer(async (req, res) => {
    const path = normalize(decodeURIComponent((req.url ?? '/').split('?')[0]!));
    const filePath = join(ROOT, path === '/' ? '/examples/browser/index.html' : path);
    if (!filePath.startsWith(ROOT)) {
        res.writeHead(403).end();
        return;
    }
    try {
        const body = await readFile(filePath);
        res.writeHead(200, { 'content-type': TYPES[extname(filePath)] ?? 'application/octet-stream' });
        res.end(body);
    } catch {
        res.writeHead(404).end('Not found');
    }
}).listen(PORT, () => console.log(`Serving ${ROOT} on http://localhost:${PORT}`));
