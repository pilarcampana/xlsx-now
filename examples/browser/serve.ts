// Static file server for the browser example. Built on `server4test` so the
// server is a plain object we can create, await and close in-process — the
// automated test (run-in-chromium-test.ts) imports `startExampleServer()`
// instead of spawning `node serve.js` and killing the child.
import { Server4Test } from 'server4test';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const DEFAULT_PORT = Number(process.env['PORT'] ?? 8080);

/** Path of the ESM example page, relative to the server root. */
export const EXAMPLE_PATH = '/examples/browser/';

/** Path of the UMD example page (same server, it is served from the repo root). */
export const UMD_EXAMPLE_PATH = '/examples/umd/index.html';

/**
 * Starts the example server and resolves once it is actually listening.
 * Stop it with `await server.closeServer()`.
 */
export async function startExampleServer(port: number = DEFAULT_PORT): Promise<Server4Test> {
    const server = new Server4Test({
        port,
        verbose: false,
        // Served from the repo root: index.html loads the compiled module from
        // /dist and resolves the bare "client-zip" specifier to /node_modules.
        'public-dir': resolve('.'),
        'serve-content': {
            // "" (no extension) has to be allowed for `index` to kick in on
            // directory URLs; ".ts" is only needed by devtools when following
            // source maps.
            allowedExts: ['', 'html', 'js', 'mjs', 'json', 'map', 'ts', 'ico'],
            index: ['index.html'],
        },
    });
    await server.start();
    return server;
}

// Run directly (`npm run example:browser`): serve until interrupted.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
    const server = await startExampleServer();
    console.log(`Serving ${resolve('.')} on http://localhost:${server.port}`);
    console.log(`  ESM example: http://localhost:${server.port}${EXAMPLE_PATH}`);
    console.log(`  UMD example: http://localhost:${server.port}${UMD_EXAMPLE_PATH}`);
}
