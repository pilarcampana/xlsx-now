# xlsx-now

XLSX fast outputs — a streaming XLSX writer with real cell styles (bold
headers, highlighted primary-key columns), designed to run **unmodified in
Node and in the browser**.

## Why not just fork `xlsx-write-stream`?

We evaluated [`xlsx-write-stream`](https://www.npmjs.com/package/xlsx-write-stream)
(published from [`apify/xlsx-stream`](https://github.com/apify/xlsx-stream))
as a fork base. Two findings ruled it out:

- **No path to styles.** Its README states it explicitly: *"does not support
  formatting, charts, comments and a myriad of other OOXML features."* It does
  have a partial `s="<index>"` mechanism on cells (used for number/date
  formats), but the style registry itself (`styles.xml`) is static, with a
  single non-bold font and no fills — bold headers and a PK-column fill are
  not just missing, they're out of scope by design.
- **No path to the browser.** It builds the zip container with
  [`archiver`](https://www.npmjs.com/package/archiver), which depends on
  Node's `fs`/`zlib`/streams. That's the majority of the codebase, and it's
  exactly the part that can't run client-side.

A fork only pays off when you can rebase future upstream work onto your
changes. Here, upstream is small (~60 commits, stable, narrow scope) and
isn't heading toward styles or the browser — we'd diverge on nearly
everything relevant. It was more useful as a **reference**: its
`s="<index>"` pattern for per-cell styling and its use of inline strings
(`t="inlineStr"`, no shared-strings table, which keeps a writer stateless
and streamable) are both reused here.

## Architecture

The writer is split in two layers:

- **`src/core/*.js` — pure, dependency-free XLSX XML generation.** No I/O,
  no zip library import. `styles.js` defines a small style registry
  (`DEFAULT`, `HEADER`, `PK`, `PK_HEADER`) and `sheet.js` streams one
  `<row>` at a time from an `AsyncIterable` of records — nothing is
  buffered. This is what makes it isomorphic: it's just string generation.
- **The zip container is pluggable.** `src/core/createXlsxStream.js` takes
  the zip writer as a `makeZip` parameter rather than importing one, so the
  core never pulls in a platform-specific dependency. The package entry
  point (`xlsx-now`) wires in
  [`client-zip`](https://github.com/Touffy/client-zip) by default, which
  builds a ZIP64 archive as a `ReadableStream<Uint8Array>` without knowing
  the total size upfront — exactly the "I don't know how many rows are
  coming" case this targets. `client-zip` is documented as browser/Deno
  targeted, but since it's plain Web Streams with no DOM APIs, it also
  works from Node (verified below via `Readable.fromWeb`). To swap it, pass
  your own `makeZip`, or import `xlsx-now/core` and supply one.

This split is what proves the isomorphism claim: the same `src/core` files,
byte-for-byte, ran in Node and in real Chromium and produced structurally
equivalent, valid `.xlsx` files (see verification below).

## Install

```sh
npm install xlsx-now
```

## Node usage

`createXlsxStream({ columns, rows, sheetName })` returns a Web
`ReadableStream<Uint8Array>`. In Node, convert it to a Node stream with
`Readable.fromWeb` and pipe it wherever you like — a file, an HTTP response,
etc. `rows` accepts anything iterable, sync or async, so the same function
covers both a plain in-memory array and a live async iterator.

### From an array of data

```js
import { createWriteStream } from 'node:fs';
import { Readable } from 'node:stream';
import { createXlsxStream } from 'xlsx-now';

const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
];

const rows = [
    { id: 1, name: 'Widget 1', price: 3.33 },
    { id: 2, name: 'Widget 2', price: 6.66 },
    { id: 3, name: 'Widget 3', price: 9.99 },
];

const webStream = createXlsxStream({ columns, rows, sheetName: 'Widgets' });

await new Promise((resolve, reject) => {
    Readable.fromWeb(webStream)
        .pipe(createWriteStream('widgets.xlsx'))
        .on('finish', resolve)
        .on('error', reject);
});
```

### From an async iterator (stream)

Same call, but `rows` is an async generator instead of an array — each
record is turned into a `<row>` and written out as soon as it arrives, so
nothing accumulates in memory even for a data set you're still receiving
(e.g. a DB cursor or an NDJSON response from another service).

```js
import { createWriteStream } from 'node:fs';
import { Readable } from 'node:stream';
import { createXlsxStream } from 'xlsx-now';

const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
];

async function* fetchRowsFromUpstream() {
    // e.g. reading a DB cursor, or `for await (const line of ndjsonResponse.body)`
    for (let i = 1; i <= 100_000; i++) {
        yield { id: i, name: `Widget ${i}`, price: Math.round(i * 3.33 * 100) / 100 };
    }
}

const webStream = createXlsxStream({
    columns,
    rows: fetchRowsFromUpstream(),
    sheetName: 'Widgets',
});

await new Promise((resolve, reject) => {
    Readable.fromWeb(webStream)
        .pipe(createWriteStream('widgets.xlsx'))
        .on('finish', resolve)
        .on('error', reject);
});
```

See `examples/node/write-file.mjs` for the runnable version of the
streaming example.

## Try it

```sh
npm install

# Node: streams 200 simulated rows into out/example-node.xlsx
npm run example:node

# Browser: serves examples/browser at http://localhost:8080 — click
# "Generate & download" (uses the File System Access API to stream
# straight to disk if available, otherwise falls back to a Blob download)
npm run example:browser

# Or run the browser example headlessly in the pre-installed Chromium
# and save its output to out/example-browser.xlsx, for a quick sanity check
npm run example:browser:test
```

Both examples were validated by loading the resulting files with an
independent library (`openpyxl`, Python) and confirming: valid zip/xlsx
structure, header row bold, PK column (`id`) filled in both the header and
data rows, non-PK columns unstyled.

## Current limitations

- **Reading `.xlsx`.** Not attempted — this is a well-solved problem
  (SheetJS, `exceljs`); no reason to build it from scratch here.
- **True OS-level streaming download in the browser without File System
  Access API.** Firefox/Safari don't support `showSaveFilePicker`, so on
  those the current fallback still generates incrementally but has to
  materialize a `Blob` before the browser's normal download flow can start.
- **Number/date formatting on top of styles.** `cell.js` already writes
  date values as Excel serial numbers; giving date columns a real date
  `numFmt` (the same idea as the PK/header styles) is a small, natural
  follow-up, not a redesign.
- **Compression.** `client-zip` stores files uncompressed (no `deflate`),
  so output files are larger than they need to be — worth benchmarking
  against a compressing zip writer.
