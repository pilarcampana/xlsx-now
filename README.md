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

The public API is a **stream transform**: records go in one side, one per
chunk, and the bytes of the `.xlsx` come out the other. It is meant to sit in
the middle of a pipe, not at the head of one.

It is **Web Streams throughout**, in Node as much as in the browser. Node has
had `ReadableStream`, `WritableStream` and `TransformStream` as globals since
v18, so there is no Node-flavoured variant of the writer to keep in step —
one class, one behaviour, both environments:

| export | what it is |
| --- | --- |
| `XlsxStream` | a `TransformStream<Row, Uint8Array>`, for `.pipeThrough()` |
| `createXlsxStream` | a `ReadableStream` that *pulls* the records, for sources that aren't streams |
| `XlsxWriter` | the engine under both: `writeRow(record)`, `finish()`, no streams at all |

`XlsxWriter` knows nothing about streams — every byte goes to a sink as soon
as it exists — and the two stream faces are a dozen lines each, because the
standard already does the work.

Leaning on the standard rather than a hand-rolled pair is the point. The paths
that matter here are the ones that break silently when written by hand, and
they come for free:

- the row source fails → the file's stream errors, so the consumer gets the
  failure instead of a truncated file;
- the destination goes away → the row source stops producing;
- the output is not being read → backpressure reaches the row source;
- the row source closes → the worksheet footer and the central directory are
  written, in order.

All four are verified, on every face, not assumed. Underneath, the writer is
split in two layers:

- **XLSX XML generation — no I/O at all.** `styles.ts` defines a small style
  registry (`DEFAULT`, `HEADER`, `PK`, `PK_HEADER`) and `sheet.ts` turns one
  record into one `<row>` — plain synchronous string functions, nothing
  buffered. It's just string generation, which is half of what makes this
  isomorphic.
- **The zip container — `fflate`, wrapped in `zip.ts`.** See
  [The zip container](#the-zip-container) for why that library. `zip.ts` is a
  thin adapter, not a zip implementation: `ZipWriter` takes bytes in and hands
  whatever `fflate` produces straight to a sink. Both sides are push-based, so
  there is no queue in between. Every byte of the archive format — local
  headers, CRCs, data descriptors, central directory — is written by `fflate`.

Both halves are pure JS with no `fs`, `zlib` or DOM anywhere, which is what
proves the isomorphism claim: the same `src/core` files, byte-for-byte, ran in
Node and in real Chromium and produced structurally equivalent, valid `.xlsx`
files (see verification below).

Anything that *can't* be isomorphic lives outside `core`, one folder per
environment, and ships as its own entry point:

| module | contents |
| --- | --- |
| `xlsx-now` (`src/core`) | `XlsxStream`, `createXlsxStream`, `XlsxWriter`, the types — runs in both |
| `xlsx-now/node` (`src/node`) | `createFileWritable`, `writeXlsxFile` |
| `xlsx-now/browser` (`src/browser`) | `downloadXlsx`, `createXlsxBlob` |

Neither of those two contains a writer: they only supply the *destination*
the standard has no answer for on that platform — a file on Node, the save
dialog or a `Blob` in the browser.

## The zip container

An `.xlsx` is a zip archive, but not any zip archive will do. This project
needs four things at once, and the fourth one is what rules out most of the
field:

1. **Deflate compression.** The XML is extremely repetitive; see the
   [benchmark](#compression-benchmark) — it compresses to about a tenth.
2. **ZIP 2.0, never ZIP64.** Office rejects ZIP64 containers.
3. **Streaming without knowing the size upfront.** The whole point of the
   project is not knowing how many rows are coming. That forces *data
   descriptors*: the sizes and CRC of each entry are written after its data
   instead of in the header.
4. **The same code in Node and in the browser.** `archiver`, `yazl` and
   `zip-stream` all depend on Node's `zlib`, and `jszip` buffers the whole
   archive before emitting anything.

[`fflate`](https://github.com/101arrowz/fflate) satisfies all four: pure JS,
no dependencies, ~8 kB, and its `Zip`/`ZipDeflate` classes emit output as data
is pushed in. Verified in the generated files: entries declare version 2.0,
compression method deflate, and bit 3 of the general-purpose flag set with
zeroed sizes in the local header.

The previous container, `client-zip`, failed (1) and (2) — its own README says
*"MS Office documents must be stored using ZIP version 2.0, use client-zip^1
to generate those"*. Files produced before this change declared version 4.5
(ZIP64) on every entry and stored the parts uncompressed.

`compressionLevel` (0-9, default 6) is an option on every form of the writer;
`0` stores the parts uncompressed, which is the old behaviour.

## Usage

Every form takes the same static configuration in its constructor — columns,
sheet name, compression — and then receives the records. `columns` is the only
required option.

```js
const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
];
```

### `new XlsxStream(...)` in a pipe chain

A `TransformStream<Row, Uint8Array>`: one record per chunk in, the file's
bytes out.

```js
import { XlsxStream } from 'xlsx-now';

await rows                                         // ReadableStream<Row>
    .pipeThrough(new XlsxStream({ columns, sheetName: 'Widgets' }))
    .pipeTo(destination);                          // WritableStream<Uint8Array>
```

The same call in Node, with an NDJSON file at one end and an `.xlsx` at the
other. Every link is a `TransformStream`, so the chain itself is portable —
only the two ends are platform-specific:

```js
import { openAsBlob } from 'node:fs';
import { XlsxStream } from 'xlsx-now';
import { createFileWritable } from 'xlsx-now/node';

await (await openAsBlob('widgets.ndjson')).stream()
    .pipeThrough(new TextDecoderStream())
    .pipeThrough(new LineSplitStream())            // ~10 lines, see the example
    .pipeThrough(new JsonParseStream())
    .pipeThrough(new XlsxStream({ columns, sheetName: 'Widgets' }))
    .pipeTo(createFileWritable('widgets.xlsx'));
```

[`examples/node/write-file.ts`](examples/node/write-file.ts) is exactly this,
runnable, splitter and parser included.

### Records that are not a stream: `createXlsxStream(...)`

An array, a generator, a database cursor. `rows` accepts anything iterable,
sync or async, and the result is the finished file as a
`ReadableStream<Uint8Array>`:

```js
import { createXlsxStream } from 'xlsx-now';

const xlsxStream = createXlsxStream({
    columns,
    rows: [
        { id: 1, name: 'Widget 1', price: 3.33 },
        { id: 2, name: 'Widget 2', price: 6.66 },
    ],
    sheetName: 'Widgets',
});
```

This one pulls: a record is read only when the consumer asks for more bytes,
so memory stays flat even for a data set still being received. In Node,
`writeXlsxFile(path, options)` from `xlsx-now/node` takes the same options and
writes the file in one call.

### A note on `createFileWritable`

Node has no native web-stream writer to a file, so `xlsx-now/node` bridges
`fs.WriteStream` — the only Node stream anywhere in the package. It does
**not** use Node's own `Writable.toWeb` for that, because it does not carry
the file's backpressure tightly enough: writing a million rows uncompressed
(247 MB, where the writer outruns the disk by the widest margin) peaks at
348 MB of RSS through `Writable.toWeb`, against 111 MB straight into a web
sink. Waiting on `drain`, which is the whole of the difference, brings it back
to 109 MB.

### Browser

`xlsx-now/browser` saves the generated file, streaming straight to disk
through the File System Access API where it exists and falling back to a
`Blob` download where it does not. The return value says which of the two ran.

```js
import { downloadXlsx } from 'xlsx-now/browser';

const route = await downloadXlsx('widgets.xlsx', {
    columns,
    rows: fetchRowsFromUpstream(),
    sheetName: 'Widgets',
});
// route === 'file-system-access' | 'blob'
```

It has to be called from a user gesture, since the File System Access API
opens a native save dialog. `createXlsxBlob(options)` is there for when the
`Blob` itself is what's wanted.

## Compression benchmark

[`examples/node/benchmark.ts`](examples/node/benchmark.ts) writes the same
data at each compression level and reports size, wall time and peak process
memory. Node 22, 1,000,000 rows × 5 columns (`ROWS=1000000 npm run benchmark`):

| level | file size | vs. stored | time | peak RSS |
| --- | --- | --- | --- | --- |
| 0 (stored) | 247.6 MB | — | 3.9 s | 118 MB |
| 1 | 38.2 MB | 15.4% | 6.9 s | 134 MB |
| 6 (default) | 28.4 MB | 11.5% | 12.8 s | 148 MB |
| 9 | 28.5 MB | 11.5% | 14.7 s | 153 MB |

Level 9 costs 15% more time than level 6 for 0.1% less size, so 6 is the
default. Level 1 is the option worth knowing about: a third of the time for a
file still 6.5× smaller than stored.

Peak RSS barely moves with the row count — 200,000 rows peak at 123 MB against
1,000,000 rows at 148 MB — which is the streaming claim measured rather than
asserted. Most of that figure is the Node baseline, not the workbook.

## TypeScript, no bundler

Everything is written in TypeScript under `strict` (plus
`noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`) and the shipped
ESM is compiled with `tsc` alone — no bundler touches it, which keeps the
"same file runs on both platforms" claim easy to check. There is no bundler
in the repository at all: the optional [UMD build](#umd-build) described
below is a second `tsc` invocation over the same sources, not a bundling
step.

`tsconfig.json` uses `module: NodeNext`, so relative imports carry their
`.js` extension in both source and output and the emitted ESM in `dist/`
loads unchanged in Node **and** straight from a `<script type="module">` in
the browser. The one specifier a browser can't resolve on its own —
the bare `fflate` — is mapped by an
[import map](examples/browser/index.html) instead of a build step.

The public surface is `src/core/index.ts`, which exports `XlsxWriter`,
`XlsxStream` and `createXlsxStream` plus the types callers need (`Column`,
`Row`, `CellValue`, `CompressionLevel`), with the environment-specific faces
under `src/node/index.ts` and `src/browser/index.ts`.

```sh
npm run build      # tsc -> dist/ (JS + .d.ts + source maps), then the UMD bundle
npm run typecheck  # tsc --noEmit
```

## UMD build

The ESM output above is the primary artifact, but it can't be consumed from a
classic `<script>` tag (no module resolution) or from a `require()` call. For
those, `npm run build` also emits a UMD bundle:

```
dist/umd/xlsx-now.umd.js   # global `xlsxNow`, also AMD- and CommonJS-aware
```

No bundler produces it. `tsc -p tsconfig.umd.json` compiles the same
`src/core` sources a second time with `module: amd` and `outFile`, which is
`tsc`'s own way of emitting a whole module graph as a single file: every
module becomes a `define(...)` call, they come out in dependency order, and
bare specifiers — here only `fflate` — are left as unresolved named
dependencies. `scripts/build-umd.ts` then splices that output into the
hand-written envelope in [`src/umd/wrapper.js`](src/umd/wrapper.js), which
supplies the UMD detection and a ~20-line AMD runtime, and shifts the source
map down by the lines it prepended.

So `fflate` is *not* inlined: the bundle asks its host for it, the same way
the ESM build does. Under CommonJS that is a real `require('fflate')`, under
AMD a declared dependency, and from a `<script>` tag it is `window.fflate`,
which means fflate's own UMD build has to be loaded first (the bundle throws
a clear error if it is not).

The UMD build is a straight repackaging of `src/core/index.ts` — same exports,
same signatures, no separate entry point and no API differences from the ESM
path.

```html
<script src="https://unpkg.com/fflate"></script>
<script src="node_modules/xlsx-now/dist/umd/xlsx-now.umd.js"></script>
<script>
  const stream = xlsxNow.createXlsxStream({
      columns: [{ name: 'id', pk: true }, { name: 'name' }],
      rows: [{ id: 1, name: 'Widget 1' }],
      sheetName: 'Widgets',
  });
</script>
```

```js
const { createXlsxStream } = require('xlsx-now/umd'); // CommonJS
```

The `xlsx-now/umd` subpath is typed (it reuses `dist/src/core/index.d.ts`,
since the exports are the same) and the
package's `unpkg`/`jsdelivr` fields point at the bundle, so a CDN URL like
`https://unpkg.com/xlsx-now` serves it directly. Since the package is
`"type": "module"`, `dist/umd/package.json` marks just that directory as
`"type": "commonjs"` — that's what lets the `.js` bundle keep an extension
browsers serve happily while `require()` still parses it as CommonJS.

Verified on all three consumption paths (global via script tag, AMD `define`,
and `require()`), and the resulting files pass the same
[validation](#try-it) as the ESM output — bold headers, filled PK column,
200 data rows.

```sh
npm run example:umd:node          # require() the bundle -> out/example-umd-node.xlsx
npm run example:umd:browser       # then open http://localhost:8080/examples/umd/index.html
npm run example:umd:browser:test  # same page headlessly -> out/example-umd-browser.xlsx
```

## Try it

```sh
npm install

# Each example script runs `tsc` first, so no separate build step is needed.

# Node: streams 200 simulated rows into out/example-node.xlsx
npm run example:node

# Browser: serves the repo at http://localhost:8080/examples/browser/ —
# click "Generate & download" (uses the File System Access API to stream
# straight to disk if available, otherwise falls back to a Blob download)
npm run example:browser

# Or run the browser example headlessly in the pre-installed Chromium
# and save its output to out/example-browser.xlsx, for a quick sanity check
npm run example:browser:test

# Size/time/memory per compression level (ROWS=1000000 for the big run)
npm run benchmark
```

The static server behind every browser script (ESM and UMD alike — it serves
the repo root, so one server covers both pages) is
[`server4test`](https://www.npmjs.com/package/server4test), used
programmatically: `examples/browser/serve.ts` exports
`startExampleServer(port)`, which returns a `Server4Test` whose `start()`
only resolves once the port is really listening. The Chromium check imports
that function instead of spawning `node serve.js`, so there is no child
process, no fixed `sleep` waiting for the port, and shutdown is just
`await server.closeServer()` in a `finally`.

Every generated file is checked by
[`scripts/validate-xlsx.ts`](scripts/validate-xlsx.ts):

```sh
npm run validate -- out/example-node.xlsx 201
```

It reads the file back with implementations that had nothing to do with
writing it — [`yauzl`](https://github.com/thejoshwolfe/yauzl) plus Node's own
`zlib` for the container, [`exceljs`](https://github.com/exceljs/exceljs) for
the workbook. Validating with `fflate` would only prove that the writer agrees
with itself.

It asserts the workbook side (header row bold, PK column `id` filled in both
the header and the data rows, non-PK columns unstyled, expected row count) and
the container side (every entry at ZIP version 2.0, no ZIP64 record, worksheet
deflated, CRC recomputed from the inflated bytes, and a local header with bit
3 set and zeroed sizes — the proof it was written without knowing the row
count).

The checks were confirmed to bite, not just pass: a file written with
`compressionLevel: 0` fails on the deflate check, a wrong expected row count
fails, and flipping one byte of compressed data fails on inflation.

## What this PoC does *not* cover yet

- **Reading `.xlsx`.** Not attempted — this is a well-solved problem
  (SheetJS, `exceljs`); no reason to build it from scratch here.
- **True OS-level streaming download in the browser without File System
  Access API.** Firefox/Safari don't support `showSaveFilePicker`, so on
  those the current fallback still generates incrementally but has to
  materialize a `Blob` before the browser's normal download flow can start.
- **Number/date formatting on top of styles.** `cell.ts` already writes
  date values as Excel serial numbers; giving date columns a real date
  `numFmt` (the same idea as the PK/header styles) is a small, natural
  follow-up, not a redesign.
- **Confirmation in Excel itself.** The container is now shaped the way Office
  documents it wants (ZIP 2.0, deflate), and the files read back correctly
  under `yauzl` and `exceljs`, but no file has been opened in a real Excel
  installation from this environment.
- **Archives above 4 GB.** No ZIP64 is emitted, so an entry over 4 GB
  uncompressed would overflow silently. Unreachable in practice — Excel stops
  at 1,048,576 rows, which is roughly 250 MB of worksheet XML — but it is a
  real limit and not a guarded one.
