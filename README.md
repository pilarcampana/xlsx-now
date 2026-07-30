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

- **`src/core/*.ts` — pure, dependency-free XLSX XML generation.** No I/O,
  no zip library import. `styles.ts` defines a small style registry
  (`DEFAULT`, `HEADER`, `PK`, `PK_HEADER`) and `sheet.ts` streams one
  `<row>` at a time from an `AsyncIterable` of records — nothing is
  buffered. This is what makes it isomorphic: it's just string generation.
- **The zip container is injected, not imported.** `createXlsxStream`
  takes a `makeZip` function as a parameter instead of importing a zip
  library directly. Both examples pass in
  [`client-zip`](https://github.com/Touffy/client-zip) (`makeZip`), which
  builds a ZIP64 archive as a `ReadableStream<Uint8Array>` without knowing
  the total size upfront — exactly the "I don't know how many rows are
  coming" case this targets. `client-zip` is documented as browser/Deno
  targeted, but since it's plain Web Streams with no DOM APIs, it also
  works from Node (verified below via `Readable.fromWeb`).

This split is what proves the isomorphism claim: the same `src/core` files,
byte-for-byte, ran in Node and in real Chromium and produced structurally
equivalent, valid `.xlsx` files (see verification below).

## TypeScript, no bundler on the main path

Everything is written in TypeScript under `strict` (plus
`noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`) and the shipped
ESM is compiled with `tsc` alone — no bundler touches it, which keeps the
"same file runs on both platforms" claim easy to check. A bundler
(Rollup) is involved in exactly one place, the optional [UMD
build](#umd-build) described below, and it consumes `tsc`'s output rather
than the TypeScript sources.

`tsconfig.json` uses `module: NodeNext`, so relative imports carry their
`.js` extension in both source and output and the emitted ESM in `dist/`
loads unchanged in Node **and** straight from a `<script type="module">` in
the browser. The one specifier a browser can't resolve on its own —
the bare `client-zip` — is mapped by an
[import map](examples/browser/index.html) instead of a build step.

The public surface is `src/core/index.ts`, which exports
`createXlsxStream` plus the types callers need (`Column`, `Row`,
`CellValue`, `MakeZip`). `MakeZip` is declared structurally, so
`client-zip`'s `makeZip` satisfies it without `src/core` ever depending on
that package.

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

It is produced by Rollup (`rollup.config.mjs`) from the JS `tsc` already
emitted, so TypeScript compilation still happens in exactly one place and the
bundler only converts module format and inlines dependencies.

**`client-zip` is inlined into the UMD bundle**, and `makeZip` becomes
optional there. That is the one deliberate difference from the ESM entry:
`src/core` stays dependency-free and still takes `makeZip` as a parameter, but
`client-zip` ships ESM only, so a UMD that required an injected `makeZip`
would be unusable from a script tag. The injection point remains open — pass
`makeZip` explicitly to use a different zip builder. The extra entry point
lives in [`src/umd/index.ts`](src/umd/index.ts); `src/core` is untouched.

```html
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

The `xlsx-now/umd` subpath is typed (`dist/src/umd/index.d.ts`) and the
package's `unpkg`/`jsdelivr` fields point at the bundle, so a CDN URL like
`https://unpkg.com/xlsx-now` serves it directly. Since the package is
`"type": "module"`, `dist/umd/package.json` marks just that directory as
`"type": "commonjs"` — that's what lets the `.js` bundle keep an extension
browsers serve happily while `require()` still parses it as CommonJS.

Verified on all three consumption paths (global via script tag, AMD `define`,
and `require()`), and the resulting files pass the same `openpyxl` validation
as the ESM output — bold headers, filled PK column, 200 data rows.

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
- **Compression.** `client-zip` stores files uncompressed (no `deflate`).
  Fine for a PoC; worth benchmarking file size vs. a compressing zip
  writer before this goes further.
