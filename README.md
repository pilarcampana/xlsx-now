# xlsx-now

XLSX fast outputs — a streaming XLSX writer with real cell styles (fonts,
fills, borders, alignment, number and date formats), formulas, column widths,
a frozen header row and as many worksheets as the stream cares to open,
designed to run **unmodified in Node and in the browser**. The stream carries
the workbook: rows, and the commands that open a sheet or spell a line out.

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

The public API is a **stream transform**: rows go in one side, one per chunk,
and the bytes of the `.xlsx` come out the other. It is meant to sit in
the middle of a pipe, not at the head of one.

It is **Web Streams throughout**, in Node as much as in the browser. Node has
had `ReadableStream`, `WritableStream` and `TransformStream` as globals since
v18, so there is no Node-flavoured variant of the writer to keep in step —
one class, one behaviour, both environments:

| export | what it is |
| --- | --- |
| `XlsxStream` | a `TransformStream<SheetInput, Uint8Array>`, for `.pipeThrough()` |
| `createXlsxStream` | a `ReadableStream` that *pulls* the rows — or the `sheets`, one source per worksheet — for sources that aren't streams |
| `XlsxWriter` | the engine under both: `writeRow(message)`, `finish()`, no streams at all |

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
split in layers:

- **XLSX XML generation — no I/O at all.** `styles.ts` is the style table:
  it takes a flat description of what a cell should look like and takes it
  apart into the four tables xlsx keeps — the number format, the font, the
  fill, the border — deduplicating each one, and hands back the index the
  cell writes as its `s`. `sheet.ts` turns one row of cells into one `<row>`.
  Plain synchronous string functions, nothing buffered. It's just string
  generation, which is half of what makes this isomorphic.
- **The columns mode, on top of that.** `columns.ts` is the whole of it: given
  the columns it returns the freeze they imply, the header row, and the
  function that reads one record by key into a row of cells. Nothing below it
  knows what a column is.
- **The widths a sheet measures for itself.** `autoWidth.ts` counts
  characters and nothing else: every cell on its way out is measured into the
  column it lands in, and what comes back is one width per column for
  `sheet.ts` to merge into the `<cols>` it was going to write anyway. It is
  the one thing here the writer cannot hand over as it goes — see
  [Columns sized by what they hold](#columns-sized-by-what-they-hold-autowidthmax).
- **The commands, alongside the rows.** `command.ts` defines the messages that
  are not rows: `#worksheet`, which the writer turns into the end of one
  worksheet part and the start of the next, and `#line`, which is a row said
  outright — with a height, a style or the hiding of a row, none of which a
  bare array has room for. See
  [Several worksheets in one stream](#several-worksheets-in-one-stream) for
  how the workbook is put together without ever knowing how many sheets are
  coming.
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

Every form takes the same static configuration in its constructor — sheet
name, compression, what the sheet freezes and, in the columns mode, the
columns — and then receives the messages. Nothing is required, and anything
the constructor takes can arrive on the stream instead.

A message is one of three things:

| message | what it is |
| --- | --- |
| `[1, 'Widget 1']` | a row of cells, where the position is the column |
| `{ id: 1, name: 'Widget 1' }` | a record, read by the sheet's `columns` |
| `{ '#worksheet': 'Summary' }` | a command: everything after it goes to a new sheet |
| `{ '#line': 'array', values: [...] }` | a command: the line said outright, options and all |

Any position of a row — and any property of a record — is either a plain
value or a cell that says more about itself: `{ v, s, f, t, col }`. See
[Styles](#styles) and [The cell that says more](#the-cell-that-says-more).

The first two are the same thing with a header on top, and they are not two
modes to choose between: a sheet with `columns` takes records *and* rows of
cells, in any order. The last one is those same two lines spelled out, for
when the line has something to say beyond its cells.

A key that starts with `#` is what makes a message a command. That is the one
reserved character in the API, and it costs nothing: a column's `name` — what
the header row shows — can be anything, `#` included, and only its `key` is
restricted. A key that starts with `#` and is not a command the writer knows
is refused by name, rather than going in as a row of blanks.

### The columns mode

`columns` declares the sheet, and every row is a record read by key.

```js
const columns = [
    { name: 'id', key: 'id', pk: true },
    { name: 'name', key: 'name' },
    { name: 'price', key: 'price' },
];
```

Each column is `name` (the header text), `key` (the property read from every
record, defaulting to `name`) and `pk`, which marks it as a primary key: pk
columns get the highlight fill, and they are also what the sheet freezes.

The sheet gets a bold header row of the column names, and every record becomes
one row. That is all this mode is — it is written *as* the rows mode below,
in `columns.ts`, and nothing underneath it knows what a column is.

A sheet with columns still takes rows of cells, which is how a separator, a
note or a totals line goes in without a column to hold it. A record on a sheet
with *no* columns is the one combination that fails, and it says so.

A record's values are cells like any other, so one value can say how it looks
without the record giving up on being read by key:

```js
{ id: 1, name: 'Widget 1', price: { v: 3.33, s: 'money' } }
```

### The rows mode

No columns, no keys: a row is an array, and the position is the column.

```js
import { createXlsxStream } from 'xlsx-now';

const xlsxStream = createXlsxStream({
    freezeRows: 1,
    freezeColumns: 1,
    rows: [
        [{ v: 'id', s: { bold: true } }, { v: 'name', s: { bold: true } }],
        [1, 'Widget 1'],
        [2, 'Widget 2', undefined, new Date()],
    ],
});
```

A position holds either a value — the same `string | number | boolean | Date |
null` the columns mode takes — or a cell that says more about itself, which is
the section below.

Three things a position can be, and they are not the same thing:

| in the array | in the sheet |
| --- | --- |
| `undefined` (or a hole, `[1, , 3]`) | no cell at all |
| `null` or `''` | an empty cell, written only if it carries a style |
| `{ s }` with no value | the styled cell — asking for the style is asking for the cell |

Rows can be as long or short as they happen to be; nothing has to line up.

### The cell that says more

`{ v, s, f, t, col }`. Every field is optional, the value included.

| field | what it is |
| --- | --- |
| `v` | the value |
| `s` | the style: a declared one by name, or one written out |
| `f` | the formula, with or without a leading `=` |
| `t` | what the cell holds, as xlsx spells it — read off `v` when it is not said |
| `col` | the column to write it in: `'J'`, or `10` for the same one |

```js
[
    { v: 1234.5, s: 'money' },                     // a declared style, by name
    { v: new Date(), s: { numFmt: 'dd/mm/yyyy' } }, // a style written out
    { v: 45, f: 'SUM(B2:B10)' },                   // a formula, and its result
    { v: '007', t: 'inlineStr' },                  // a code, not the number 7
    { v: 'Total', col: 'J' },                      // in J, not in the next one
]
```

**`f`, and the value beside it.** What xlsx stores next to a formula is the
*cached* result — what a reader shows until it recalculates. With a `v` the
cell reads right straight away; without one it is blank until something
recalculates it, which is a decision worth making on purpose.

**`t`, for when the value is not what it looks like.** `'n'`, `'b'`, `'str'`,
`'inlineStr'`, `'e'`. Left out it is read off `v`, which is right nearly
always; saying it outright is how a code that looks like a number stays text,
or a number that arrived as text goes in as a number.

`'n'` is what a `<c>` holds when it says nothing, so it is never written out —
six bytes per cell, on the type most of a sheet is made of. The two spellings
of text are the cell's own business rather than the caller's, and this is what
they mean in the file:

| type | where the text goes | when |
| --- | --- | --- |
| `inlineStr` | `<is><t>` inside the cell | text the cell holds itself |
| `str` | the `<v>`, next to the `<f>` | the cached result of a formula |

[ECMA-376][ecma376] §18.18.11 spells `str` out as *"cell containing a formula
string"*: it is not a cheaper `inlineStr`, it is the other one of the pair. A
cell with an `f` writes its text as `str` — there is no `<is>` next to an
`<f>` — and a cell without one writes it as `inlineStr`, whichever of the two
was asked for. Text with nothing around it would be the shared string table,
`t="s"`, and this writer has none.

`xml:space="preserve"` goes on the `<t>` of a string whose edges would
otherwise be trimmed, and on no other — the whitespace *inside* an element is
never XML's to touch, so a string that starts and ends with a letter pays
nothing for the attribute.

[ecma376]: https://ecma-international.org/publications-and-standards/standards/ecma-376/
[msoi]: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/17d11129-219b-4e2c-88db-45844d21e528

**`col`, and why there is no sparse form.** Columns are numbered from 1, as
the sheet shows them, and letters work in any case. Whatever follows carries
on from there, so a line that touches two far-apart columns costs two cells,
not seventy-eight:

```js
{ '#line': 'array', values: [{ v: 'Total', col: 'A' }, { v: 12, col: 'BZ' }] }
```

A line only moves forward. A `col` pointing at a column the line has already
written, or already gone past, is refused — two cells in one column is a file
Excel opens as one of them, and which one is nobody's decision to leave to it.

### Styles

A style is one flat object, and the writer takes it apart into the four tables
xlsx keeps it in:

```js
{
    font: 'Calibri', size: 11, bold: true, italic: true, underline: 'double',
    strike: true, script: 'super', color: '#003366',
    bg: '#FFE699',
    align: 'center', valign: 'middle', wrap: true, rotate: 90, indent: 1,
    shrink: true,
    numFmt: '#,##0.00',
    border: { all: 'thin', bottom: { style: 'thick', color: '#f00' } },
    locked: false, hideFormula: false,
}
```

Colours are hex — `#RGB`, `#RRGGBB`, `RRGGBB` or `AARRGGBB`, the `#` optional
— and anything that is not one is refused where it was written, not turned
into an attribute that makes the file fail to open. `numFmt` is the format
code as Excel spells it (`'yyyy-mm-dd'`, `'0.00%'`), or a number for one of
Excel's built-in formats. `border.all` covers the four sides, under whatever
a side says for itself; the diagonal is its own, with `diagonalUp` and
`diagonalDown` to say which way it runs.

**Declared once, asked for by name.** Nothing *has* to be declared — a cell
can carry a style outright — but a name keeps one look in one place, and it is
what `base` builds on:

```js
const xlsxStream = createXlsxStream({
    styles: {
        money: { numFmt: '#,##0.00', align: 'right' },
        moneyTotal: { base: 'money', bold: true, border: { top: 'thin' } },
    },
    rows: [
        [{ v: 3.33, s: 'money' }],
        [{ v: 9.99, s: 'moneyTotal' }],
        [{ v: 9.99, s: { base: 'money', color: '#c00' } }],  // the same, plus one thing
    ],
});
```

A name nobody declared is an error, not a cell that quietly comes out plain.

**Where a style can go.** Three levels, and Excel resolves them in this order:

| level | how |
| --- | --- |
| the cell | `{ v, s }` |
| the row | `{ '#line': 'array', values: [...], s }` |
| the column | `columnFormats` |

**Nothing has to be known upfront.** The table is built as the rows go by and
`xl/styles.xml` is written at the end, next to `xl/workbook.xml` and for the
same reason — the order of the entries inside the archive is nobody's business
but the central directory's. What it holds is bounded by how many *different*
styles the workbook has, not by how many rows: two cells asking for the same
thing, however each of them spelled it, get the same index.

### Dates

A date is a number in a sheet, and a number with no format is shown as the
five-digit serial it is. So a `Date` whose cell asks for no format of its own
gets one. A style that says `numFmt` is left alone — asking for a format is
how a caller says it wants that one.

A `Date` is an entry of [`types`](#types-the-workbook-knows-types) like any
other, and the two options below are what its conversion reads.

**There is no preview to write.** A date cell holds two things: the serial
number, and the id of a format. Nothing in the file says what the day *looks
like* — that is worked out by whoever opens it, every time. (A formula caches
its result and a date does not: there is nowhere in a `<c>` to put one.)

**The default is the reader's own short date.** `numFmtId` 14 and 22 are the
two built-ins the reader spells for itself: [ECMA-376][ecma376] §18.8.30 lists
them as `mm-dd-yy` and `m/d/yy h:mm`, and Microsoft's implementation notes for
that same clause ([MS-OI29500][msoi]) say Excel shows them in the short date
of the system it is running on. They are what "Short Date" in Excel's own
format menu applies. So a date written by this writer reads `15/01/2024` in
Buenos Aires and `1/15/2024` in Chicago — each user sees a date written the
way they write dates, which is the point.

**Or one format for everybody**, when the file is going somewhere that does
not have a locale — a spec, an import, an archive:

```js
createXlsxStream({ dateFormat: 'yyyy-mm-dd', rows });   // ISO, everywhere
createXlsxStream({ dateFormat: 'dd/mm/yyyy', rows });   // one country's own
```

| option | what it is | default |
| --- | --- | --- |
| `dateFormat` | a `Date` with no time of day | `14`, the built-in short date |
| `dateTimeFormat` | a `Date` that carries one | `dateFormat` + ` hh:mm:ss` |

The time of day is added to `dateFormat` on its own, so only the date is ever
asked for: an hour is written the same way everywhere and the order of a day
and a month is not. A built-in is an id and not a format code — there is
nothing to add a time to — so `dateFormat: 14` pairs with the built-in `22`,
and any other built-in has to say its `dateTimeFormat` outright rather than
have one guessed at.

That pair is where the two ways part: `22` is the reader's short date with
the time to the *minute*, since that is the built-in Excel has, while a
`dateFormat` written out gets the whole `hh:mm:ss`. The seconds are in the
value either way — this is what is shown, not what is stored — and a workbook
that wants them shown under a local date says so:
`{ dateFormat: 14, dateTimeFormat: 'dd/mm/yyyy hh:mm:ss' }`.

`autoWidthMax` measures a date by the format it will be shown in. Under a
built-in that is the reader's business, so what it measures is the widest a
short date runs to.

The serial is the **wall clock the caller reads**, not the UTC instant. A
sheet has no time zone: `new Date(2024, 0, 15)` in Buenos Aires has to come
out of the file reading `2024-01-15`, and taking `getTime()` for the serial
would write `2024-01-14 21:00`. So the date's own `getTimezoneOffset()` —
daylight saving and all — is taken off first, and what gets written is the
same reading `getFullYear()` and `getHours()` give.

### Types the workbook knows: `types`

A sheet holds four things: a number, a boolean, a string, or nothing. Anything
else has to become one of them on the way in, and `types` is where a workbook
is told how.

It is a `Map` from a class to the conversion for it. `defaultTypes` is what a
workbook uses when it is told nothing, and `withType` builds one map from
another:

```js
import { createXlsxStream, defaultTypes, withType } from 'xlsx-now';

class HourRange {
    constructor(from, to) { this.from = from; this.to = to; }
    toString() { return `${this.from} a ${this.to}`; }
}

// Once, wherever the application's own types live:
export const appTypes = withType(defaultTypes, HourRange, {
    convert: (range) => ({ v: range.toString() }),
});

// And from then on, an HourRange is a value like any other:
createXlsxStream({ columns, rows, types: appTypes });
```

The point is that nothing at the call site has to remember: an `HourRange`
that reaches any cell of that workbook, in any sheet and in any column,
converts itself. What used to be a `map` before the writer is declared once
and named once.

**A conversion returns a value, not a cell.** Only `v` is required — a string,
a number or a boolean, which is what the writer already writes. The rest is
what the value would have lost by becoming one of those:

| field | what it is |
| --- | --- |
| `v` | the value itself, as the writer already writes it |
| `t` | what the cell says it holds; read off `v` when left out |
| `numFmt` | the number format the cell falls back to |
| `width` | how many characters it *shows*, for `autoWidthMax` |

`numFmt` applies exactly as a date's does: a style that already says `numFmt`
has said what it wants and is left alone. A `t` written on the cell wins over
the one the conversion gave, since asking for a type outright is asking for
that one.

**`width` is worked out from the format when it is left out.** A conversion
that gives a `numFmt` has already said that `v` is not what the cell shows, so
measuring `v` there is not measuring imprecisely — it is measuring the wrong
thing. A duration written as a fraction of a day is the clearest case:

```js
withType(defaultTypes, Interval, {
    convert: (interval) => ({ v: interval.ms / 86400000, numFmt: '[h]:mm:ss' }),
});
```

Half an hour is `0.020833333333333332` in the file — twenty characters of a
number nobody will ever see — against the seven of the `0:30:00` the cell
shows. So a format code, which is a template of what comes out, is measured by
its own length: nine here, close enough to size a column by, and exactly right
for a `'yyyy-mm-dd'`.

It is an estimate and not a measurement. A format whose output grows with the
value — `'#,##0.00'` against a million — comes out short, and a built-in
format is an id with no code to measure at all. Both are what `width` is for:
a conversion that knows its own magnitude says so, as `dateValue` does.

**What comes in the box:**

| class | written as |
| --- | --- |
| `Date` | the serial, under [the workbook's date formats](#dates) |
| `BigInt` | a number while a cell can hold one exactly, text past that |
| `URL` | its `href` |

`Date` is one entry among the others and not a case above them, which is the
whole test of whether this generalizes: a class of your own goes in the same
way and costs the same. A conversion is handed the workbook's `DateFormats`,
so a type of yours that is a date in any sense reads the same `dateFormat`
everything else does — `dateValue` is exported for exactly that:

```js
withType(defaultTypes, Timestamp, {
    convert: (own, context) => dateValue(own.toDate(), context),
});
```

**A `BigInt` past 2<sup>53</sup> becomes text.** A cell stores a double, so
15 to 16 digits is all the precision there is: written as a number, a longer
one comes back out of the file with its last digits replaced by zeros — a
wrong value that looks like a right one. Text keeps every digit, and a caller
who would rather have the rounded number says `t: 'n'` on the cell.

**`types` replaces the map, it does not add to it.** That is what makes it one
option instead of a list, and it means a map that leaves `Date` out is a
workbook where a date is an error rather than a wrong number. It is read once,
when the writer is constructed, so nothing that happens to the map afterwards
changes what that workbook writes — and two workbooks written side by side
cannot change what the other one knows.

**A class nobody registered is refused by name.** There is no one right way to
write a `Map`, a `Set`, a nested array or a class the library has never seen,
so none of them is written out as whatever `String()` makes of it:

```
A cell is a value of a type the workbook knows, and "HourRange" is not one of
them: add it to the writer's "types", with withType(defaultTypes, ...).
```

The lookup walks the prototype chain, so `class Timestamp extends Date` needs
no entry of its own, and it is keyed on the class itself rather than on its
name — nothing here depends on a name a minifier is free to rewrite. Every
prototype is remembered the first time it is seen, so a sheet of a million
dates costs one lookup and 999,999 hits.

**An object nobody claimed** can be caught with an entry under `Object`, which
is the class every prototype chain ends at. It is consulted last, after an
object has been given the chance to be [a cell that says more](#the-cell-that-says-more):
`{ v: 1, s: 'money' }` is a cell, not a value, and stays one.

### The columns of a sheet: `columnFormats`

How wide a column is, whether it is shown, and what its cells look like
without a style of their own:

```js
const xlsxStream = createXlsxStream({
    styles: { money: { numFmt: '#,##0.00' } },
    columnFormats: {
        B: { width: 24 },
        C: { width: 12, s: 'money' },
        D: { hidden: true },
    },
    rows,
});
```

By position works too, which is the same thing said the other way:

```js
columnFormats: [{ width: 8 }, { width: 24 }, { width: 12, s: 'money' }]
```

This is the sheet's own layout and `columns` is how a record is read: they are
declared apart because they do not have to line up, and a sheet written from
arrays has no `columns` at all and can still say that its column C is twelve
characters wide.

Like `columns` and the freezes, `columnFormats` can be given in the writer
options as the workbook's default and again on a `#worksheet` command for the
sheet it opens.

The column style costs nothing per cell — Excel applies it to every cell of
the column that carries no style of its own, so it is written once in the
worksheet's `<cols>` and never stamped on a row.

### Columns sized by what they hold: `autoWidthMax`

`autoWidthMax` sizes every column of the sheet by what is written in it: the
longest cell of the column, counted in characters, is its width — up to that
many characters, which is where a column of long text stops growing.

```js
const xlsxStream = createXlsxStream({
    columns,
    autoWidthMax: 40,
    rows,
});
```

It works the same in either mode — a sheet of records is measured through the
row each one is read into, and a sheet of arrays through the cells' positions,
which is where they were going to be written anyway. The header row is
measured like any other, so a column of short values under a long name comes
out as wide as the name.

**What a width is given to.** Every column that had something written in it.
A column of blanks measured nothing and gets no width, and neither does one
nobody wrote in at all: they keep whatever `columnFormats` says about them,
and Excel's default width when it says nothing either.

**What is measured.** What the cell will show: the characters of a string, the
digits of a number, `TRUE`/`FALSE`, and — for a date — the format it gets,
which under the [default built-in](#dates) is the widest a short date runs to
(`dd/mm/yyyy`, and the same with the time after it). A formula is measured by the cached
result it carries, and not at all when it carries none: there is nothing in
the file to be wide for until a reader recalculates it. The one thing this
cannot see is a *number format*: `1234.5` is measured as the six characters it
is written as, not as the `1.234,50` a `numFmt` may show it as. A column whose
format makes its values longer is a column to give a width to outright.

**What lands in the `<col>` is not the count.** A column's `width` is measured
in multiples of the widest digit of the normal font *plus* five pixels of
padding — two of margin on each side, one for the gridline — and stored in
1/256ths ([ECMA-376][ecma376] §18.3.1.13):

```
width = Truncate([{characters} * {digit width} + {5px padding}] / {digit width} * 256) / 256
```

which is why a column autofitted to eight characters is the `8.7109375` Excel
writes, and not `8`. Ten characters of date come out as `width="10.7109375"`.
Writing the count itself is what a column exactly as wide as its longest value
looks like: the text is clipped, and a date or a number under it shows as
`##########`. The digit width is Calibri 11's, 7 pixels at 96 dpi, which is
the font a workbook has until one of its styles says otherwise.

**A width given outright wins.** `columnFormats` is where the sheet says what
it wants, and nothing measures over it — the two go together, and a format
that says everything but the width gets the measured one filled in:

```js
columnFormats: { A: { width: 3 }, D: { hidden: true } },
autoWidthMax: 40,   // A stays at 3, D is measured and stays hidden
```

**What it costs.** `<cols>` is written before the first row of the worksheet
and the widths are not known until the last one, so a sheet that measures
itself is held in memory until it closes and then goes into the archive whole.
That is the whole of the cost, and it is per sheet: a `#worksheet` command
closes the one being measured and starts the next, so a workbook of many
sheets never holds more than the one it is writing. Without `autoWidthMax` —
the default — nothing is measured, nothing is held, and the sheet goes out in
batches as it is written, exactly as before.

Like `columns`, `columnFormats` and the freezes, `autoWidthMax` can be given
in the writer options as the workbook's default and again on a `#worksheet`
command for the sheet it opens.

### Frozen rows and columns

`freezeRows` and `freezeColumns` fix that many rows at the top and columns at
the left, so they stay on screen as the sheet scrolls. Both default to 0 — and
in the columns mode, to what the columns imply, which is the section below.
Given explicitly they win, in either mode.

A freeze is a single split at one position: it always takes everything before
it along, which is why these are two counts and not a choice of which rows or
columns to fix.

### Frozen header and pk columns

The header row is always frozen, so it stays on screen as the sheet scrolls,
and so are the pk columns — but only while they are the sheet's first
columns. With `columns` as above, `id` is column A and the file opens with row
1 and column A fixed; two leading pks freeze two columns, and so on.

A freeze is a single split at one position, so a pk sitting after an ordinary
column can't be frozen without dragging every column before it along. When the
pks are mixed in among the rest, only the header row is frozen. Same when
*every* column is a pk: freezing all of them would leave nothing to scroll.

Nothing to configure: it follows from `pk` and the column order — the columns
mode works out `freezeRows: 1` and however many leading pks there are, and
hands them to the rows mode as the defaults. It costs one `<pane>` element at
the top of the worksheet, before the first row.

### Lines said outright: `#line`

A row array and a record are recognized by their shape, which is all most
lines need. `#line` is the same line said outright, and it buys two things
that shape alone cannot express:

```js
{ '#line': 'row',   values: { id: 1, name: 'Ana' } }   // read by the columns
{ '#line': 'array', values: [1, 'Ana'] }               // position is the column
{ '#line': 'empty' }                                   // a row and nothing in it
```

**Options on the row itself.** `height` (in points), `hidden` and an `s` for
the whole row have nowhere to go on a bare array — every position there is a
cell — and that is the main reason for the command:

```js
{ '#line': 'array', values: ['Total'], s: { bold: true }, height: 22 }
{ '#line': 'empty', hidden: true }
```

The row's `s` is the same style a cell takes, and it sits *under* whatever the
cells carry themselves: a bold row with one red cell is both.

**Saying which one it is.** An array *is* already unambiguous, and so is a
record — `'row'` and `'array'` exist so a generic producer can be explicit
instead of relying on what its values happen to look like.

An empty line needs no command, incidentally: a bare `[]` is a row with
nothing in it, and always was. `{ '#line': 'empty' }` is the same row with
somewhere to hang `height` or `hidden`.

### Several worksheets in one stream

A `#worksheet` message closes the sheet being written and opens a new one.
Nothing else changes: the same stream keeps carrying rows, and they land on
whichever sheet is open when they arrive.

```js
import { createXlsxStream } from 'xlsx-now';

const xlsxStream = createXlsxStream({
    columns,
    sheetName: 'Widgets',
    rows: [
        { id: 1, name: 'Widget 1', price: 3.33 },
        { id: 2, name: 'Widget 2', price: 6.66 },
        { '#worksheet': 'Summary', columns: [{ name: 'metric', pk: true }, { name: 'value' }] },
        { metric: 'rows', value: 2 },
        { metric: 'total price', value: 9.99 },
    ],
});
```

The command carries the sheet's own configuration — `columns`,
`columnFormats`, `autoWidthMax`, `freezeRows`, `freezeColumns` — and what it leaves out falls
back to the writer options,
which are the workbook's defaults. So a table split across sheets repeats
nothing:

```js
rows: [
    ...january,
    { '#worksheet': 'February' },   // same columns, same freezes, new sheet
    ...february,
]
```

`columns: []` is how a sheet opts *out* of the columns the workbook declared:
no header row, no freeze, rows of cells alone.

Sent **before any row**, a command configures the first sheet rather than
adding a second one — nothing is written until the first message arrives.
That is why `sheetName` and `columns` are optional everywhere: a stream can
declare itself entirely on the way in.

```js
rows: [
    { '#worksheet': 'Widgets', columns, freezeColumns: 1 },
    ...records,
]
```

Sheet names are made to fit what Excel accepts, rather than refused: a name it
would reject is a file that will not open, and by the time anyone finds out
the rows are long gone. So the characters Excel forbids (`\ / ? * [ ] :`, and
a leading or trailing `'`) are dropped, anything past 31 characters is cut, a
sheet that arrives with no name gets `Sheet<n>`, and a name another sheet
already took — which Excel compares without regard to case — gets a `(2)`, a
`(3)`, and so on, with the number fitted inside the same 31 characters.

#### How the workbook is assembled without knowing its sheets

Two parts of an `.xlsx` have to name every worksheet: `[Content_Types].xml`
and `xl/workbook.xml` (with its `.rels`). Neither can be written while sheets
are still arriving — and `[Content_Types].xml`, per OPC, has to be the *first*
part in the archive.

The way out is that a zip's entries can be written in any order (the central
directory comes last anyway), and that content types can be assigned by
extension:

- `[Content_Types].xml` goes out first, before any sheet exists, declaring
  `<Default Extension="xml">` as the worksheet type and overriding it for the
  only two `.xml` parts that are not worksheets, `workbook.xml` and
  `styles.xml`. However many sheets arrive, they are already typed.
- `xl/workbook.xml` and `xl/_rels/workbook.xml.rels` — the parts that list the
  sheets by name — are written **last**, once no more of them can come. So is
  `xl/styles.xml`, for the same reason: what the cells asked for is not known
  until the last of them is in.

So the archive reads `[Content_Types].xml`, `_rels/.rels`, one part per
worksheet, and then the styles and the workbook that names the sheets. Still
one pass, still nothing buffered, and a workbook of a hundred sheets costs no
more memory than a workbook of one.

### `new XlsxStream(...)` in a pipe chain

A `TransformStream`: one message per chunk in, the file's bytes out. A chunk
is a row of cells, a record, or a `#worksheet` command — `SheetInput`, the one
type every face of the writer takes.

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

### Rows that are not a stream: `createXlsxStream(...)`

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

This one pulls: a row is read only when the consumer asks for more bytes,
so memory stays flat even for a data set still being received. In Node,
`writeXlsxFile(path, options)` from `xlsx-now/node` takes the same options and
writes the file in one call.

### One source per worksheet: `sheets`

A workbook whose shape is known before its rows are — a report per sheet, a
query behind each one — has no reason to interleave `#worksheet` commands into
a single stream. `sheets` takes the worksheets spelled out instead, each with
rows of its own:

```js
await writeXlsxFile('reportes.xlsx', {
    sheets: [
        { name: 'Ventas', columns: ventasColumns, rows: db.cursor(ventasSql) },
        { name: 'Costos', columns: costosColumns, rows: db.cursor(costosSql) },
    ],
});
```

`sheets` is itself a `ForAwaitable`, so the list can be an array as above or a
source that yields one report at a time. Each sheet carries the whole of
`SheetOptions` — `columns`, `columnFormats`, `autoWidthMax`, `freezeRows`,
`freezeColumns` — and whatever it leaves out falls back to the writer options,
exactly as a `#worksheet` command does.

Nothing is buffered and nothing runs early: a sheet's `rows` is read only
while that sheet is the one being written, so the query behind the second
report starts when the first one has finished going out. `rows` and `sheets`
are two ways to say the same thing, so pass one or the other, not both.

The reason it is only here, and not on `XlsxStream`, is that a chunk of a
`TransformStream` is one value: a pipe chain cannot carry a stream of streams,
which is what the `#worksheet` command exists to encode. `sheets` is that same
encoding, written for you — it flattens to the very same messages, at no
measurable cost.

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
the bare `fflate` — is mapped by an
[import map](examples/browser/index.html) instead of a build step.

The public surface is `src/core/index.ts`, which exports `XlsxWriter`,
`XlsxStream` and `createXlsxStream` plus the types callers need (`Column`,
`Row`, `CellValue`, `SheetInput`, `WorksheetCommand`, `LineCommand`,
`XlsxSheet`, `RowOptions`, `CompressionLevel`) and the `WORKSHEET` and `LINE` keys
themselves, with the environment-specific faces under `src/node/index.ts` and
`src/browser/index.ts`.

```sh
npm run build      # tsc -> dist/ (JS + .d.ts + source maps), then the UMD bundle
npm run typecheck  # tsc --noEmit
```

## Tests

```sh
npm test         # mocha
npm run test-ci  # mocha under c8: console summary, coverage/lcov.info, coverage/*.html
```

The tests live in [`test/`](test), are written in TypeScript against the
sources in `src/` (never against `dist/`), and run on Node. The browser side
is a second stage: what runs today is what Node can run, which turns out to
be almost everything — `Blob`, `Response` and the Web Streams are all native
here, so `src/browser/index.ts` is covered too, with only the DOM around it
([`test/helpers/dom.ts`](test/helpers/dom.ts)) faked. Driving the real thing
is still `npm run example:browser:test`.

Every generated file is read back with implementations that had nothing to do
with writing it — [`yauzl`](https://github.com/thejoshwolfe/yauzl) plus Node's
own `zlib` for the container, [`exceljs`](https://github.com/exceljs/exceljs)
for the workbook — for the same reason
[`scripts/validate-xlsx.ts`](scripts/validate-xlsx.ts) does: checking with
`fflate` would only prove that the writer agrees with itself. The container
properties an OOXML consumer expects (ZIP 2.0, no ZIP64, deflate, sizes after
the data, CRC recomputed from the inflated bytes) are asserted on the tests'
own output, not just on the examples'.

### Coverage, and why it is `c8` and not `nyc`

`npm test` compiles `src/` and `test/` together into `dist-test/`
([`tsconfig.test.json`](tsconfig.test.json)) with the same settings as the
shipped build — same `module: NodeNext`, same strictness — and mocha loads
that. So the tests run on the very ESM `npm run build` emits, which for a
package whose claim is "the output loads unchanged in Node and in the
browser" is the only version worth testing.

That rules out `nyc`: it instruments through Istanbul's `require` hook, which
only ever sees CommonJS. Point it at an ES module and every file comes back
at 0% — the tests would have had to run on a CommonJS transliteration of the
sources for the coverage tool's benefit. `c8` reads V8's own coverage
instead, so there is nothing to instrument and no module format it cannot
see, and `tsc`'s source maps put the report back on the `.ts` files.

One thing about [`.c8rc.json`](.c8rc.json) is worth knowing before editing
it: `src` points at `dist-test/src`, the **compiled** output, not at `src/`.
That is what `--all` — the flag that makes a module no test ever imports show
up at 0% instead of not showing up at all — scans, and pointing it at the
TypeScript instead silently zeroes the whole report. `types.js` is excluded
for the opposite reason: it is types only, so it compiles to an empty module
that can never be covered.

Coverage stands at 100% of the lines, statements and functions of `src/`, and
one branch short of 100% of branches: the `if (!this.batch) return` guard in
`XlsxWriter.pushBatch`, which nothing driving the writer from outside can
reach (`finish()` always appends the worksheet footer before pushing).

## UMD build

The ESM output above is the primary artifact, but it can't be consumed from a
classic `<script>` tag (no module resolution) or from a `require()` call. For
those, `npm run build` also emits UMD bundles:

```
dist/umd/xlsx-now.umd.js           # global `xlsxNow`, also AMD- and CommonJS-aware
dist/umd/xlsx-now-browser.umd.js   # global `xlsxNowBrowser`: the browser helpers
```

It is produced by Rollup (`rollup.config.mjs`) from the JS `tsc` already
emitted, so TypeScript compilation still happens in exactly one place and the
bundler only joins this package's own modules into one file and converts the
module format. Dependencies are declared `external`, so `fflate` is **not**
copied into the bundle: it stays a `require('fflate')` under CommonJS, a
declared dependency under AMD, and the `fflate` global under a plain
`<script>` tag — which is the one case that needs it loaded first.

The UMD build is a straight repackaging of `src/core/index.ts` — same exports,
same signatures, no separate entry point and no API differences from the ESM
path.

There are **two** bundles because they cover different environments, and the
split mirrors the one `src/` already makes. The first one is the writer, which
runs anywhere. The second is `src/browser/index.ts` — `downloadXlsx` and
`createXlsxBlob`, which reach for the DOM and the File System Access API — and
since a UMD build is above all for the browser, leaving them out would leave
out the part a page most wants. The core counts as external there too, on the
same principle: the browser bundle is 3 kB of helpers, not a second copy of the
writer for the page to download twice.

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

Loading the helpers on top of that — after the core bundle, whose global they
read — replaces the whole `Response`/`Blob`/anchor dance with one call:

```html
<script src="node_modules/xlsx-now/dist/umd/xlsx-now-browser.umd.js"></script>
<script>
  // From a click handler: it may open the browser's save dialog.
  xlsxNowBrowser.downloadXlsx('widgets.xlsx', {
      columns: [{ name: 'id', pk: true }, { name: 'name' }],
      rows: [{ id: 1, name: 'Widget 1' }],
      sheetName: 'Widgets',
  });
</script>
```

```js
const { createXlsxStream } = require('xlsx-now/umd');     // CommonJS
const { downloadXlsx } = require('xlsx-now/umd/browser'); // browser only
```

Both subpaths are typed by reusing the declarations of the sources they
repackage (`xlsx-now/umd` reuses `dist/src/core/index.d.ts`,
`xlsx-now/umd/browser` reuses `dist/src/browser/index.d.ts`), since the
exports are the same. The package's `unpkg`/`jsdelivr` fields point at the core
bundle, so a CDN URL like
`https://unpkg.com/xlsx-now` serves it directly. Since the package is
`"type": "module"`, `dist/umd/package.json` marks just that directory as
`"type": "commonjs"` — that's what lets the `.js` bundle keep an extension
browsers serve happily while `require()` still parses it as CommonJS.

Verified on all three consumption paths (global via script tag, AMD `define`,
and `require()`), and the resulting files pass the same
[validation](#try-it) as the ESM output — bold headers, filled PK column,
frozen header row and PK column, 200 data rows. The browser bundle is verified
along with them: the UMD example page generates its file through
`xlsxNowBrowser`, so the run below exercises both bundles and the handover
between them.

```sh
npm run example:umd:node          # require() the bundle -> out/example-umd-node.xlsx
npm run example:umd:browser       # then open http://localhost:8080/examples/umd/index.html
npm run example:umd:browser:test  # same page headlessly -> out/example-umd-browser.xlsx
```

## Try it

```sh
npm install

# Each example script runs `tsc` first, so no separate build step is needed.

# Node: streams 200 simulated rows into out/example-node.xlsx, plus a
# second "Summary" sheet the same stream declares on its way out
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
the header and the data rows, non-PK columns unstyled, the header row and the
leading PK columns frozen, expected row count, and no two sheets sharing a
name) and the container side (every entry at ZIP version 2.0, no ZIP64 record, worksheet
deflated, CRC recomputed from the inflated bytes, and a local header with bit
3 set and zeroed sizes — the proof it was written without knowing the row
count).

Which columns are frozen is *not* read from the writer's configuration: the
validator counts the leading filled (PK) columns in the header it read back,
and requires the pane to match.

The checks were confirmed to bite, not just pass: a file written with
`compressionLevel: 0` fails on the deflate check, a wrong expected row count
fails, flipping one byte of compressed data fails on inflation, and a
worksheet written without its `<pane>` — or with the wrong number of frozen
columns — fails the freeze check.

## What this PoC does *not* cover yet

- **The test suite in a browser.** [The tests](#tests) run on Node, where
  `src/browser/index.ts` is covered against a faked DOM; running the same
  suite in a real browser is the second stage. Until then, the closest thing
  is `npm run example:browser:test`, which drives the example page through
  Chromium.
- **Reading `.xlsx`.** Not attempted — this is a well-solved problem
  (SheetJS, `exceljs`); no reason to build it from scratch here.
- **True OS-level streaming download in the browser without File System
  Access API.** Firefox/Safari don't support `showSaveFilePicker`, so on
  those the current fallback still generates incrementally but has to
  materialize a `Blob` before the browser's normal download flow can start.
- **`Temporal` out of the box.** `Temporal.PlainDate` and the rest are exactly
  what [`types`](#types-the-workbook-knows-types) is for, and adding them is a
  `withType` away — but they are not in `defaultTypes`, because this
  repository's Node and its `lib` setting have neither the runtime nor the
  types to write them against and test them.
- **Shared formulas and array formulas.** A cell's `f` is its own; the
  `shared`/`array` forms, where one expression covers a range, are not
  emitted. Nothing about them is ruled out by the design.
- **Merged cells, conditional formats, data validation, charts.** All of them
  are further parts or further elements of the worksheet, and none is
  attempted here.
- **Confirmation in Excel itself.** The container is now shaped the way Office
  documents it wants (ZIP 2.0, deflate), and the files read back correctly
  under `yauzl` and `exceljs` — and a two-sheet file converts cleanly through
  a headless LibreOffice Calc, which is a third reader and a stricter one
  about the package than `exceljs` is. Still, no file has been opened in a
  real Excel installation from this environment.
- **Archives above 4 GB.** No ZIP64 is emitted, so an entry over 4 GB
  uncompressed would overflow silently. Unreachable in practice — Excel stops
  at 1,048,576 rows, which is roughly 250 MB of worksheet XML — but it is a
  real limit and not a guarded one.
