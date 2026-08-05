import { WORKSHEET, type SheetInput, type SheetOptions } from './command.js';
import type { ForAwaitable } from './types.js';
import { XlsxWriter, type XlsxWriterOptions } from './xlsxWriter.js';

/**
 * One worksheet as a thing of its own: what it is called, whatever it decides
 * for itself, and the messages that fill it. It is what `sheets` is made of,
 * for a workbook whose shape is known before its rows are — a report per
 * sheet, a query behind each one.
 *
 * ```js
 * { name: 'Ventas', columns, rows: cursor }
 * ```
 *
 * `rows` is a source of its own, so nothing has to be in hand when the sheet
 * is described: the query behind it runs when the sheet's turn comes.
 */
export type XlsxSheet = SheetOptions & {
    /** Made to fit what Excel accepts; see `sheetName` in `parts.ts`. */
    name: string;
    /** This sheet's messages: rows of cells, records, and `#line` commands. */
    rows: ForAwaitable<SheetInput>;
};

/**
 * What the workbook is written from: one flat source of messages, or one
 * sheet at a time. The two are the same thing said differently — `sheets` is
 * flattened into the messages `rows` would have carried — and asking for both
 * is a mistake rather than a merge, so the type takes one or the other.
 */
export type CreateXlsxStreamOptions = XlsxWriterOptions &
    (
        | {
              /** The messages of the workbook: rows of cells, records, and commands. */
              rows: ForAwaitable<SheetInput>;
              sheets?: undefined;
          }
        | {
              /** The workbook a sheet at a time, each with its own rows. */
              sheets: ForAwaitable<XlsxSheet>;
              rows?: undefined;
          }
    );

/** One iterator for both kinds of source; `await` on a sync result is free. */
function iterate<T>(rows: ForAwaitable<T>): AsyncIterator<T> | Iterator<T> {
    return Symbol.asyncIterator in rows
        ? rows[Symbol.asyncIterator]()
        : rows[Symbol.iterator]();
}

/**
 * The sheets, as the messages the writer already takes: the command that
 * opens each one, then its rows. This is the whole of what `sheets` adds —
 * the nesting is undone here, and nothing downstream knows it existed.
 *
 * One sheet at a time is not a choice: the archive holds a single entry open,
 * so a sheet cannot start before the one before it has ended. Driving the
 * sheets from a `for await` is what makes that structural instead of a rule
 * the caller has to keep.
 */
async function* flatten(sheets: ForAwaitable<XlsxSheet>): AsyncGenerator<SheetInput> {
    for await (const sheet of sheets) {
        const { name, rows, ...options } = sheet;
        // The command carries whatever the sheet decided for itself; what it
        // leaves out falls back to the writer options, as any command does.
        yield { ...options, [WORKSHEET]: name };
        yield* rows;
    }
}

/** The messages of the workbook, however the caller chose to describe them. */
function messages(options: CreateXlsxStreamOptions): ForAwaitable<SheetInput> {
    const { rows, sheets } = options;
    if (sheets) {
        if (rows) {
            throw new Error(
                'Both "rows" and "sheets" were given: they are two ways to say the same ' +
                    'thing, so pass one or the other.',
            );
        }
        return flatten(sheets);
    }
    if (!rows) {
        throw new Error(
            'Nothing to write: pass "rows" with the messages of the workbook, or ' +
                '"sheets" with a source per worksheet.',
        );
    }
    return rows;
}

/**
 * The source form of the writer, for records that are not already a stream:
 * an array, a generator, a database cursor. Returns the finished file as a
 * Web `ReadableStream<Uint8Array>`, ready to be piped at a file, an HTTP
 * response or a `Blob`.
 *
 * `rows` accepts anything iterable, sync or async, and carries the whole
 * workbook: rows of cells, records read by the sheet's columns, and the
 * `{ '#worksheet': name }` commands that open one sheet after another. The
 * stream pulls: messages are read only when the consumer asks for more bytes,
 * which is what keeps memory flat however many of them are coming.
 *
 * `sheets` says the same thing with the worksheets spelled out instead of
 * commanded — one `XlsxSheet` per sheet, each with rows of its own:
 *
 * ```js
 * createXlsxStream({ sheets: reports })   // { name, columns, rows } each
 * ```
 *
 * It pulls exactly the same way, one sheet after another, and a sheet's rows
 * are read only while that sheet is the one being written.
 */
export function createXlsxStream(options: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    const iterator = iterate(messages(options));
    let writer!: XlsxWriter;
    let emitted = false;

    return new ReadableStream<Uint8Array>({
        start(controller) {
            // `rows` rides along in the options the writer gets; it reads the
            // ones it knows and this is the only place the extra one exists.
            writer = new XlsxWriter((bytes) => {
                emitted = true;
                controller.enqueue(bytes);
            }, options);
        },

        async pull(controller) {
            // Read records until the writer actually produces something, so
            // one `pull` always makes progress instead of spinning on the
            // rows that are still accumulating into the current batch.
            emitted = false;
            while (!emitted) {
                const next = await iterator.next();
                if (next.done) {
                    writer.finish();
                    controller.close();
                    return;
                }
                writer.writeRow(next.value);
            }
        },

        async cancel(reason) {
            await iterator.return?.(reason);
        },
    });
}
