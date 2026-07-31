import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { dataRowXml, FIRST_DATA_ROW, headerRowXml, SHEET_FOOTER, SHEET_HEADER } from './sheet.js';
import { stylesXml } from './styles.js';
import type { Column, Row } from './types.js';
import { DEFAULT_COMPRESSION_LEVEL, ZipWriter, type CompressionLevel } from './zip.js';

const WORKSHEET_PART = 'xl/worksheets/sheet1.xml';

/**
 * How much worksheet XML is accumulated before it is handed to the zip.
 *
 * The worksheet is built one `<row>` at a time, and deflating each row on its
 * own is both slower and slightly worse (~2% larger output) than working on
 * bigger blocks. Batching does not buffer the workbook: only this much text
 * is ever held, no matter how many rows go through.
 */
const PUSH_BATCH_BYTES = 64 * 1024;

export interface XlsxStreamOptions {
    columns: readonly Column[];
    sheetName?: string;
    /**
     * Deflate effort, 0-9. Defaults to 6; `0` writes the parts uncompressed,
     * which is faster but leaves the file roughly ten times bigger.
     */
    compressionLevel?: CompressionLevel;
}

/**
 * Builds the transformer behind `XlsxStream`. Everything it needs lives in
 * this closure, so the class constructor is a single `super(...)` call.
 */
function xlsxTransformer({
    columns,
    sheetName = 'Sheet1',
    compressionLevel = DEFAULT_COMPRESSION_LEVEL,
}: XlsxStreamOptions): Transformer<readonly Row[], Uint8Array> {
    const encoder = new TextEncoder();
    // Assigned by `start`, which the stream always runs before `transform`
    // and `flush`.
    let zip!: ZipWriter;
    let batch = '';
    let rowNumber = FIRST_DATA_ROW;

    function pushBatch(): void {
        if (!batch) return;
        zip.push(encoder.encode(batch));
        batch = '';
    }

    /**
     * The archive is left unreadable by any failure mid-write, so there is
     * nothing to salvage: drop what `fflate` is holding and let the error
     * error the stream, which is what the consumer sees.
     */
    function discardOnFailure(step: () => void): void {
        try {
            step();
        } catch (err) {
            zip.terminate();
            throw err;
        }
    }

    return {
        start(controller) {
            zip = new ZipWriter((bytes) => controller.enqueue(bytes), compressionLevel);
            discardOnFailure(() => {
                zip.writeEntry('[Content_Types].xml', contentTypesXml());
                zip.writeEntry('_rels/.rels', rootRelsXml());
                zip.writeEntry('xl/workbook.xml', workbookXml(sheetName));
                zip.writeEntry('xl/styles.xml', stylesXml());
                zip.writeEntry('xl/_rels/workbook.xml.rels', workbookRelsXml());
                // The worksheet stays open for the whole life of the stream:
                // it is the one part whose length nobody knows yet.
                zip.startEntry(WORKSHEET_PART);
            });
            batch = SHEET_HEADER + headerRowXml(columns);
        },

        transform(rows) {
            discardOnFailure(() => {
                for (const record of rows) {
                    batch += dataRowXml(rowNumber, record, columns);
                    rowNumber++;
                    if (batch.length >= PUSH_BATCH_BYTES) pushBatch();
                }
            });
        },

        flush() {
            discardOnFailure(() => {
                batch += SHEET_FOOTER;
                pushBatch();
                zip.endEntry();
                zip.end();
            });
        },
    };
}

/**
 * A styled `.xlsx` as a `TransformStream`: batches of records go in the
 * writable side, the bytes of the file come out the readable side. Pipe rows
 * through it and pipe the result wherever the file has to go.
 *
 * ```js
 * rowBatches.pipeThrough(new XlsxStream({ columns })).pipeTo(destination)
 * ```
 *
 * Chunks are `Row[]` rather than single records, because that is the shape
 * row sources actually have (a database cursor hands over a page at a time);
 * a source that produces one row at a time writes `[row]`.
 *
 * Nothing is buffered: the file is emitted as the rows arrive, and the
 * standard's own backpressure — readable side full, writable side not ready —
 * is what stops rows from being consumed faster than they can be written out.
 *
 * The same module runs unmodified in Node and in the browser: the XML is
 * plain string generation and the zip container underneath is pure JS
 * (`fflate`), so nothing here touches `fs`, `zlib` or the DOM.
 */
export class XlsxStream extends TransformStream<readonly Row[], Uint8Array> {
    constructor(options: XlsxStreamOptions) {
        super(xlsxTransformer(options));
    }
}
