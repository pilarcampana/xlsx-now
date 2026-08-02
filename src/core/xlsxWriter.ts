import { columnsMode } from './columns.js';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { cellRowXml, sheetHeaderXml, SHEET_FOOTER } from './sheet.js';
import { stylesXml } from './styles.js';
import type { CellRow, Column, Row } from './types.js';
import { DEFAULT_COMPRESSION_LEVEL, ZipWriter, type CompressionLevel } from './zip.js';

const WORKSHEET_PART = 'xl/worksheets/sheet1.xml';

/**
 * How much worksheet XML is accumulated before it is handed to the zip.
 *
 * The worksheet is built one `<row>` at a time, and deflating each row on its
 * own is both slower and slightly worse (~2% larger output) than working on
 * bigger blocks. Counted in characters, not UTF-8 bytes: for the XML this
 * produces the two are the same until the data has non-Latin text, where a
 * character can reach three bytes. It is a threshold for batching, not a
 * limit anyone depends on, and no amount of rows makes it grow.
 */
const PUSH_BATCH_CHARS = 64 * 1024;

const encoder = new TextEncoder();

export interface XlsxWriterOptions {
    /**
     * Turns on the columns mode: the sheet gets a header row of column names,
     * every row is an incoming record read by key, and the freezes below
     * default to the header row and the leading pk columns. Left out, rows are
     * arrays of cells and there is nothing to declare.
     */
    columns?: readonly Column[];
    sheetName?: string;
    /**
     * Deflate effort, 0-9. Defaults to 6; `0` writes the parts uncompressed,
     * which is faster but leaves the file roughly ten times bigger.
     */
    compressionLevel?: CompressionLevel;
    /** Rows fixed at the top of the sheet. Defaults to 0, or to 1 with `columns`. */
    freezeRows?: number;
    /**
     * Columns fixed at the left of the sheet. Defaults to 0, or with
     * `columns` to however many leading ones are pks.
     */
    freezeColumns?: number;
}

/** Which kind of row a given set of options takes. */
export type RowOf<O extends XlsxWriterOptions> = O extends { columns: readonly Column[] }
    ? Row
    : CellRow;

/**
 * Writes a styled `.xlsx`, one row at a time, handing every byte to `sink` as
 * soon as it exists. No I/O and no streams: this is the shared engine the
 * environment-specific stream classes drive, and the only thing it ever holds
 * is the batch of worksheet XML on its way to the zip.
 */
export class XlsxWriter<O extends XlsxWriterOptions = XlsxWriterOptions> {
    private readonly zip: ZipWriter;
    private readonly toCellRow: (row: RowOf<O>) => CellRow;
    private batch: string;
    private rowNumber = 1;

    constructor(sink: (bytes: Uint8Array) => void, options: O) {
        const {
            columns,
            sheetName = 'Sheet1',
            compressionLevel = DEFAULT_COMPRESSION_LEVEL,
        } = options;
        const mode = columns ? columnsMode(columns) : undefined;
        // Which of the two modes is running is the whole of the difference
        // from here on. `RowOf<O>` is what the caller is held to, and no
        // runtime check on the options narrows it — hence the two casts.
        this.toCellRow = mode
            ? (row) => mode.toCellRow(row as Row)
            : (row) => row as CellRow;
        const freeze = {
            rows: options.freezeRows ?? mode?.freeze.rows ?? 0,
            columns: options.freezeColumns ?? mode?.freeze.columns ?? 0,
        };
        this.zip = new ZipWriter(sink, compressionLevel);
        this.discardOnFailure(() => {
            this.zip.writeEntry('[Content_Types].xml', contentTypesXml());
            this.zip.writeEntry('_rels/.rels', rootRelsXml());
            this.zip.writeEntry('xl/workbook.xml', workbookXml(sheetName));
            this.zip.writeEntry('xl/styles.xml', stylesXml());
            this.zip.writeEntry('xl/_rels/workbook.xml.rels', workbookRelsXml());
            // The worksheet stays open for the whole life of the writer: it is
            // the one part whose length nobody knows yet.
            this.zip.startEntry(WORKSHEET_PART);
        });
        this.batch = sheetHeaderXml(freeze);
        // Enough columns and the header row alone fills a batch, so writing it
        // can reach the zip and fail there like any other row does.
        if (mode) this.discardOnFailure(() => this.writeCellRow(mode.headerRow));
    }

    /**
     * Any failure mid-write leaves the archive unreadable, so there is nothing
     * to salvage: drop what `fflate` is holding and let the error through to
     * whoever is driving.
     */
    private discardOnFailure(step: () => void): void {
        try {
            step();
        } catch (err) {
            this.zip.terminate();
            throw err;
        }
    }

    private pushBatch(): void {
        if (!this.batch) return;
        this.zip.push(encoder.encode(this.batch));
        this.batch = '';
    }

    private writeCellRow(row: CellRow): void {
        this.batch += cellRowXml(this.rowNumber, row);
        this.rowNumber++;
        if (this.batch.length >= PUSH_BATCH_CHARS) this.pushBatch();
    }

    writeRow(row: RowOf<O>): void {
        this.discardOnFailure(() => this.writeCellRow(this.toCellRow(row)));
    }

    /** Closes the worksheet and the archive. No row goes in after this. */
    finish(): void {
        this.discardOnFailure(() => {
            this.batch += SHEET_FOOTER;
            this.pushBatch();
            this.zip.endEntry();
            this.zip.end();
        });
    }
}
