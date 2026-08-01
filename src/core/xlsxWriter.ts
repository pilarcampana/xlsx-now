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
 * bigger blocks. Counted in characters, not UTF-8 bytes: for the XML this
 * produces the two are the same until the data has non-Latin text, where a
 * character can reach three bytes. It is a threshold for batching, not a
 * limit anyone depends on, and no amount of rows makes it grow.
 */
const PUSH_BATCH_CHARS = 64 * 1024;

const encoder = new TextEncoder();

export interface XlsxWriterOptions {
    columns: readonly Column[];
    sheetName?: string;
    /**
     * Deflate effort, 0-9. Defaults to 6; `0` writes the parts uncompressed,
     * which is faster but leaves the file roughly ten times bigger.
     */
    compressionLevel?: CompressionLevel;
}

/**
 * Writes a styled `.xlsx`, one record at a time, handing every byte to `sink`
 * as soon as it exists. No I/O and no streams: this is the shared engine the
 * environment-specific stream classes drive, and the only thing it ever holds
 * is the batch of worksheet XML on its way to the zip.
 */
export class XlsxWriter {
    private readonly zip: ZipWriter;
    private readonly columns: readonly Column[];
    private batch: string;
    private rowNumber = FIRST_DATA_ROW;

    constructor(
        sink: (bytes: Uint8Array) => void,
        {
            columns,
            sheetName = 'Sheet1',
            compressionLevel = DEFAULT_COMPRESSION_LEVEL,
        }: XlsxWriterOptions,
    ) {
        this.columns = columns;
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
        this.batch = SHEET_HEADER + headerRowXml(columns);
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

    writeRow(record: Row): void {
        this.discardOnFailure(() => {
            this.batch += dataRowXml(this.rowNumber, record, this.columns);
            this.rowNumber++;
            if (this.batch.length >= PUSH_BATCH_CHARS) this.pushBatch();
        });
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
