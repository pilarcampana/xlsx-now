import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { sheetXmlChunks } from './sheet.js';
import { stylesXml } from './styles.js';
import type { Column, ForAwaitable, Row, ZipEntry } from './types.js';
import { createZipStream, DEFAULT_COMPRESSION_LEVEL, type CompressionLevel } from './zip.js';

export interface CreateXlsxStreamOptions {
    columns: readonly Column[];
    rows: ForAwaitable<Row>;
    sheetName?: string;
    /**
     * Deflate effort, 0-9. Defaults to 6; `0` writes the parts uncompressed,
     * which is faster but leaves the file roughly ten times bigger.
     */
    compressionLevel?: CompressionLevel;
}

/**
 * Builds a styled .xlsx as a Web ReadableStream<Uint8Array>, without ever
 * holding the full workbook (or the full row set) in memory.
 *
 * The same module runs unmodified in Node and in the browser: the XML is
 * plain string generation and the zip container underneath is pure JS
 * (`fflate`), so nothing here touches `fs`, `zlib` or the DOM.
 */
export function createXlsxStream({
    columns,
    rows,
    sheetName = 'Sheet1',
    compressionLevel = DEFAULT_COMPRESSION_LEVEL,
}: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    const files: ZipEntry[] = [
        { name: '[Content_Types].xml', input: contentTypesXml() },
        { name: '_rels/.rels', input: rootRelsXml() },
        { name: 'xl/workbook.xml', input: workbookXml(sheetName) },
        { name: 'xl/styles.xml', input: stylesXml() },
        { name: 'xl/_rels/workbook.xml.rels', input: workbookRelsXml() },
        { name: 'xl/worksheets/sheet1.xml', input: sheetXmlChunks(columns, rows) },
    ];

    return createZipStream(files, compressionLevel);
}
