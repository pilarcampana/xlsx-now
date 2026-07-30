import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { sheetXmlChunks } from './sheet.js';
import { stylesXml } from './styles.js';
import type { Column, ForAwaitable, MakeZip, Row, ZipEntry } from './types.js';

export interface CreateXlsxStreamOptions {
    columns: readonly Column[];
    rows: ForAwaitable<Row>;
    sheetName?: string;
    /**
     * Zip builder injected by the caller (e.g. `makeZip` from `client-zip`),
     * so this module never imports a platform-specific zip library.
     */
    makeZip: MakeZip;
}

/**
 * Builds a styled .xlsx as a Web ReadableStream<Uint8Array>, without ever
 * holding the full workbook (or the full row set) in memory.
 *
 * Platform-agnostic on purpose: it never imports a zip library directly.
 * Callers inject `makeZip` (e.g. from the `client-zip` package) so the same
 * module runs unmodified in Node and in the browser — only the entry point
 * that supplies `makeZip` differs per platform.
 */
export function createXlsxStream({
    columns,
    rows,
    sheetName = 'Sheet1',
    makeZip,
}: CreateXlsxStreamOptions): ReadableStream<Uint8Array> {
    if (!makeZip) {
        throw new Error('createXlsxStream requires a `makeZip` implementation (e.g. from "client-zip").');
    }

    const files: ZipEntry[] = [
        { name: '[Content_Types].xml', input: contentTypesXml() },
        { name: '_rels/.rels', input: rootRelsXml() },
        { name: 'xl/workbook.xml', input: workbookXml(sheetName) },
        { name: 'xl/styles.xml', input: stylesXml() },
        { name: 'xl/_rels/workbook.xml.rels', input: workbookRelsXml() },
        { name: 'xl/worksheets/sheet1.xml', input: sheetXmlChunks(columns, rows) },
    ];

    return makeZip(files);
}
