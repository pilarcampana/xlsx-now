import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { sheetXmlChunks } from './sheet.js';
import { stylesXml } from './styles.js';

/**
 * Builds a styled .xlsx as a Web ReadableStream<Uint8Array>, without ever
 * holding the full workbook (or the full row set) in memory.
 *
 * Platform-agnostic on purpose: it never imports a zip library directly.
 * Callers inject `makeZip` (e.g. from the `client-zip` package) so the same
 * module runs unmodified in Node and in the browser — only the entry point
 * that supplies `makeZip` differs per platform.
 *
 * @param {object} options
 * @param {{ name: string, key?: string, pk?: boolean }[]} options.columns
 * @param {AsyncIterable<Record<string, unknown>> | Iterable<Record<string, unknown>>} options.rows
 * @param {string} [options.sheetName]
 * @param {(files: unknown, options?: unknown) => ReadableStream<Uint8Array>} options.makeZip
 * @returns {ReadableStream<Uint8Array>}
 */
export function createXlsxStream({ columns, rows, sheetName = 'Sheet1', makeZip }) {
    if (!makeZip) {
        throw new Error('createXlsxStream requires a `makeZip` implementation (e.g. from "client-zip").');
    }

    const files = [
        { name: '[Content_Types].xml', input: contentTypesXml() },
        { name: '_rels/.rels', input: rootRelsXml() },
        { name: 'xl/workbook.xml', input: workbookXml(sheetName) },
        { name: 'xl/styles.xml', input: stylesXml() },
        { name: 'xl/_rels/workbook.xml.rels', input: workbookRelsXml() },
        { name: 'xl/worksheets/sheet1.xml', input: sheetXmlChunks(columns, rows) },
    ];

    return makeZip(files);
}
