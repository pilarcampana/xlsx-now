import { makeZip } from 'client-zip';
import { createXlsxStream as createXlsxStreamWithZip } from './core/createXlsxStream.js';

export { STYLE } from './core/styles.js';

/**
 * Builds a styled .xlsx as a Web ReadableStream<Uint8Array>, streaming rows
 * out as they arrive instead of buffering the whole workbook.
 *
 * Works unchanged in Node and in the browser.
 *
 * @param {object} options
 * @param {{ name: string, key?: string, pk?: boolean }[]} options.columns
 * @param {AsyncIterable<Record<string, unknown>> | Iterable<Record<string, unknown>>} options.rows
 * @param {string} [options.sheetName]
 * @param {Function} [options.makeZip] Override the zip writer (defaults to client-zip).
 * @returns {ReadableStream<Uint8Array>}
 */
export function createXlsxStream({ columns, rows, sheetName, makeZip: customMakeZip }) {
    return createXlsxStreamWithZip({
        columns,
        rows,
        sheetName,
        makeZip: customMakeZip ?? makeZip,
    });
}
