/**
 * UMD entry point (bundled by Rollup into `dist/umd/xlsx-now.umd.js`).
 *
 * `src/core` stays dependency-free on purpose: it takes `makeZip` as a
 * parameter instead of importing a zip library. That works well for ESM
 * consumers, but a `<script src="...umd.js">` consumer has no module
 * resolution and `client-zip` ships ESM only — so a UMD that still required
 * an injected `makeZip` would be unusable from a plain script tag.
 *
 * This entry therefore bundles `client-zip` and defaults `makeZip` to it,
 * while keeping the injection point open for callers that want a different
 * zip builder. `src/core` itself is untouched and still imports nothing.
 */
import { makeZip } from 'client-zip';
import { createXlsxStream as createXlsxStreamCore } from '../core/createXlsxStream.js';
import type { CreateXlsxStreamOptions } from '../core/createXlsxStream.js';
import type { MakeZip } from '../core/types.js';

export { STYLE, stylesXml, type StyleIndex } from '../core/styles.js';
export type { CellValue, Column, ForAwaitable, MakeZip, Row, ZipEntry } from '../core/types.js';

/** The bundled zip builder used when the caller doesn't inject one. */
export const defaultMakeZip: MakeZip = makeZip;

export interface CreateXlsxStreamUmdOptions extends Omit<CreateXlsxStreamOptions, 'makeZip'> {
    /** Defaults to the bundled `client-zip` builder. */
    makeZip?: MakeZip;
}

/**
 * Same as the core `createXlsxStream`, except `makeZip` is optional: the
 * bundled `client-zip` builder is used when it is omitted.
 */
export function createXlsxStream(options: CreateXlsxStreamUmdOptions): ReadableStream<Uint8Array> {
    return createXlsxStreamCore({ ...options, makeZip: options.makeZip ?? makeZip });
}
