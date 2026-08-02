// Browser-only helpers. These reach for the DOM and the File System Access
// API, which is exactly why they live outside src/core — core has to keep
// loading unchanged in Node.
import { createXlsxStream, type CreateXlsxStreamOptions } from '../core/createXlsxStream.js';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Not in the DOM lib's `Window`, and absent in Firefox and Safari. */
interface WindowWithSaveFilePicker {
    showSaveFilePicker?: (options?: { suggestedName?: string }) => Promise<FileSystemFileHandle>;
}

/** How `downloadXlsx` ended up saving the file. */
export type DownloadRoute = 'file-system-access' | 'blob';

/**
 * The generated `.xlsx` as a `Blob`. Generation is still incremental, but the
 * whole file ends up in memory — use `downloadXlsx`, which avoids that when
 * the browser allows it.
 */
export async function createXlsxBlob(options: CreateXlsxStreamOptions): Promise<Blob> {
    // Response is a convenient built-in ReadableStream -> Blob adapter.
    const blob = await new Response(createXlsxStream(options)).blob();
    // `slice` is the way to stamp the MIME type without copying the bytes.
    return blob.slice(0, blob.size, XLSX_MIME);
}

/**
 * Saves the generated `.xlsx` under `filename`, streaming straight to disk
 * through the File System Access API when it is available (nothing is held in
 * memory, and the save dialog opens before generation finishes) and falling
 * back to a `Blob` download where it is not. The returned value says which of
 * the two ran.
 *
 * Must be called from a user gesture: the File System Access API opens a
 * native save dialog.
 */
export async function downloadXlsx(
    filename: string,
    options: CreateXlsxStreamOptions,
): Promise<DownloadRoute> {
    const showSaveFilePicker = (window as WindowWithSaveFilePicker).showSaveFilePicker;

    if (showSaveFilePicker) {
        const handle = await showSaveFilePicker.call(window, { suggestedName: filename });
        const writable = await handle.createWritable();
        await createXlsxStream(options).pipeTo(writable);
        return 'file-system-access';
    }

    // Firefox and Safari: generated incrementally all the same, but it has to
    // materialize as a Blob before the browser's normal download flow starts.
    const url = URL.createObjectURL(await createXlsxBlob(options));
    try {
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = filename;
        anchor.click();
    } finally {
        URL.revokeObjectURL(url);
    }
    return 'blob';
}
