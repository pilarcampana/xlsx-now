import type { CellValue } from './types.js';
import type { StyleIndex } from './styles.js';

// Days between 1900-01-01 and 1970-01-01 (Excel's epoch quirk on Windows).
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86400000;

export function sanitizeText(value: unknown): string {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// 0-based column index -> spreadsheet column letters ("A", "B", ..., "AA", ...).
export function columnLetters(index: number): string {
    let n = index + 1;
    let letters = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

export function cellRef(colIndex: number, rowNumber: number): string {
    return `${columnLetters(colIndex)}${rowNumber}`;
}

// Renders a single <c> element. `styleIndex` is a 0-based index into
// styles.xml's <cellXfs>, or 0 (falsy) for the default style.
export function cellXml(value: CellValue, ref: string, styleIndex: StyleIndex): string {
    const s = styleIndex ? ` s="${styleIndex}"` : '';

    if (value === null || value === undefined || value === '') {
        return styleIndex ? `<c r="${ref}"${s}/>` : '';
    }
    if (value instanceof Date) {
        const officeTimestamp = value.getTime() / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
        return `<c r="${ref}" t="n"${s}><v>${officeTimestamp}</v></c>`;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
        return `<c r="${ref}" t="n"${s}><v>${value}</v></c>`;
    }
    if (typeof value === 'boolean') {
        return `<c r="${ref}" t="b"${s}><v>${value ? 1 : 0}</v></c>`;
    }
    // Inline strings (not shared strings): keeps the writer stateless/streamable.
    return `<c r="${ref}" t="inlineStr"${s}><is><t xml:space="preserve">${sanitizeText(value)}</t></is></c>`;
}
