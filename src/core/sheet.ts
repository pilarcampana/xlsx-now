import { cellRef, cellXml } from './cell.js';
import { STYLE, type StyleIndex } from './styles.js';
import type { CellValue, Column, Row } from './types.js';

export const SHEET_HEADER =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
export const SHEET_FOOTER = '</sheetData></worksheet>';

/** The worksheet's first row is the header, so data records start here. */
export const FIRST_DATA_ROW = 2;

function styleFor(isHeaderRow: boolean, isPkColumn: boolean): StyleIndex {
    if (isHeaderRow && isPkColumn) return STYLE.PK_HEADER;
    if (isHeaderRow) return STYLE.HEADER;
    if (isPkColumn) return STYLE.PK;
    return STYLE.DEFAULT;
}

function rowXml(
    rowNumber: number,
    values: readonly CellValue[],
    columns: readonly Column[],
    isHeaderRow: boolean,
): string {
    let cells = '';
    for (let i = 0; i < columns.length; i++) {
        const style = styleFor(isHeaderRow, Boolean(columns[i]!.pk));
        cells += cellXml(values[i], cellRef(i, rowNumber), style);
    }
    return `<row r="${rowNumber}">${cells}</row>`;
}

/** Row 1: the column names, in the header style. */
export function headerRowXml(columns: readonly Column[]): string {
    return rowXml(
        1,
        columns.map((c) => c.name),
        columns,
        true,
    );
}

/** One `<row>` for an incoming record, reading each column by its key. */
export function dataRowXml(
    rowNumber: number,
    record: Row,
    columns: readonly Column[],
): string {
    return rowXml(
        rowNumber,
        columns.map((c) => record[c.key ?? c.name]),
        columns,
        false,
    );
}
