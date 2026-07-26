import { cellRef, cellXml } from './cell.js';
import { STYLE } from './styles.js';

const SHEET_HEADER =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
const SHEET_FOOTER = '</sheetData></worksheet>';

function styleFor(isHeaderRow, isPkColumn) {
    if (isHeaderRow && isPkColumn) return STYLE.PK_HEADER;
    if (isHeaderRow) return STYLE.HEADER;
    if (isPkColumn) return STYLE.PK;
    return STYLE.DEFAULT;
}

function rowXml(rowNumber, values, columns, isHeaderRow) {
    let cells = '';
    for (let i = 0; i < columns.length; i++) {
        const style = styleFor(isHeaderRow, Boolean(columns[i].pk));
        cells += cellXml(values[i], cellRef(i, rowNumber), style);
    }
    return `<row r="${rowNumber}">${cells}</row>`;
}

// Streams the worksheet XML: header, then one <row> per incoming record,
// emitted as soon as each record arrives — nothing is buffered in full.
export async function* sheetXmlChunks(columns, rows) {
    yield SHEET_HEADER;
    yield rowXml(
        1,
        columns.map((c) => c.name),
        columns,
        true,
    );

    let rowNumber = 2;
    for await (const record of rows) {
        const values = columns.map((c) => record[c.key ?? c.name]);
        yield rowXml(rowNumber, values, columns, false);
        rowNumber++;
    }

    yield SHEET_FOOTER;
}
