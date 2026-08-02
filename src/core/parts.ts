import { sanitizeText } from './cell.js';

const WORKSHEET_CONTENT_TYPE =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

/** Excel's own limits on what a worksheet can be called. */
const NAME_MAX_LENGTH = 31;
const FORBIDDEN_IN_NAME = /[\\/?*[\]:]/;

/** The part a worksheet is written to. `number` is 1-based, as Excel numbers them. */
export function worksheetPart(number: number): string {
    return `xl/worksheets/sheet${number}.xml`;
}

/**
 * How many worksheets the file carries is not known until the last row is in,
 * so the content types cannot name them one by one: worksheets are typed by
 * extension instead, and the two parts that are not worksheets override that
 * default.
 *
 * This is what lets the part stay first in the archive — where the OPC spec
 * wants it — while the sheets are still arriving. Any future `.xml` part that
 * is not a worksheet (shared strings, doc props) needs an `Override` of its
 * own here, exactly as these two have.
 */
export function contentTypesXml(): string {
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Default Extension="xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>` +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'
    );
}

export function rootRelsXml(): string {
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    );
}

/**
 * The workbook, listing its sheets in the order they were written: sheet `i`
 * is `rId(i)`, which is the numbering `workbookRelsXml` writes.
 */
export function workbookXml(sheetNames: readonly string[]): string {
    const sheets = sheetNames
        .map(
            (name, index) =>
                `<sheet name="${sanitizeText(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
        )
        .join('');
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        `<sheets>${sheets}</sheets>` +
        '</workbook>'
    );
}

/** The worksheets, in order, and the styles under the id after the last of them. */
export function workbookRelsXml(sheetCount: number): string {
    let worksheets = '';
    for (let i = 1; i <= sheetCount; i++) {
        worksheets +=
            `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`;
    }
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        worksheets +
        `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        '</Relationships>'
    );
}

/**
 * A name Excel refuses is a file that will not open, and the writer would
 * only find that out long after the rows are gone — so every sheet name is
 * checked as its sheet opens, against the limits Excel itself applies: 1 to
 * 31 characters, none of `\ / ? * [ ] :`, and no two sheets alike (which it
 * compares without regard to case).
 */
export function checkSheetName(name: string, taken: readonly string[]): void {
    if (typeof name !== 'string' || !name) {
        throw new Error('A worksheet name cannot be empty.');
    }
    if (name.length > NAME_MAX_LENGTH) {
        throw new Error(
            `Worksheet name "${name}" is longer than the ${NAME_MAX_LENGTH} characters Excel allows.`,
        );
    }
    if (FORBIDDEN_IN_NAME.test(name)) {
        throw new Error(
            `Worksheet name "${name}" carries one of the characters Excel forbids: \\ / ? * [ ] :`,
        );
    }
    const lowered = name.toLowerCase();
    if (taken.some((used) => used.toLowerCase() === lowered)) {
        throw new Error(`Worksheet name "${name}" is already taken: no two sheets can share one.`);
    }
}
