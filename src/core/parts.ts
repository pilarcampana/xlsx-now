import { sanitizeText } from './cell.js';

const WORKSHEET_CONTENT_TYPE =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

/** Excel's own limits on what a worksheet can be called. */
const NAME_MAX_LENGTH = 31;
const FORBIDDEN_IN_NAME = /[\\/?*[\]:]/g;
/** Excel refuses a name that starts or ends with one of these, too. */
const FORBIDDEN_AT_THE_ENDS = /^'+|'+$/g;

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
 * The name Excel will accept for what the caller asked to call the sheet.
 *
 * A name Excel refuses is a file that will not open, and by the time anyone
 * finds out the rows are long gone — so rather than refuse the name, it is
 * made to fit, the way a spreadsheet importing foreign data does: the
 * characters it forbids (`\ / ? * [ ] :`, and a leading or trailing `'`) are
 * dropped, anything past 31 characters is cut, an empty name falls back to
 * `Sheet<number>`, and a name another sheet already took — which Excel
 * compares without regard to case — gets a `(2)`, a `(3)`, and so on.
 *
 * `number` is the sheet's position, 1-based, and only names the ones that
 * arrive with nothing to be called.
 */
export function sheetName(asked: unknown, taken: readonly string[], number: number): string {
    const cleaned =
        typeof asked === 'string'
            ? asked.replace(FORBIDDEN_IN_NAME, '').replace(FORBIDDEN_AT_THE_ENDS, '').trim()
            : '';
    const wanted = (cleaned || `Sheet${number}`).slice(0, NAME_MAX_LENGTH);

    // The number has to fit inside the same 31 characters, so what it costs
    // comes off the name rather than off the end of the result.
    const used = new Set(taken.map((name) => name.toLowerCase()));
    let candidate = wanted;
    for (let n = 2; used.has(candidate.toLowerCase()); n++) {
        const suffix = ` (${n})`;
        candidate = wanted.slice(0, NAME_MAX_LENGTH - suffix.length) + suffix;
    }
    return candidate;
}
