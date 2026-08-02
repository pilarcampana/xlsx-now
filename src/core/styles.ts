import type { CellStyle } from './types.js';

// Fixed style registry for the PoC. Indices must match the <cellXfs> order
// in stylesXml() below — this is the same "s=<index>" mechanism xlsx-write-stream
// already uses for number/date formats, extended here to bold headers and a
// primary-key column fill.
//
// The table is a bitmask of the style attributes, so an index is the
// combination of the attributes a cell asked for and nothing has to be
// registered while rows are being written.
const BOLD = 1;
const HIGHLIGHT = 2;

export const STYLE = {
    DEFAULT: 0,
    BOLD,
    HIGHLIGHT,
    BOLD_HIGHLIGHT: BOLD | HIGHLIGHT,
} as const;

/** A 0-based index into styles.xml's <cellXfs>. */
export type StyleIndex = (typeof STYLE)[keyof typeof STYLE];

/** The entry in the table above that holds `style`'s combination. */
export function styleIndex(style: CellStyle | undefined): StyleIndex {
    if (!style) return STYLE.DEFAULT;
    return ((style.bold ? BOLD : 0) | (style.highlight ? HIGHLIGHT : 0)) as StyleIndex;
}

export function stylesXml(): string {
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2">' +
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
        '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
        '</fonts>' +
        '<fills count="3">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor indexed="64"/></patternFill></fill>' +
        '</fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="4">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + // 0 DEFAULT
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' + // 1 BOLD
        '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' + // 2 HIGHLIGHT
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' + // 3 BOLD_HIGHLIGHT
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'
    );
}
