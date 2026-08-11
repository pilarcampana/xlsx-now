// Packages built by hand, for the reader to be pointed at.
//
// The round trip through this repository's own writer is one kind of test and
// it is in `readXlsx.test.ts`; this is the other kind. A file made here can
// say things the writer never writes — a shared string table, a namespace
// prefix, a row that leaves out its number, the 1904 epoch — which is most of
// what a reader has to get right about files it did not make.
import { ZipWriter } from '../../src/core/zip.js';
import { recordingSink } from './streams.js';

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** A zip of exactly these entries, in this order. */
export function zipOf(parts: Record<string, string>): Uint8Array {
    const { sink, bytes } = recordingSink();
    const zip = new ZipWriter(sink);
    for (const [name, text] of Object.entries(parts)) zip.writeEntry(name, text);
    zip.end();
    return bytes();
}

function relationships(entries: readonly string[]): string {
    return `<Relationships xmlns="${RELATIONSHIPS_NS}">${entries.join('')}</Relationships>`;
}

function relationship(id: string, type: string, target: string): string {
    return `<Relationship Id="${id}" Type="${RELATIONSHIP_TYPE}/${type}" Target="${target}"/>`;
}

export interface PackageParts {
    /** The `<sheetData>` of each sheet, in order, by the name of the sheet. */
    sheets: Record<string, string>;
    /** The whole `xl/styles.xml`, for a workbook that has one. */
    styles?: string;
    /** The `<si>` entries of the shared string table, already spelled out. */
    sharedStrings?: string;
    /** Attributes for `<workbookPr>`, which is where `date1904` lives. */
    workbookPr?: string;
    /** Written in front of every element name, to read a file that has one. */
    prefix?: string;
}

/**
 * The smallest package that is still an xlsx: the relationships, a workbook,
 * one part per sheet, and whichever of styles and shared strings were asked
 * for.
 *
 * The parts go into the archive in the order `exceljs` uses — the worksheets
 * ahead of the shared strings they point into, and the workbook last of all —
 * so that a reader that quietly depended on a friendlier order would fail
 * here rather than in someone's hands.
 */
export function xlsxPackage(parts: PackageParts): Uint8Array {
    const names = Object.keys(parts.sheets);
    const tag = (name: string): string => (parts.prefix ? `${parts.prefix}:${name}` : name);
    const xmlns = parts.prefix ? `xmlns:${parts.prefix}` : 'xmlns';

    const entries: Record<string, string> = {
        '_rels/.rels': relationships([
            relationship('rIdWorkbook', 'officeDocument', 'xl/workbook.xml'),
        ]),
    };
    names.forEach((name, index) => {
        entries[`xl/worksheets/sheet${index + 1}.xml`] =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<${tag('worksheet')} ${xmlns}="${SPREADSHEET_NS}">` +
            `<${tag('sheetData')}>${parts.sheets[name] ?? ''}</${tag('sheetData')}>` +
            `</${tag('worksheet')}>`;
    });
    if (parts.sharedStrings !== undefined) {
        entries['xl/sharedStrings.xml'] =
            `<sst xmlns="${SPREADSHEET_NS}">${parts.sharedStrings}</sst>`;
    }
    if (parts.styles !== undefined) entries['xl/styles.xml'] = parts.styles;

    const sheetRelationships = names.map((_name, index) =>
        relationship(`rIdSheet${index + 1}`, 'worksheet', `worksheets/sheet${index + 1}.xml`),
    );
    if (parts.sharedStrings !== undefined) {
        sheetRelationships.push(relationship('rIdStrings', 'sharedStrings', 'sharedStrings.xml'));
    }
    if (parts.styles !== undefined) {
        sheetRelationships.push(relationship('rIdStyles', 'styles', 'styles.xml'));
    }
    entries['xl/_rels/workbook.xml.rels'] = relationships(sheetRelationships);

    entries['xl/workbook.xml'] =
        `<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIP_TYPE}">` +
        (parts.workbookPr === undefined ? '' : `<workbookPr ${parts.workbookPr}/>`) +
        `<sheets>` +
        names
            .map(
                (name, index) =>
                    `<sheet name="${name}" sheetId="${index + 1}" r:id="rIdSheet${index + 1}"/>`,
            )
            .join('') +
        `</sheets></workbook>`;
    return zipOf(entries);
}

/** A styles part whose `cellXfs` are exactly these number formats, in order. */
export function stylesOf(formats: readonly (string | number)[]): string {
    const custom = formats.filter((format): format is string => typeof format === 'string');
    const idOf = (format: string | number): number =>
        typeof format === 'number' ? format : 164 + custom.indexOf(format);
    return (
        `<styleSheet xmlns="${SPREADSHEET_NS}">` +
        `<numFmts count="${custom.length}">` +
        custom
            .map((code, index) => `<numFmt numFmtId="${164 + index}" formatCode="${code}"/>`)
            .join('') +
        `</numFmts>` +
        // The named styles, which sit in a list of their own right before the
        // cell ones and are numbered separately from them.
        `<cellStyleXfs count="1"><xf numFmtId="9"/></cellStyleXfs>` +
        `<cellXfs count="${formats.length}">` +
        formats.map((format) => `<xf numFmtId="${idOf(format)}"/>`).join('') +
        `</cellXfs></styleSheet>`
    );
}
