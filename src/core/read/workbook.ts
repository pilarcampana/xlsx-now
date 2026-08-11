// Finding the parts, the way the package says to.
//
// An xlsx is not a zip with agreed-upon file names in it: it is an OPC
// package, where every part is reached by following a relationship from
// another part. `xl/worksheets/sheet1.xml` is a convention, not a rule — the
// workbook says which relationship each of its sheets is, and the
// relationship says which part that is. Going by the conventional names
// instead works right up until it does not, and then it silently reads the
// wrong sheet, or reads them in the wrong order.
//
// So the trail is walked properly: `_rels/.rels` says where the workbook is,
// the workbook's own rels say where its sheets are, and the workbook says
// which of them come in which order and what they are called.
import { parseXml } from './xml.js';

/** Relationship types, by the last step of the URI that names them. */
const OFFICE_DOCUMENT = 'officeDocument';
const WORKSHEET = 'worksheet';
const SHARED_STRINGS = 'sharedStrings';
const STYLES = 'styles';

/** The part the package's own relationships describe: the package itself. */
export const PACKAGE_ROOT = '';

export interface Relationship {
    /** The last step of the type URI: `worksheet`, `styles`, and so on. */
    type: string;
    /** The part it points at, as a name in the archive. */
    part: string;
}

/**
 * An attribute by its local name, prefix or no prefix.
 *
 * `r:id` on a `<sheet>` is the relationship it points at, and the `r` is
 * whatever the file bound that namespace to. Elements go through `localName`
 * in the parser; attributes are few enough to ask for one at a time.
 */
function attribute(attributes: Record<string, string>, name: string): string | undefined {
    const direct = attributes[name];
    if (direct !== undefined) return direct;
    for (const [key, value] of Object.entries(attributes)) {
        if (key.slice(key.indexOf(':') + 1) === name) return value;
    }
    return undefined;
}

/** The directory a part lives in, with its trailing slash, or `''` at the root. */
function directoryOf(part: string): string {
    return part.slice(0, part.lastIndexOf('/') + 1);
}

/** Where a part's relationships are: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`. */
export function relsFor(part: string): string {
    const directory = directoryOf(part);
    return `${directory}_rels/${part.slice(directory.length)}.rels`;
}

/**
 * A relationship target as a name in the archive.
 *
 * Relative to the part the relationships *belong to*, which is not the part
 * they are written in: `xl/_rels/workbook.xml.rels` holds the relationships
 * of `xl/workbook.xml`, and its targets are relative to `xl/`. Resolving them
 * against the `_rels/` directory they are written in instead puts every part
 * of the package one directory too deep.
 *
 * A target that starts at the root says so with a leading slash, and one is
 * allowed to climb with `..` — which a package built by moving parts around
 * does use.
 */
export function resolvePart(base: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const steps: string[] = [];
    for (const step of (directoryOf(base) + target).split('/')) {
        if (step === '.' || step === '') continue;
        if (step === '..') steps.pop();
        else steps.push(step);
    }
    return steps.join('/');
}

/**
 * The relationships of one part, by id.
 *
 * `owner` is the part they describe — `xl/workbook.xml`, or `''` for the
 * package itself, whose relationships are the `_rels/.rels` at the root.
 */
export function readRelationships(xml: string, owner: string): Map<string, Relationship> {
    const relationships = new Map<string, Relationship>();
    const part = relsFor(owner);
    parseXml(
        xml,
        {
            open(name, attributes) {
                if (name !== 'Relationship') return;
                const id = attributes['Id'];
                const type = attributes['Type'];
                const target = attributes['Target'];
                if (id === undefined || type === undefined || target === undefined) return;
                // An external target is a link to something outside the
                // package — a workbook on a network drive — and there is no
                // part here to read for it.
                if (attribute(attributes, 'TargetMode') === 'External') return;
                relationships.set(id, {
                    type: type.slice(type.lastIndexOf('/') + 1),
                    part: resolvePart(owner, target),
                });
            },
        },
        part,
    );
    return relationships;
}

/** The first relationship of a type, which is how the singular parts are found. */
export function partOfType(
    relationships: Map<string, Relationship>,
    type: string,
): string | undefined {
    for (const relationship of relationships.values()) {
        if (relationship.type === type) return relationship.part;
    }
    return undefined;
}

/** Where the workbook part is, according to the package. */
export function workbookPart(rootRels: string): string {
    const part = partOfType(readRelationships(rootRels, PACKAGE_ROOT), OFFICE_DOCUMENT);
    if (part === undefined) {
        throw new Error('This package has no workbook: nothing in _rels/.rels points at one.');
    }
    return part;
}

export interface WorkbookSheet {
    name: string;
    /** The relationship id, which is what says where the worksheet part is. */
    relationshipId: string;
}

export interface Workbook {
    /** The sheets, in the order the workbook declares them. */
    sheets: WorkbookSheet[];
    /**
     * Whether the workbook counts its days from 1904 instead of 1900 — the
     * epoch the Macintosh Excel used, still written by files that came from
     * one. It shifts every serial by 1462 days, and a reader that ignores it
     * is off by four years and a day on every date in the file.
     */
    date1904: boolean;
}

export function readWorkbook(xml: string, part: string): Workbook {
    const sheets: WorkbookSheet[] = [];
    let date1904 = false;
    parseXml(
        xml,
        {
            open(name, attributes) {
                if (name === 'workbookPr') {
                    const flag = attributes['date1904'];
                    date1904 = flag === '1' || flag === 'true';
                } else if (name === 'sheet') {
                    const relationshipId = attribute(attributes, 'id');
                    if (relationshipId === undefined) {
                        throw new Error(
                            `The workbook declares a sheet with no relationship to a worksheet part.`,
                        );
                    }
                    sheets.push({ name: attributes['name'] ?? '', relationshipId });
                }
            },
        },
        part,
    );
    return { sheets, date1904 };
}

export { OFFICE_DOCUMENT, SHARED_STRINGS, STYLES, WORKSHEET };
