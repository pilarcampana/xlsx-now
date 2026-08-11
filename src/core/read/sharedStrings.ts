// `xl/sharedStrings.xml`: the table a cell of type `s` points into.
//
// The writer has no such table on purpose — a string goes into the cell it
// belongs to and nowhere else, which is what lets a sheet be written without
// holding on to anything. Reading has no choice: almost every file out there
// has one, and a `<v>3</v>` in a cell of type `s` means nothing without it.
//
// It is also why the reader cannot be a single forward pass. The table is
// indexed at random from any row, and it is not required to come first in the
// archive — `exceljs` writes it *after* the worksheet that uses it.
import { parseXml } from './xml.js';

const PART_NAME = 'xl/sharedStrings.xml';

/**
 * Every `<si>` of the table, in order, as one string each.
 *
 * A `<si>` is either a run of text or several of them with their own
 * formatting — `<r><t>bold</t></r><r><t> and not</t></r>` — and this is a
 * reader of data, so the runs come back joined and the formatting does not
 * come back at all.
 *
 * What is left out is `<rPh>`: the phonetic guide a Japanese entry carries
 * alongside its text, in `<t>` elements of its own. Joining those in would
 * put the reading of a word into the middle of the sentence it belongs to.
 */
export function readSharedStrings(xml: string): string[] {
    const strings: string[] = [];
    let current: string | undefined;
    let inText = false;
    let phonetic = 0;

    parseXml(
        xml,
        {
            open(name) {
                if (name === 'si') current = '';
                else if (name === 'rPh') phonetic++;
                else if (name === 't') inText = phonetic === 0;
            },
            text(text) {
                if (inText && current !== undefined) current += text;
            },
            close(name) {
                if (name === 'si') {
                    strings.push(current ?? '');
                    current = undefined;
                } else if (name === 'rPh') phonetic--;
                else if (name === 't') inText = false;
            },
        },
        PART_NAME,
    );
    return strings;
}
