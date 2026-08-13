// The reader, put together.
//
// Two layers, and the lower one is the real one: `openXlsx` opens a package,
// reads the small parts, and hands back a sheet for each worksheet whose rows
// can be walked with `for await`. `readXlsx` is the twenty lines on top that
// collect those rows into a grid, which is what most callers want and what
// none of them should have to hold if they don't.
import type { StyledCell } from '../types.js';
import { readDates, type ReadDates } from './dates.js';
import { NO_FORMATS, readNumberFormats, type NumberFormats } from './numberFormats.js';
import { bytesAccess, type RandomAccess } from './randomAccess.js';
import { readSharedStrings } from './sharedStrings.js';
import type { ReadMode, ReadModes, ReadRow, ReadValue, SheetData } from './types.js';
import {
    PACKAGE_ROOT,
    partOfType,
    readRelationships,
    readWorkbook,
    relsFor,
    SHARED_STRINGS,
    STYLES,
    workbookPart,
    type Relationship,
} from './workbook.js';
import type { CellContext, RawCell } from './worksheet.js';
import { cellValue, readRows, styledCell } from './worksheet.js';
import { decodeChunks } from './xml.js';
import { entryChunks, readCentralDirectory, readEntryText, type ZipEntry } from './zipReader.js';

/** An archive to read: bytes in hand, or anything that can be seeked in. */
export type XlsxSource = Uint8Array | RandomAccess;

export interface ReadOptions<M extends ReadMode> {
    /** What each cell comes back as. Defaults to `values`. */
    mode?: M;
    /**
     * What a date cell is built as: a `Temporal` value, a `Date` read in UTC, a
     * `Date` read locally, or the ISO text. Defaults to `temporal`, which is
     * the only one of the four that says exactly what the sheet says — see
     * `ReadDates`.
     *
     * `temporal` needs a `Temporal` in the environment, native or a polyfill,
     * and a package opened without one fails here rather than at the first
     * date it reads.
     */
    dates?: ReadDates;
}

/** One worksheet of an open package, not read yet. */
export interface XlsxSheetReader<C> {
    /** The name the workbook gives it. */
    readonly name: string;
    /** The rows, as they are parsed. Nothing is held between two of them. */
    rows(): AsyncIterable<ReadRow<C>>;
    /** The whole sheet as a grid, for when holding it is the point. */
    read(): Promise<SheetData<C>>;
}

/** An open package: its sheets, in the order the workbook declares them. */
export interface XlsxReader<C> {
    readonly sheets: readonly XlsxSheetReader<C>[];
}

/**
 * Opens a package: reads everything except the rows.
 *
 * The workbook, the relationships, the styles and the shared strings all come
 * in here, because a row cannot be read without them — which sheet a
 * worksheet is, whether a number is a date, what string a `<v>7</v>` means.
 * The worksheets themselves are not touched until someone asks for their
 * rows, and each one can be asked for on its own and in any order: the parts
 * are reached through the central directory, so there is no first sheet to
 * get past.
 */
export async function openXlsx<M extends ReadMode = 'values'>(
    source: XlsxSource,
    options: ReadOptions<M> = {},
): Promise<XlsxReader<ReadModes[M]>> {
    // Before the file is touched: an option that cannot be honoured is not
    // something to find out about with a sheet already half read.
    const dates = readDates(options.dates);
    const access = source instanceof Uint8Array ? bytesAccess(source) : source;
    const entries = await readCentralDirectory(access);

    function entry(part: string): ZipEntry {
        const found = entries.get(part);
        if (!found) throw new Error(`This package declares a part it does not carry: "${part}".`);
        return found;
    }
    const textOf = (part: string): Promise<string> => readEntryText(access, entry(part));
    /** A part named by a relationship, when the package has one of that type. */
    async function optional(
        relationships: Map<string, Relationship>,
        type: string,
    ): Promise<string | undefined> {
        const part = partOfType(relationships, type);
        return part === undefined ? undefined : textOf(part);
    }

    const workbookName = workbookPart(await textOf(relsFor(PACKAGE_ROOT)));
    const relationships = readRelationships(await textOf(relsFor(workbookName)), workbookName);
    const workbook = readWorkbook(await textOf(workbookName), workbookName);

    const sharedStringsXml = await optional(relationships, SHARED_STRINGS);
    const stylesXml = await optional(relationships, STYLES);
    const formats: NumberFormats = stylesXml === undefined ? NO_FORMATS : readNumberFormats(stylesXml);
    const context: CellContext = {
        sharedStrings: sharedStringsXml === undefined ? [] : readSharedStrings(sharedStringsXml),
        formats,
        date1904: workbook.date1904,
        dates,
    };

    // The one cast in the reader, and what it stands in for is the link
    // between the `mode` asked for and the shape it gives back — which the
    // signature states and a value cannot carry.
    const convert = (
        options.mode === 'cells'
            ? (raw: RawCell) => styledCell(raw, context)
            : (raw: RawCell) => cellValue(raw, context)
    ) as (raw: RawCell) => ReadModes[M];

    const sheets = workbook.sheets.map((sheet): XlsxSheetReader<ReadModes[M]> => {
        const relationship = relationships.get(sheet.relationshipId);
        if (!relationship) {
            throw new Error(
                `The sheet "${sheet.name}" points at the relationship ${sheet.relationshipId}, which the workbook does not have.`,
            );
        }
        const part = relationship.part;
        const rows = (): AsyncIterable<ReadRow<ReadModes[M]>> =>
            readRows(decodeChunks(entryChunks(access, entry(part))), convert, part);
        return { name: sheet.name, rows, read: () => collect(sheet.name, rows()) };
    });
    return { sheets };
}

/**
 * The rows of a sheet as the grid they make up.
 *
 * Dense in rows and ragged in columns: every row up to the last one that
 * holds anything is there, empty ones as empty arrays, and each row ends
 * where its own cells do. A row the file carries but leaves empty — a row
 * element written only to give the row a height — does not make the sheet any
 * taller, since `maxRow` is meant to say where the data ends.
 */
async function collect<C>(name: string, rows: AsyncIterable<ReadRow<C>>): Promise<SheetData<C>> {
    const cells: (C | undefined)[][] = [];
    let maxCol = 0;
    for await (const row of rows) {
        if (row.cells.length === 0) continue;
        if (row.index < 1) throw new Error(`A sheet numbers its rows from 1, and one says ${row.index}.`);
        cells[row.index - 1] = row.cells;
        if (row.cells.length > maxCol) maxCol = row.cells.length;
    }
    for (let index = 0; index < cells.length; index++) cells[index] ??= [];
    return { name, cells, maxCol, maxRow: cells.length };
}

/**
 * Every sheet of a workbook, read whole.
 *
 * ```js
 * const [first] = await readXlsx(bytes);
 * first.cells[0]?.[1]                       // what B1 holds
 *
 * const [asCells] = await readXlsx(bytes, { mode: 'cells' });
 * asCells.cells[0]?.[1]                     // { v: 45306, s: { numFmt: 14 } }
 * ```
 */
export async function readXlsx<M extends ReadMode = 'values'>(
    source: XlsxSource,
    options: ReadOptions<M> = {},
): Promise<SheetData<ReadModes[M]>[]> {
    const reader = await openXlsx(source, options);
    const sheets: SheetData<ReadModes[M]>[] = [];
    for (const sheet of reader.sheets) sheets.push(await sheet.read());
    return sheets;
}

export type { ReadMode, ReadRow, ReadValue, SheetData, StyledCell };
