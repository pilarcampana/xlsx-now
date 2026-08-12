// The reader, put together.
//
// Two layers, and the lower one is the real one: `openXlsx` opens a package,
// reads the small parts, and hands back a sheet for each worksheet whose rows
// can be walked with `for await`. `readXlsx` is the twenty lines on top that
// collect those rows into a grid, which is what most callers want and what
// none of them should have to hold if they don't.
import type { StyledCell } from '../types.js';
import {
    dateReaderOf,
    type DateContext,
    type DateOf,
    type DateOption,
    type DateReader,
    type DateReaderName,
} from './dates.js';
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

// The defaults here are the *widest* each parameter can be, which is not what
// `openXlsx` defaults them to: a bare `ReadOptions` is somebody naming the
// shape of the options object, and every option belongs in it. What a call
// that says nothing actually gets — `values` and `temporal` — is settled on
// the function, where there is a call to infer it from.
export interface ReadOptions<M extends ReadMode = ReadMode, D extends DateOption = DateOption> {
    /** What each cell comes back as. Defaults to `values`. */
    mode?: M;
    /**
     * What a number under a date format comes back as. Defaults to
     * `'temporal'`.
     *
     * - `'temporal'` — `Temporal.PlainDate`, `PlainDateTime` or `PlainTime`,
     *   whichever the value and its format call for. Needs a `Temporal` in
     *   the runtime, and says so when the package is opened rather than when
     *   the first date turns up.
     * - `'localDate'` — a `Date` reading the same wall clock on the machine
     *   that reads it, which is what dates came back as before there was
     *   anything to choose.
     * - `'utcDate'` — a `Date` whose UTC reading is the sheet's wall clock.
     * - `'isoString'` — `'2024-01-15'`, `'2024-01-15T10:30:00'` or
     *   `'10:30:00'`, with no zone on the end.
     * - `'serial'` — the number as the file has it, with nothing made of it.
     *
     * Or a `DateReader` of the caller's own, which is how a `DateTime` of
     * some other library gets read without this package knowing about it.
     */
    dates?: D;
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
export async function openXlsx<
    M extends ReadMode = 'values',
    D extends DateOption = 'temporal',
>(
    source: XlsxSource,
    options: ReadOptions<M, D> = {},
): Promise<XlsxReader<ReadModes<DateOf<D>>[M]>> {
    // Before the file is touched at all: whether dates can be read the way
    // they were asked for is a question about this runtime, not about this
    // workbook, and a workbook with no dates in it is not a reason for the
    // answer to change.
    const dates = dateReaderOf(options.dates);
    dates.check?.();

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

    // The one cast left in the reader, and what it stands in for is the link
    // between the `mode` asked for and the shape it gives back — which the
    // signature states and a value cannot carry.
    const convert = (
        options.mode === 'cells'
            ? (raw: RawCell) => styledCell(raw, context)
            : (raw: RawCell) => cellValue(raw, context)
    ) as (raw: RawCell) => ReadModes<DateOf<D>>[M];

    const sheets = workbook.sheets.map((sheet): XlsxSheetReader<ReadModes<DateOf<D>>[M]> => {
        const relationship = relationships.get(sheet.relationshipId);
        if (!relationship) {
            throw new Error(
                `The sheet "${sheet.name}" points at the relationship ${sheet.relationshipId}, which the workbook does not have.`,
            );
        }
        const part = relationship.part;
        const rows = (): AsyncIterable<ReadRow<ReadModes<DateOf<D>>[M]>> =>
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
 * asCells.cells[0]?.[1]                     // { v: PlainDate, s: { numFmt: 14 } }
 *
 * const [asDates] = await readXlsx(bytes, { dates: 'localDate' });
 * asDates.cells[0]?.[1]                     // a Date, as it always was
 * ```
 */
export async function readXlsx<
    M extends ReadMode = 'values',
    D extends DateOption = 'temporal',
>(
    source: XlsxSource,
    options: ReadOptions<M, D> = {},
): Promise<SheetData<ReadModes<DateOf<D>>[M]>[]> {
    const reader = await openXlsx(source, options);
    const sheets: SheetData<ReadModes<DateOf<D>>[M]>[] = [];
    for (const sheet of reader.sheets) sheets.push(await sheet.read());
    return sheets;
}

export type {
    DateContext,
    DateOf,
    DateOption,
    DateReader,
    DateReaderName,
    ReadMode,
    ReadRow,
    ReadValue,
    SheetData,
    StyledCell,
};
