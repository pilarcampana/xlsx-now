import { WidthMeter } from './autoWidth.js';
import { writeDates, type WriteDates } from './cell.js';
import {
    checkRecord,
    isLineCommand,
    isWorksheetCommand,
    lineCells,
    lineRecord,
    noColumnsError,
    WORKSHEET,
    type LineCommand,
    type SheetInput,
    type SheetOptions,
} from './command.js';
import { columnsMode, type ColumnsMode } from './columns.js';
import { MergeTable } from './merges.js';
import {
    contentTypesXml,
    rootRelsXml,
    workbookRelsXml,
    workbookXml,
    worksheetPart,
    sheetName,
} from './parts.js';
import {
    cellRowXml,
    columnStyles,
    sheetFooterXml,
    sheetHeaderXml,
    type ColumnFormats,
    type Freeze,
    type RowOptions,
} from './sheet.js';
import {
    DateFormats,
    StyleTable,
    type DateFormatOptions,
    type StyleRef,
    type StyleSpec,
} from './styles.js';
import type { CellRow, Row } from './types.js';
import { defaultTypes, ValueTypes, type TypeMap } from './valueTypes.js';
import { DEFAULT_COMPRESSION_LEVEL, ZipWriter, type CompressionLevel } from './zip.js';

/**
 * How much worksheet XML is accumulated before it is handed to the zip.
 *
 * The worksheet is built one `<row>` at a time, and deflating each row on its
 * own is both slower and slightly worse (~2% larger output) than working on
 * bigger blocks. Counted in characters, not UTF-8 bytes: for the XML this
 * produces the two are the same until the data has non-Latin text, where a
 * character can reach three bytes. It is a threshold for batching, not a
 * limit anyone depends on, and no amount of rows makes it grow.
 */
const PUSH_BATCH_CHARS = 64 * 1024;

const encoder = new TextEncoder();

export interface XlsxWriterOptions extends SheetOptions, DateFormatOptions {
    /**
     * Name of the first worksheet; defaults to `Sheet1`. Every other sheet is
     * named by the `#worksheet` command that opens it — and so is the first
     * one, when a command arrives before any row. Whatever the name, it is
     * made to fit what Excel accepts; see `sheetName` in `parts.ts`.
     */
    sheetName?: string;
    /**
     * Styles to reuse by name, so a cell can ask for one with `s: 'money'`
     * instead of writing it out again. Nothing has to be declared here — a
     * cell can carry a style outright, and the table is built as the rows go
     * by either way — but a name is what keeps one look in one place, and it
     * is what a `base` builds on.
     */
    styles?: Readonly<Record<string, StyleSpec>>;
    /**
     * Deflate effort, 0-9. Defaults to 6; `0` writes the parts uncompressed,
     * which is faster but leaves the file roughly ten times bigger.
     */
    compressionLevel?: CompressionLevel;
    /**
     * Every type this workbook can write a value of, keyed by the class it is
     * an instance of. Defaults to `defaultTypes`, and **replaces** it whole
     * rather than adding to it — a map that leaves `Date` out is a workbook
     * where a date is an error, which is the loud half of being able to say
     * exactly what a workbook knows.
     *
     * `withType` is how one is built from another:
     *
     * ```js
     * const appTypes = withType(defaultTypes, HourRange, {
     *     convert: (range) => ({ v: range.toString() }),
     * });
     * ```
     *
     * It is read once, here, so nothing that happens to the map afterwards
     * changes what this workbook writes.
     */
    types?: TypeMap;
    /**
     * Which clock a `Date` is read by on its way into a cell: `local`, the
     * default, is the one the caller is looking at — what `getFullYear()` and
     * `getHours()` give — and `utc` is what `getUTCHours()` gives.
     *
     * A sheet has no time zone, so writing an instant down means picking the
     * clock that reads it, and only the caller knows which one their dates
     * came from: a `Date` built from a local calendar is `local`, and one that
     * came from an ISO text with a `Z` on it, or from a database that stores
     * instants, is `utc`.
     *
     * It is the only thing this decides. A `Temporal.PlainDate`, a
     * `PlainDateTime` and a `PlainTime` are wall clocks already and go in as
     * they read, whatever this says.
     */
    dates?: WriteDates;
}

/**
 * Writes a styled `.xlsx`, one message at a time, handing every byte to
 * `sink` as soon as it exists. No I/O and no streams: this is the shared
 * engine the environment-specific stream classes drive, and the only thing it
 * ever holds is the batch of worksheet XML on its way to the zip.
 *
 * A message is a row — an array of cells, or a record read by the sheet's
 * columns — or a `#worksheet` command, which closes the sheet being written
 * and opens the next one. That is what decides the order of the parts:
 * everything that has to name the sheets is written at the end, once no more
 * of them can arrive.
 */
export class XlsxWriter {
    private readonly zip: ZipWriter;
    /** What every sheet starts from, before its own command overrides it. */
    private readonly defaults: XlsxWriterOptions;
    /** The workbook's styles, filled in as its cells ask for things. */
    private readonly styles: StyleTable;
    /**
     * What this workbook can write a value of, indexed for the lookup every
     * cell that is not already a native value goes through. Built once from
     * the `types` map, and the only thing that knows a `Date` from a number.
     */
    private readonly types: ValueTypes;
    /** The sheets so far, in order; the last of them may still be open. */
    private readonly sheetNames: string[] = [];
    /** The columns of the sheet being written, if it has any. */
    private mode: ColumnsMode | undefined;
    /**
     * What the sheet's cells measure into. Every sheet has one; a sheet with
     * no `autoWidthMax` gets a meter that measures nothing.
     */
    private widths = new WidthMeter(undefined);
    /**
     * The style each column of the sheet gives its cells. Held for the rows
     * to be written against: `<col style>` reaches the cells that are not in
     * the file, and every cell that is has to carry the style itself.
     */
    private columns: readonly (StyleRef | undefined)[] = [];
    /**
     * The merged ranges of the sheet being written. Held to the end of the
     * sheet because that is where they go: `<mergeCells>` comes after
     * `<sheetData>`, so the rows stream out and the ranges wait for the
     * footer.
     */
    private merges = new MergeTable();
    /**
     * What the header of the sheet being written will be made of. Opening a
     * sheet does not write it: `<cols>` carries widths the cells may still be
     * measuring, so the header is put together — and the worksheet part
     * started — at the first moment the sheet's bytes have to go out, which is
     * `pushBatch`.
     */
    private pendingHeader: { freeze: Freeze; columnFormats: ColumnFormats | undefined } | undefined;
    private open = false;
    private batch = '';
    private rowNumber = 1;

    constructor(sink: (bytes: Uint8Array) => void, options: XlsxWriterOptions = {}) {
        this.defaults = options;
        // What a date falls back to is the workbook's, and it is what the
        // types are handed: one `dateFormat` for every type that is a date in
        // any sense, not one per class that happens to be one.
        this.types = new ValueTypes(options.types ?? defaultTypes(), {
            dates: new DateFormats(options),
            clock: writeDates(options.dates),
        });
        this.styles = new StyleTable(options.styles);
        this.zip = new ZipWriter(sink, options.compressionLevel ?? DEFAULT_COMPRESSION_LEVEL);
        this.discardOnFailure(() => {
            // The parts that depend on nothing, and so can go out before
            // anything has arrived. The first worksheet waits for the first
            // message instead: it may be a command, and then that command is
            // what configures it.
            this.zip.writeEntry('[Content_Types].xml', contentTypesXml());
            this.zip.writeEntry('_rels/.rels', rootRelsXml());
        });
    }

    /**
     * Any failure mid-write leaves the archive unreadable, so there is nothing
     * to salvage: drop what `fflate` is holding and let the error through to
     * whoever is driving.
     */
    private discardOnFailure(step: () => void): void {
        try {
            step();
        } catch (err) {
            this.zip.terminate();
            throw err;
        }
    }

    /**
     * The batch, out to the zip — and, the first time a sheet gets here, the
     * header and the worksheet part to put it in. The part is started at this
     * point and not when the sheet was opened, so that everything the header
     * says is settled by the time it is written: `<cols>` goes before the
     * first row and carries widths the cells may have been measuring.
     */
    private pushBatch(): void {
        const pending = this.pendingHeader;
        if (pending) {
            this.pendingHeader = undefined;
            // The worksheet stays open until the next command or `finish`: it
            // is the one part whose length nobody knows yet.
            this.zip.startEntry(worksheetPart(this.sheetNames.length));
            const widths = this.widths.columnWidths();
            this.batch =
                sheetHeaderXml(pending.freeze, this.styles, pending.columnFormats, widths) +
                this.batch;
        }
        this.zip.push(encoder.encode(this.batch));
        this.batch = '';
    }

    private writeCellRow(row: CellRow, options?: RowOptions): void {
        this.batch += cellRowXml(
            this.rowNumber,
            row,
            this.styles,
            this.types,
            options,
            this.widths,
            this.columns,
            this.merges,
        );
        this.rowNumber++;
        // A sheet that is measuring itself has nowhere to push to: its `<cols>`
        // is written from rows that have not arrived yet, so it waits for its
        // last one. Any other sheet goes out as it is written.
        if (!this.widths.measures && this.batch.length >= PUSH_BATCH_CHARS) this.pushBatch();
    }

    /**
     * Opens a worksheet, and writes its header row when it has columns.
     * `sheet` is the command that opened it, or the writer options for the
     * first one; either way, whatever it leaves out falls back to the
     * options, and then to what the columns imply.
     */
    private openSheet(asked: unknown, sheet: SheetOptions): void {
        const name = sheetName(asked, this.sheetNames, this.sheetNames.length + 1);
        const columns = sheet.columns ?? this.defaults.columns;
        const columnFormats = sheet.columnFormats ?? this.defaults.columnFormats;
        const autoWidthMax = sheet.autoWidthMax ?? this.defaults.autoWidthMax;
        // A header row with no column in it is nobody's intention, so an
        // empty list reads as the rows mode — which is how a sheet opts out
        // of the columns the workbook declared.
        const mode = columns?.length ? columnsMode(columns, this.types) : undefined;
        const freeze = {
            rows: sheet.freezeRows ?? this.defaults.freezeRows ?? mode?.freeze.rows ?? 0,
            columns:
                sheet.freezeColumns ?? this.defaults.freezeColumns ?? mode?.freeze.columns ?? 0,
        };

        this.sheetNames.push(name);
        this.widths = new WidthMeter(autoWidthMax);
        this.merges = new MergeTable();
        this.columns = columnStyles(columnFormats);
        this.pendingHeader = { freeze, columnFormats };
        this.batch = '';
        this.rowNumber = 1;
        this.mode = mode;
        this.open = true;
        // Enough columns and the header row alone fills a batch, so writing it
        // can reach the zip and fail there like any other row does.
        if (mode) this.writeCellRow(mode.headerRow);
    }

    private closeSheet(): void {
        // A `rowSpan` reaching past the last row of the sheet is a range
        // Excel repairs the file to be rid of, and it is a row the caller
        // meant to write and did not.
        const unfinished = this.merges.unfinishedAt(this.rowNumber);
        if (unfinished) {
            throw new Error(
                `The merge "${unfinished.ref}" needs a row ${unfinished.through} and the sheet ` +
                    `ends at row ${this.rowNumber - 1}: a merge cannot reach past the last row ` +
                    'of its sheet.',
            );
        }
        // Whatever is left of the sheet, header included when nothing of it
        // has gone out yet — a sheet with no row in it is still a sheet.
        this.batch += sheetFooterXml(this.merges);
        this.pushBatch();
        this.zip.endEntry();
        this.open = false;
    }

    /** The first sheet, as the writer options alone describe it. */
    private openFirstSheet(): void {
        this.openSheet(this.defaults.sheetName, this.defaults);
    }

    /** One record, read by the sheet's columns — the one line that needs them. */
    private recordCells(record: Row): CellRow {
        if (!this.mode) throw noColumnsError();
        return this.mode.toCellRow(record);
    }

    /**
     * A record is read by the sheet's columns, an array is already the row it
     * will be written as. The two travel together: which one a message is
     * says nothing about what the next one has to be.
     */
    private autodetectedCells(input: CellRow | Row): CellRow {
        if (Array.isArray(input)) return input as CellRow;
        const record = input as Row;
        // What claims no command has to be a record — and a key that starts
        // with `#` says it meant to claim one.
        checkRecord(record);
        return this.recordCells(record);
    }

    /**
     * A line said outright: its cells as the kind describes them, and the
     * command itself as what the row asks for past them — a `LineCommand`
     * *is* a `RowOptions`, so there is nothing to pick out of it.
     */
    private writeLineCommand(command: LineCommand): void {
        const cells = lineCells(command) ?? this.recordCells(lineRecord(command));
        this.writeCellRow(cells, command);
    }

    /** One message: a row of cells, a record, or a command. */
    writeRow(input: SheetInput): void {
        this.discardOnFailure(() => {
            if (isWorksheetCommand(input)) {
                // Before any row, the command *is* the first sheet rather
                // than a second one.
                if (this.open) this.closeSheet();
                this.openSheet(input[WORKSHEET], input);
                return;
            }
            if (!this.open) this.openFirstSheet();
            if (isLineCommand(input)) this.writeLineCommand(input);
            else this.writeCellRow(this.autodetectedCells(input));
        });
    }

    /** Closes the last worksheet and the archive. No message goes in after this. */
    finish(): void {
        this.discardOnFailure(() => {
            // A workbook with no sheet in it is not a workbook, so a writer
            // nobody gave anything to still closes with one empty sheet.
            if (!this.open) this.openFirstSheet();
            this.closeSheet();
            // Now — and only now — every sheet is known, and so is every
            // style any of their cells asked for. The order of the entries
            // inside the archive is nobody's business but the central
            // directory's, which is what lets these three be written last.
            this.zip.writeEntry('xl/styles.xml', this.styles.xml());
            this.zip.writeEntry('xl/workbook.xml', workbookXml(this.sheetNames));
            this.zip.writeEntry(
                'xl/_rels/workbook.xml.rels',
                workbookRelsXml(this.sheetNames.length),
            );
            this.zip.end();
        });
    }
}
