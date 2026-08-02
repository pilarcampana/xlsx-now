import { WidthMeter } from './autoWidth.js';
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
    sheetHeaderXml,
    SHEET_FOOTER,
    type ColumnFormats,
    type Freeze,
    type RowOptions,
} from './sheet.js';
import { StyleTable, type StyleSpec } from './styles.js';
import type { CellRow, Row } from './types.js';
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

export interface XlsxWriterOptions extends SheetOptions {
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
    /** The sheets so far, in order; the last of them may still be open. */
    private readonly sheetNames: string[] = [];
    /** The columns of the sheet being written, if it has any. */
    private mode: ColumnsMode | undefined;
    /** What the sheet's cells are measuring, when it sizes its columns by them. */
    private widths: WidthMeter | undefined;
    /**
     * The header of a sheet whose columns are being measured: it cannot be
     * written until the sheet closes, so what it will be made of waits here —
     * the meter included — and the rows pile up behind it.
     */
    private pendingHeader:
        | { freeze: Freeze; columnFormats: ColumnFormats | undefined; widths: WidthMeter }
        | undefined;
    private open = false;
    private batch = '';
    private rowNumber = 1;

    constructor(sink: (bytes: Uint8Array) => void, options: XlsxWriterOptions = {}) {
        this.defaults = options;
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

    private pushBatch(): void {
        if (!this.batch) return;
        this.zip.push(encoder.encode(this.batch));
        this.batch = '';
    }

    private writeCellRow(row: CellRow, options?: RowOptions): void {
        this.batch += cellRowXml(this.rowNumber, row, this.styles, options, this.widths);
        this.rowNumber++;
        // A sheet being measured has nowhere to push to: its `<cols>` is
        // written from rows that have not arrived yet, so the whole worksheet
        // waits until it closes.
        if (!this.widths && this.batch.length >= PUSH_BATCH_CHARS) this.pushBatch();
    }

    /**
     * Starts a worksheet part, and its header row when it has columns.
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
        const mode = columns?.length ? columnsMode(columns) : undefined;
        const freeze = {
            rows: sheet.freezeRows ?? this.defaults.freezeRows ?? mode?.freeze.rows ?? 0,
            columns:
                sheet.freezeColumns ?? this.defaults.freezeColumns ?? mode?.freeze.columns ?? 0,
        };

        this.sheetNames.push(name);
        this.widths = autoWidthMax === undefined ? undefined : new WidthMeter(autoWidthMax);
        if (this.widths) {
            // Nothing of this sheet can go out yet: `<cols>` is written before
            // the first row and the widths come from the last one.
            this.pendingHeader = { freeze, columnFormats, widths: this.widths };
            this.batch = '';
        } else {
            // The worksheet stays open until the next command or `finish`: it
            // is the one part whose length nobody knows yet.
            this.zip.startEntry(worksheetPart(this.sheetNames.length));
            this.pendingHeader = undefined;
            this.batch = sheetHeaderXml(freeze, this.styles, columnFormats);
        }
        this.rowNumber = 1;
        this.mode = mode;
        this.open = true;
        // Enough columns and the header row alone fills a batch, so writing it
        // can reach the zip and fail there like any other row does.
        if (mode) this.writeCellRow(mode.headerRow);
    }

    private closeSheet(): void {
        this.batch += SHEET_FOOTER;
        if (this.pendingHeader) {
            // Now the columns have been measured, so the header they were
            // waiting for can be written and the sheet goes into the archive
            // whole, in the place it would have taken anyway.
            const { freeze, columnFormats, widths } = this.pendingHeader;
            this.batch =
                sheetHeaderXml(freeze, this.styles, columnFormats, widths.columnWidths()) +
                this.batch;
            this.pendingHeader = undefined;
            this.zip.startEntry(worksheetPart(this.sheetNames.length));
        }
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
