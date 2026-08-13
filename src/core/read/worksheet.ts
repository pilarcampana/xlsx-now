// `xl/worksheets/sheetN.xml`: the data itself.
//
// This is the part that can be big — a sheet at Excel's limit is a quarter of
// a gigabyte of XML — and the only one of them the reader never holds whole.
// The chunks come out of the zip, go through the parser, and rows come out
// the other end as they close.
//
// Everything above it is small by comparison and read in one piece: the
// workbook, the relationships, the styles, and the shared strings. That last
// one is the honest limit of how little a reader can hold, and it is the
// format's doing — a cell says `<v>7</v>` and means the seventh entry of a
// table it does not carry.
import { columnIndex } from '../cell.js';
import type { CellType, StyledCell } from '../types.js';
import { readDate, readIsoDate, type ReadDates } from './dates.js';
import type { NumberFormats } from './numberFormats.js';
import type { ReadRow, ReadValue } from './types.js';
import { XmlParser } from './xml.js';

/** Days between the 1900 epoch and the 1904 one a Macintosh workbook uses. */
const DAYS_1904_TO_1900 = 1462;

/** A cell as the file spells it, before anything has been made of it. */
export interface RawCell {
    /** The `t`: what the cell says it holds. Absent means a number. */
    type: string | undefined;
    /** The `s`: which entry of `cellXfs` says how it is shown. */
    style: number | undefined;
    /** The text of `<v>`, or of the `<is>` a cell carries its own string in. */
    value: string | undefined;
    /** The expression of `<f>`, without the `=` a sheet does not store. */
    formula: string | undefined;
}

/** What a cell is read against: everything the worksheet itself does not say. */
export interface CellContext {
    sharedStrings: readonly string[];
    formats: NumberFormats;
    date1904: boolean;
    /** What a date is built as; see `ReadDates`. */
    dates: ReadDates;
}

/** `B12` as the coordinates it names, counting columns from 0 and rows from 1. */
export function parseCellReference(
    reference: string,
): { column: number; row: number } | undefined {
    const match = /^([A-Za-z]+)([0-9]+)$/.exec(reference);
    if (!match) return undefined;
    const column = columnIndex(match[1] as string);
    if (column === undefined) return undefined;
    return { column, row: Number(match[2]) };
}

/** The number in a `<v>`, or a failure that names what was there instead. */
function numberOf(raw: RawCell): number {
    const value = Number(raw.value);
    if (!Number.isFinite(value)) {
        throw new Error(`A numeric cell holds "${raw.value}", which is not a number.`);
    }
    return value;
}

/** The date a serial means, under whichever epoch the workbook counts from. */
function dateOf(serial: number, context: CellContext): ReadValue {
    return readDate(context.date1904 ? serial + DAYS_1904_TO_1900 : serial, context.dates);
}

/**
 * A cell as its value.
 *
 * The types are the file's own: `s` points into the shared strings, `b` is a
 * boolean written as `1` or `0`, `str` is the cached result of a formula,
 * `inlineStr` is text the cell carries itself, `e` is an error, `d` is a date
 * written out in full — rare, but in the spec — and no type at all is a
 * number. Which is where the format comes in: a number under a date format is
 * a date, and that is the only thing the styles are read for.
 *
 * An error comes back as its own text — `#DIV/0!` — rather than as `null`.
 * It is what the cell shows, and a `null` would be a value that went missing
 * without anyone saying so.
 */
export function cellValue(raw: RawCell, context: CellContext): ReadValue {
    if (raw.value === undefined) return null;
    switch (raw.type) {
        case undefined:
        case 'n': {
            const value = numberOf(raw);
            return context.formats.isDate(raw.style) ? dateOf(value, context) : value;
        }
        case 's': {
            const index = numberOf(raw);
            const text = context.sharedStrings[index];
            if (text === undefined) {
                throw new Error(
                    `A cell points at the shared string ${index}, and the table has ${context.sharedStrings.length}.`,
                );
            }
            return text;
        }
        case 'b':
            return raw.value === '1';
        case 'str':
        case 'inlineStr':
        case 'e':
            return raw.value;
        case 'd':
            // The cell that spells its date out instead of numbering it, which
            // is what a day the serial cannot number has to be written as. It
            // never becomes one: see `readIsoDate` — and the 1904 epoch has
            // nothing to shift in a text that names its own day.
            return readIsoDate(raw.value, context.dates);
        default:
            throw new Error(`A cell says it holds "${raw.type}", which is not a type a sheet has.`);
    }
}

/**
 * A cell as the writer would take it back.
 *
 * `t` is only written where the writer would not work it out for itself: a
 * string, a number, a boolean and a date all say what they are by being what
 * they are, and what is left is the cached result of a formula — text that
 * has to stay text — and an error, which is text that is not a value.
 *
 * The format goes in `s`, as a style written out, because that is where the
 * writer reads one: `{ v: 45306, s: { numFmt: 14 } }` is a cell that can go
 * straight back into a workbook and come out the same.
 */
export function styledCell(raw: RawCell, context: CellContext): StyledCell {
    const cell: StyledCell = { v: cellValue(raw, context) };
    const type: CellType | undefined =
        raw.type === 'e' ? 'e' : raw.type === 'str' ? 'str' : undefined;
    if (type !== undefined) cell.t = type;
    if (raw.formula !== undefined) cell.f = raw.formula;
    const numFmt = context.formats.forStyle(raw.style);
    if (numFmt !== undefined) cell.s = { numFmt };
    return cell;
}

/**
 * The rows of a worksheet, as the chunks of it go by.
 *
 * `saxes` calls back while a chunk is being written and a generator cannot
 * yield from inside a callback, so a chunk's rows are collected and handed
 * out right after it — the same order, with the rows in hand instead of on
 * the stack. Nothing accumulates past one chunk.
 */
export async function* readRows<C>(
    chunks: AsyncIterable<string>,
    convert: (raw: RawCell) => C,
    partName: string,
): AsyncIterable<ReadRow<C>> {
    const ready: ReadRow<C>[] = [];

    let inSheetData = false;
    let rowIndex = 0;
    let cells: (C | undefined)[] = [];
    let column = 0;
    let raw: RawCell | undefined;
    /** Where the text arriving now belongs, if anywhere. */
    let target: 'value' | 'formula' | undefined;

    const parser = new XmlParser(
        {
            open(name, attributes) {
                if (name === 'sheetData') {
                    inSheetData = true;
                    return;
                }
                if (!inSheetData) return;
                switch (name) {
                    case 'row': {
                        // A row is allowed to leave its number out, and then
                        // it is simply the next one.
                        const declared = Number(attributes['r']);
                        rowIndex = Number.isInteger(declared) ? declared : rowIndex + 1;
                        cells = [];
                        column = 0;
                        break;
                    }
                    case 'c': {
                        const reference = attributes['r'];
                        const at = reference === undefined ? undefined : parseCellReference(reference);
                        if (reference !== undefined && at === undefined) {
                            throw new Error(`"${reference}" is not a cell reference.`);
                        }
                        column = at?.column ?? column;
                        const style = Number(attributes['s']);
                        raw = {
                            type: attributes['t'],
                            style: Number.isInteger(style) ? style : undefined,
                            value: undefined,
                            formula: undefined,
                        };
                        break;
                    }
                    case 'v':
                    case 't':
                        if (raw) {
                            // `??=` and not `=`: the runs of an `<is>` are
                            // several `<t>` of one value, and they join.
                            raw.value ??= '';
                            target = 'value';
                        }
                        break;
                    case 'f':
                        if (raw) {
                            raw.formula ??= '';
                            target = 'formula';
                        }
                        break;
                }
            },
            text(text) {
                if (!raw || target === undefined) return;
                if (target === 'value') raw.value += text;
                else raw.formula += text;
            },
            close(name) {
                switch (name) {
                    case 'sheetData':
                        inSheetData = false;
                        break;
                    case 'v':
                    case 't':
                    case 'f':
                        target = undefined;
                        break;
                    case 'c':
                        if (raw) {
                            cells[column] = convert(raw);
                            raw = undefined;
                            column++;
                        }
                        break;
                    case 'row':
                        if (inSheetData) ready.push({ index: rowIndex, cells });
                        break;
                }
            },
        },
        partName,
    );

    for await (const chunk of chunks) {
        parser.write(chunk);
        yield* ready.splice(0);
    }
    parser.close();
    yield* ready.splice(0);
}
