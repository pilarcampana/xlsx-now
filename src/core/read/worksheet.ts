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
import { kindOf, serialOfParts, type DateKind, type DateParts } from '../dates.js';
import type { CellType, StyledCell } from '../types.js';
import type { DateReader } from './dates.js';
import type { NumberFormats } from './numberFormats.js';
import type { RawReadValue, ReadRow } from './types.js';
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
    /** What a serial under a date format becomes: the `dates` option, resolved. */
    dates: DateReader<unknown>;
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

/**
 * The `d` type of the spec: a date written out instead of numbered, in the
 * ISO 8601 spelling and with no zone on it — which is the same wall clock
 * every other date in a sheet is.
 *
 * Parsed here rather than by `new Date`, whose reading of these depends on
 * what is in them: a bare `2024-01-15` is taken as UTC and the same string
 * with a time on it is taken as local, so the two would land a few hours
 * apart on the same machine.
 *
 * A zone on the end — a `Z`, an `+03:00` — is allowed and then ignored, and
 * the wall clock is read as it stands. Shifting it by the offset is the one
 * thing that cannot be done here: there is nowhere to shift it *to*. Every
 * other date in the sheet is a serial, which is a wall clock and carries no
 * zone, so honouring this one would leave a single cell meaning something
 * different from the rest of its column — and, on a machine in another zone,
 * something different from what it meant yesterday.
 */
const ISO_DATE =
    /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3})\d*)?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/i;

/** Those parts, or `undefined` when the text is not one of them. */
function isoParts(text: string): DateParts | undefined {
    const match = ISO_DATE.exec(text.trim());
    if (!match) return undefined;
    const at = (index: number, scale = ''): number =>
        match[index] === undefined ? 0 : Number(match[index].padEnd(scale.length, '0'));
    return {
        year: Number(match[1]),
        month: Number(match[2]),
        day: Number(match[3]),
        hour: at(4),
        minute: at(5),
        second: at(6),
        millisecond: at(7, '000'),
    };
}

/** The number in a `<v>`, or a failure that names what was there instead. */
function numberOf(raw: RawCell): number {
    const value = Number(raw.value);
    if (!Number.isFinite(value)) {
        throw new Error(`A numeric cell holds "${raw.value}", which is not a number.`);
    }
    return value;
}

/**
 * The date a serial means, as whatever the caller asked dates to be. The
 * epoch is corrected for here, so a reader is never handed a serial that
 * counts from anything but 1899-12-30.
 */
function dateOf(serial: number, kind: DateKind, context: CellContext): RawReadValue {
    const from1900 = context.date1904 ? serial + DAYS_1904_TO_1900 : serial;
    // A reader gives back whatever it was written to give back, and this is
    // where that stops being tracked: `openXlsx` states the link between the
    // option and the type in its signature, and one value cannot carry it.
    return context.dates.read(from1900, { kind }) as RawReadValue;
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
export function cellValue(raw: RawCell, context: CellContext): RawReadValue {
    if (raw.value === undefined) return null;
    switch (raw.type) {
        case undefined:
        case 'n': {
            const value = numberOf(raw);
            const kind = context.formats.dateKind(raw.style);
            return kind === undefined ? value : dateOf(value, kind, context);
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
        case 'd': {
            const parts = isoParts(raw.value);
            if (parts === undefined) {
                throw new Error(`A date cell holds "${raw.value}", which is not a date.`);
            }
            // Back to a serial and out through the same reader as every other
            // date, rather than answered here with a `Date`: a `d` cell is a
            // date written out in full instead of numbered, and which of the
            // two spellings a file happened to use is not a reason for the
            // sheet to come back holding two different types.
            return dateOf(serialOfParts(parts), kindOf(parts), context);
        }
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
