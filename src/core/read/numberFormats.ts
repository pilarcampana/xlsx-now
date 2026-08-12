// `xl/styles.xml`, read for the one thing a reader of data cannot do without.
//
// A date in a sheet is a number. Nothing in the cell says otherwise: `45306`
// is `45306` whether it means forty-five thousand of something or the 15th of
// January of 2024, and the only thing that tells them apart is the number
// format its style points at. So the reader parses the styles it was told not
// to care about — but only this far: which format each style shows, and
// whether that format is a date. Fonts, fills and borders are not read.
import type { DateKind } from '../dates.js';
import { parseXml } from './xml.js';

const PART_NAME = 'xl/styles.xml';

/** The general format: a number shown as whatever it is. Never a date. */
const GENERAL = 0;

/**
 * The formats every reader is born knowing, of the ones that are dates, and
 * which of the three things each of them shows.
 *
 * 14 to 17 are the short dates, 18 to 21 the times, 22 the one that is both,
 * and 45 to 47 the elapsed times. The two runs in between — 27 to 36 and 50
 * to 58 — are dates and times as written in Japanese, Chinese and Korean
 * locales; a file made in one of those holds ordinary dates under those ids,
 * so leaving them out would come back as bare serial numbers rather than as
 * an error.
 *
 * The ids and their codes are ECMA-376's own table (18.8.30), and the kinds
 * are read off those codes — `d-mmm-yy` is a date, `h:mm:ss` is a time,
 * `m/d/yy h:mm` is both. They are written out rather than worked out because
 * a built-in has no code in the file to work anything out from: what the
 * reader has is the number.
 */
const BUILTIN_DATE_KINDS = new Map<number, DateKind>([
    // mm-dd-yy, d-mmm-yy, d-mmm, mmm-yy
    [14, 'date'], [15, 'date'], [16, 'date'], [17, 'date'],
    // h:mm AM/PM, h:mm:ss AM/PM, h:mm, h:mm:ss
    [18, 'time'], [19, 'time'], [20, 'time'], [21, 'time'],
    // m/d/yy h:mm
    [22, 'dateTime'],
    // The Japanese, Chinese and Korean runs: dates …
    [27, 'date'], [28, 'date'], [29, 'date'], [30, 'date'], [31, 'date'], [36, 'date'],
    [50, 'date'], [51, 'date'], [54, 'date'], [57, 'date'], [58, 'date'],
    // … and the times among them.
    [32, 'time'], [33, 'time'], [34, 'time'], [35, 'time'],
    [52, 'time'], [53, 'time'], [55, 'time'], [56, 'time'],
    // mm:ss, [h]:mm:ss, mmss.0 — elapsed, which is a time with no day in it.
    [45, 'time'], [46, 'time'], [47, 'time'],
]);

/** The letters a format code shows a date or a time with. */
const DATE_LETTERS = new Set(['y', 'd', 'h', 's', 'm']);

/** One run of the same date letter, as the scan below finds them. */
interface LetterRun {
    letter: string;
    /** How many of it in a row: `mm` is two, and `mmm` is three. */
    length: number;
}

/**
 * The runs of date letters in a format code, with everything that is not one
 * left out.
 *
 * A code is a template with literals in it, and the literals are the whole
 * difference between a format and its text: `#,##0 "días"` shows a number and
 * says `d` twice, and `[Red]0.00` says nothing at all through its `Red`. So
 * the scan skips what a reader of the code would skip — quoted text, escaped
 * characters, the padding of `_` and the fill of `*`, and whatever is in
 * brackets.
 *
 * Brackets are the one exception to being skipped: `[h]`, `[mm]` and `[ss]`
 * are elapsed time, which is the only thing in brackets that is a value
 * rather than an instruction. The rest — a colour, a condition, the `[$-409]`
 * of a locale — is not.
 *
 * Runs and not letters, because `mmm` and `mm` are different things and the
 * length is what says so.
 */
function letterRuns(code: string): LetterRun[] {
    const runs: LetterRun[] = [];
    for (let index = 0; index < code.length; index++) {
        const char = code[index] as string;
        if (char === '"') {
            const end = code.indexOf('"', index + 1);
            index = end === -1 ? code.length : end;
        } else if (char === '\\' || char === '_' || char === '*') {
            index++;
        } else if (char === '[') {
            const end = code.indexOf(']', index + 1);
            const inside = code.slice(index + 1, end === -1 ? code.length : end);
            if (/^[hms]+$/i.test(inside)) {
                runs.push({ letter: inside[0]?.toLowerCase() as string, length: inside.length });
            }
            index = end === -1 ? code.length : end;
        } else {
            const letter = char.toLowerCase();
            if (!DATE_LETTERS.has(letter)) continue;
            let length = 1;
            while ((code[index + length] as string | undefined)?.toLowerCase() === letter) length++;
            runs.push({ letter, length });
            index += length - 1;
        }
    }
    return runs;
}

/**
 * Whether an `m` is minutes rather than months — the one letter in a format
 * code that means two things.
 *
 * Excel's rule, and the one every implementation copies: an `m` right after
 * an hour or right before a second is minutes, and anywhere else it is a
 * month. `h:mm` and `mm:ss` are minutes on either side of that; `mm/dd/yy` is
 * a month, and so is the `m` of `yyyy"年"m"月"d"日"`, which has no hour and no
 * second anywhere near it.
 *
 * Three or more is never minutes: `mmm` is `Jan` and `mmmm` is `January`,
 * and a clock has no use for either.
 */
function isMinutes(runs: readonly LetterRun[], at: number): boolean {
    if ((runs[at] as LetterRun).length >= 3) return false;
    return runs[at - 1]?.letter === 'h' || runs[at + 1]?.letter === 's';
}

/**
 * What a format code shows: a date, a time, both, or neither — and neither is
 * what makes a number stay a number.
 */
export function dateFormatKind(code: string): DateKind | undefined {
    const runs = letterRuns(code);
    let date = false;
    let time = false;
    for (let index = 0; index < runs.length; index++) {
        switch ((runs[index] as LetterRun).letter) {
            case 'y':
            case 'd':
                date = true;
                break;
            case 'h':
            case 's':
                time = true;
                break;
            case 'm':
                if (isMinutes(runs, index)) time = true;
                else date = true;
                break;
        }
    }
    if (date && time) return 'dateTime';
    if (date) return 'date';
    if (time) return 'time';
    return undefined;
}

/** Whether a format code writes a date or a time at all. */
export function isDateFormat(code: string): boolean {
    return dateFormatKind(code) !== undefined;
}

/**
 * What the styles of a workbook say about the values under them.
 *
 * Indexed by the `s` of a cell, which is a position in `cellXfs` — the one
 * list of the part that matters here. `cellStyleXfs` is a list of the same
 * shape right next to it, holding the named styles that cell formats inherit
 * from, and reading the two as one would shift every index by however many
 * are in the first.
 */
export interface NumberFormats {
    /**
     * The format a cell with this style shows, as the writer takes one: the
     * code itself when the workbook declared it, the id when it is built in,
     * and nothing at all for `General`. What comes back can be handed
     * straight back to the writer.
     */
    forStyle(style: number | undefined): string | number | undefined;
    /**
     * What a number under this style is: a date, a time, both, or — which is
     * the usual answer — `undefined`, meaning it is just a number.
     */
    dateKind(style: number | undefined): DateKind | undefined;
}

/** A workbook with no styles part, where nothing is a date. */
export const NO_FORMATS: NumberFormats = {
    forStyle: () => undefined,
    dateKind: () => undefined,
};

export function readNumberFormats(xml: string): NumberFormats {
    /** The codes the workbook declared, by the id it gave them. */
    const codes = new Map<number, string>();
    /** The format id of each `cellXfs` entry, in order. */
    const styles: number[] = [];
    let inCellXfs = false;

    parseXml(
        xml,
        {
            open(name, attributes) {
                if (name === 'numFmt') {
                    const id = Number(attributes['numFmtId']);
                    const code = attributes['formatCode'];
                    if (Number.isInteger(id) && code !== undefined) codes.set(id, code);
                } else if (name === 'cellXfs') inCellXfs = true;
                else if (name === 'xf' && inCellXfs) {
                    const id = Number(attributes['numFmtId'] ?? GENERAL);
                    styles.push(Number.isInteger(id) ? id : GENERAL);
                }
            },
            close(name) {
                if (name === 'cellXfs') inCellXfs = false;
            },
        },
        PART_NAME,
    );

    /** The format id a style points at; `General` for a style there is none of. */
    function formatId(style: number | undefined): number {
        if (style === undefined) return GENERAL;
        return styles[style] ?? GENERAL;
    }

    // Worked out once per format rather than per cell: a sheet of a million
    // dates asks this a million times and there are only ever a handful of
    // answers. `null` is the answer "not a date", which has to be told apart
    // from "not asked yet" — `undefined` is already taken by the first.
    const kindById = new Map<number, DateKind | null>();
    function kindOfId(id: number): DateKind | undefined {
        const known = kindById.get(id);
        if (known !== undefined) return known ?? undefined;
        const code = codes.get(id);
        const answer = code === undefined ? BUILTIN_DATE_KINDS.get(id) : dateFormatKind(code);
        kindById.set(id, answer ?? null);
        return answer;
    }

    return {
        forStyle(style) {
            const id = formatId(style);
            if (id === GENERAL) return undefined;
            return codes.get(id) ?? id;
        },
        dateKind: (style) => kindOfId(formatId(style)),
    };
}
