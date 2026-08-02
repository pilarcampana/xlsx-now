// The style table, built while the rows go by.
//
// A cell asks for a style with an `s`, and `s` is a 0-based index into
// styles.xml's <cellXfs> — so the table has to exist by the time the file is
// read, not by the time the row is written. `xl/styles.xml` is written in
// `finish()`, next to `xl/workbook.xml` and for the same reason: the order of
// the entries inside the zip is nobody's business but the central directory's,
// so a part that cannot be known upfront is simply written last.
//
// That is what lets a cell carry a whole style instead of a number. Every
// distinct combination is registered once and reused from then on, so what
// this holds is bounded by how many different styles the workbook has, not by
// how many rows it has.
import { hasTimeOfDay, sanitizeText } from './cell.js';

/**
 * A colour, as hex: `#RGB`, `#RRGGBB`, `RRGGBB` or `AARRGGBB`, with the `#`
 * optional. Everything is normalized to the ARGB that xlsx stores; without an
 * alpha it is taken as opaque.
 */
export type Color = string;

/** The line a border side is drawn with. Excel's own list, unabridged. */
export type BorderStyle =
    | 'thin'
    | 'medium'
    | 'thick'
    | 'double'
    | 'hair'
    | 'dotted'
    | 'dashed'
    | 'dashDot'
    | 'dashDotDot'
    | 'mediumDashed'
    | 'mediumDashDot'
    | 'mediumDashDotDot'
    | 'slantDashDot';

/** One side of a border: the line alone, or the line and its colour. */
export type BorderSide = BorderStyle | { style?: BorderStyle; color?: Color };

export interface BorderSpec {
    /** All four sides at once, under whatever a side says for itself. */
    all?: BorderSide;
    left?: BorderSide;
    right?: BorderSide;
    top?: BorderSide;
    bottom?: BorderSide;
    /** The diagonal line; `diagonalUp`/`diagonalDown` say which way it runs. */
    diagonal?: BorderSide;
    diagonalUp?: boolean;
    diagonalDown?: boolean;
}

/**
 * Everything a cell can look like, flat. What xlsx keeps in four separate
 * tables — the number format, the font, the fill, the border — is one object
 * here, and taking it apart is this module's job.
 */
export interface StyleSpec {
    /**
     * A style declared in the writer options to start from; what this one says
     * goes over it. It is how a cell reuses a named style and changes one
     * thing about it: `{ base: 'money', bold: true }`.
     */
    base?: string;

    /** Font name. Defaults to Calibri, which is what a sheet uses without one. */
    font?: string;
    /** Size in points. Defaults to 11. */
    size?: number;
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    /** `true` is a single underline; the rest are Excel's other three. */
    underline?: boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
    /** Superscript or subscript. */
    script?: 'super' | 'sub';
    /** Colour of the text. */
    color?: Color;

    /** Colour of the cell's background — a solid fill. */
    bg?: Color;

    align?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
    valign?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
    /** Wraps the text instead of letting it run over the next cell. */
    wrap?: boolean;
    /**
     * Degrees counterclockwise, -90 to 90. `255` is Excel's own spelling of
     * the vertical layout, where the letters stack instead of turning.
     */
    rotate?: number;
    /** Indent steps from the side the text is aligned to. */
    indent?: number;
    /** Shrinks the text until it fits, instead of wrapping or overflowing. */
    shrink?: boolean;

    /**
     * The number format, as Excel spells it: `'#,##0.00'`, `'yyyy-mm-dd'`,
     * `'0.00%'`. A number is one of Excel's built-in formats, by id.
     */
    numFmt?: string | number;

    border?: BorderSpec;

    /**
     * Whether the cell is locked once the sheet is protected. Excel's own
     * default is locked, which is what applies without this.
     */
    locked?: boolean;
    /** Hides the formula from the formula bar on a protected sheet. */
    hideFormula?: boolean;
}

/** What a cell, a row or a column asks for: a declared style by name, or one outright. */
export type StyleRef = string | StyleSpec;

/**
 * What a `Date` is shown as when its cell asks for no format of its own:
 * Excel's built-in short date, and the built-in date and time next to it.
 *
 * A sheet has no rendered dates in it — there is no preview to write. What a
 * date cell holds is the serial number and the id of a format, and the
 * spelling of the day is worked out by whoever opens the file. Ids 14 and 22
 * are the two the spec hands to the reader's own locale: ECMA-376 lists them
 * as `mm-dd-yy` and `m/d/yy h:mm`, and Microsoft's implementation notes for
 * that same clause say Excel shows them in the short date of the system it is
 * running on. So a date written under 14 reads `15/01/2024` in Buenos Aires
 * and `1/15/2024` in Chicago — it is what "Short Date" in Excel's own format
 * menu produces, and the reason it is the default here.
 *
 * `dateFormat` is how a workbook says otherwise, with a format code of its
 * own: `'yyyy-mm-dd'` is the ISO order, which reads the same everywhere and
 * is nobody's local custom.
 */
export const DEFAULT_DATE_FORMAT = 14;
export const DEFAULT_DATETIME_FORMAT = 22;

/**
 * The time of day, added to a date format that carries none — the one part of
 * a timestamp every locale writes the same way, which is why `dateFormat` is
 * a date and this is not asked for.
 */
const TIME_OF_DAY = 'hh:mm:ss';

/**
 * How wide a built-in date is taken to be by `autoWidthMax`, in characters.
 * The locale decides the real one, so what is measured is the widest a short
 * date runs to — `dd/mm/yyyy`, and the same with the time after it.
 */
const BUILTIN_DATE_WIDTH = 10;
const BUILTIN_DATETIME_WIDTH = BUILTIN_DATE_WIDTH + 1 + TIME_OF_DAY.length;

/** Excel's built-in formats take the ids below this one; ours start here. */
const FIRST_CUSTOM_NUMFMT = 164;

/** What a workbook says about the format its dates fall back to. */
export interface DateFormatOptions {
    /**
     * The format a `Date` with no time of day is shown in: a format code
     * (`'yyyy-mm-dd'`, `'dd/mm/yy'`) or the id of one of Excel's built-in
     * formats. Defaults to `14`, the built-in short date, which every reader
     * shows the way the machine it is running on writes a date.
     */
    dateFormat?: string | number;
    /**
     * The same, for a `Date` that carries a time of day. Defaults to
     * `dateFormat` with `hh:mm:ss` after it — and to the built-in `22` when
     * `dateFormat` is the built-in short date, since an id has no format code
     * to add anything to.
     */
    dateTimeFormat?: string | number;
}

/**
 * The two formats a `Date` falls back to, worked out once for the workbook.
 *
 * There is one of these behind every date in the file — the format a cell
 * gets when its style says nothing, and the width that date is measured as
 * when the sheet is sizing its columns.
 */
export class DateFormats {
    readonly date: string | number;
    readonly dateTime: string | number;

    constructor({ dateFormat, dateTimeFormat }: DateFormatOptions = {}) {
        this.date = dateFormat ?? DEFAULT_DATE_FORMAT;
        this.dateTime = dateTimeFormat ?? this.impliedDateTime();
    }

    /** What a timestamp is shown as when only the date format was given. */
    private impliedDateTime(): string | number {
        if (typeof this.date === 'string') return `${this.date} ${TIME_OF_DAY}`;
        // A built-in is an id, and an id has no format code to add a time to.
        // Excel's own pair for the short date is 14 and 22; anything else has
        // to be said outright rather than guessed at.
        if (this.date === DEFAULT_DATE_FORMAT) return DEFAULT_DATETIME_FORMAT;
        throw new Error(
            `A dateFormat of ${this.date} is one of Excel's built-in formats, and a built-in ` +
                'has no format code to add a time of day to: say dateTimeFormat as well, or ' +
                'write dateFormat out as a format code.',
        );
    }

    /** The format this value falls back to: the date, or the date and time. */
    for(value: Date): string | number {
        return hasTimeOfDay(value) ? this.dateTime : this.date;
    }

    /**
     * How many characters this value shows. A format code is measured as it
     * is written, which is what the date under it comes to; a built-in is the
     * reader's own, so it is measured as the widest one.
     */
    textLength(value: Date): number {
        const format = this.for(value);
        if (typeof format === 'string') return format.length;
        return hasTimeOfDay(value) ? BUILTIN_DATETIME_WIDTH : BUILTIN_DATE_WIDTH;
    }
}

/** The formats of a workbook that said nothing about them. */
export const DEFAULT_DATE_FORMATS = new DateFormats();

/** What a sheet looks like with no style at all — `<cellXfs>` entry 0. */
const DEFAULT_FONT =
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
/**
 * Fill 0 and fill 1 are fixed by Excel: it takes `none` and `gray125` to be
 * the first two entries whatever a file says, so a fill written over them is
 * a fill nobody sees.
 */
const RESERVED_FILLS = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
];
const EMPTY_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';
const DEFAULT_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

const HORIZONTAL = {
    left: 'left',
    center: 'center',
    right: 'right',
    fill: 'fill',
    justify: 'justify',
    centerContinuous: 'centerContinuous',
    distributed: 'distributed',
} as const;
/** `middle` is what everyone calls it; xlsx calls it `center`. */
const VERTICAL = {
    top: 'top',
    middle: 'center',
    bottom: 'bottom',
    justify: 'justify',
    distributed: 'distributed',
} as const;

/**
 * A colour as the eight hex digits xlsx stores. Anything that is not a colour
 * is refused here rather than written out as an attribute Excel will choke on
 * once the rows are already gone.
 */
export function argb(color: Color): string {
    const hex = color.replace(/^#/, '').toUpperCase();
    if (!/^[0-9A-F]+$/.test(hex) || ![3, 6, 8].includes(hex.length)) {
        throw new Error(
            `"${color}" is not a colour: write it as #RGB, #RRGGBB, RRGGBB or AARRGGBB.`,
        );
    }
    if (hex.length === 3) return `FF${hex.replace(/./g, (digit) => digit + digit)}`;
    return hex.length === 6 ? `FF${hex}` : hex;
}

function colorXml(tag: string, color: Color | undefined): string {
    return color === undefined ? '' : `<${tag} rgb="${argb(color)}"/>`;
}

/**
 * Degrees as Excel counts them: 0-90 counterclockwise, and clockwise from 91
 * on, where 91 is one degree down. `255`, the vertical layout, is its own
 * value and passes through.
 */
function textRotation(rotate: number): number {
    if (rotate === 255) return 255;
    if (!Number.isInteger(rotate) || rotate < -90 || rotate > 90) {
        throw new Error(
            `Rotation ${rotate} is not between -90 and 90 (or 255, for vertical text).`,
        );
    }
    return rotate < 0 ? 90 - rotate : rotate;
}

function fontXml(spec: StyleSpec): string {
    const underline =
        spec.underline === undefined || spec.underline === false
            ? ''
            : spec.underline === true || spec.underline === 'single'
              ? '<u/>'
              : `<u val="${spec.underline}"/>`;
    return (
        '<font>' +
        (spec.bold ? '<b/>' : '') +
        (spec.italic ? '<i/>' : '') +
        (spec.strike ? '<strike/>' : '') +
        underline +
        (spec.script ? `<vertAlign val="${spec.script === 'super' ? 'superscript' : 'subscript'}"/>` : '') +
        `<sz val="${spec.size ?? 11}"/>` +
        (spec.color === undefined ? '<color theme="1"/>' : colorXml('color', spec.color)) +
        `<name val="${sanitizeText(spec.font ?? 'Calibri')}"/><family val="2"/>` +
        // `minor` is the theme's body font, which is only Calibri while
        // nobody has asked for another one.
        (spec.font === undefined ? '<scheme val="minor"/>' : '') +
        '</font>'
    );
}

function fillXml(bg: Color): string {
    // `bgColor` is the pattern's *other* colour, and a solid pattern has none;
    // `indexed="64"` is the "whatever the window is" Excel writes there.
    return `<fill><patternFill patternType="solid">${colorXml('fgColor', bg)}<bgColor indexed="64"/></patternFill></fill>`;
}

function borderSideXml(tag: string, side: BorderSide | undefined): string {
    if (side === undefined) return `<${tag}/>`;
    const { style, color } = typeof side === 'string' ? { style: side, color: undefined } : side;
    if (style === undefined) return `<${tag}/>`;
    const rgb = colorXml('color', color);
    return rgb ? `<${tag} style="${style}">${rgb}</${tag}>` : `<${tag} style="${style}"/>`;
}

function borderXml(border: BorderSpec): string {
    const side = (own: BorderSide | undefined): BorderSide | undefined => own ?? border.all;
    return (
        '<border' +
        (border.diagonalUp ? ' diagonalUp="1"' : '') +
        (border.diagonalDown ? ' diagonalDown="1"' : '') +
        '>' +
        borderSideXml('left', side(border.left)) +
        borderSideXml('right', side(border.right)) +
        borderSideXml('top', side(border.top)) +
        borderSideXml('bottom', side(border.bottom)) +
        // The diagonal is not one of the four sides `all` draws: a border
        // around a cell is not a border across it.
        borderSideXml('diagonal', border.diagonal) +
        '</border>'
    );
}

function alignmentXml(spec: StyleSpec): string {
    let attributes = '';
    if (spec.align !== undefined) attributes += ` horizontal="${HORIZONTAL[spec.align]}"`;
    if (spec.valign !== undefined) attributes += ` vertical="${VERTICAL[spec.valign]}"`;
    if (spec.rotate !== undefined) attributes += ` textRotation="${textRotation(spec.rotate)}"`;
    if (spec.wrap) attributes += ' wrapText="1"';
    if (spec.indent !== undefined) attributes += ` indent="${spec.indent}"`;
    if (spec.shrink) attributes += ' shrinkToFit="1"';
    return attributes ? `<alignment${attributes}/>` : '';
}

function protectionXml(spec: StyleSpec): string {
    let attributes = '';
    if (spec.locked !== undefined) attributes += ` locked="${spec.locked ? 1 : 0}"`;
    if (spec.hideFormula) attributes += ' hidden="1"';
    return attributes ? `<protection${attributes}/>` : '';
}

/**
 * One of styles.xml's tables: the entries in the order they were handed out,
 * and the same rendered entry always under the same index. Deduplicating on
 * the XML is what makes two styles that were spelled differently — a named one
 * and the same one written out, `#ff0` and `FFFFFF00` — the one entry they are.
 */
class Table {
    readonly entries: string[] = [];
    private readonly indexes = new Map<string, number>();

    constructor(initial: readonly string[]) {
        for (const entry of initial) this.indexOf(entry);
    }

    indexOf(entry: string): number {
        const known = this.indexes.get(entry);
        if (known !== undefined) return known;
        const index = this.entries.length;
        this.entries.push(entry);
        this.indexes.set(entry, index);
        return index;
    }
}

/** One `<name count="n">` section, left out entirely when it holds nothing. */
function section(name: string, entries: readonly string[]): string {
    if (!entries.length) return '';
    return `<${name} count="${entries.length}">${entries.join('')}</${name}>`;
}

/**
 * The `<cellXfs>` of a workbook, filled in as its cells ask for things.
 *
 * `0` is the default style — which is why an unstyled cell can leave the `s`
 * attribute out altogether — and every combination past it is registered the
 * first time something asks for it.
 */
export class StyleTable {
    /** The styles the writer options declared, by name. */
    private readonly declared: Readonly<Record<string, StyleSpec>>;
    /** What a `Date` with no format of its own is shown as. */
    private readonly dates: DateFormats;

    private readonly numFmts = new Table([]);
    private readonly fonts = new Table([DEFAULT_FONT]);
    private readonly fills = new Table(RESERVED_FILLS);
    private readonly borders = new Table([EMPTY_BORDER]);
    private readonly xfs = new Table([DEFAULT_XF]);

    /** A declared style, merged with whatever it is based on. */
    private readonly merged = new Map<string, StyleSpec>();
    /** The two refs that repeat: a name, and a spec object by identity. */
    private readonly byName = new Map<string, number>();
    private readonly bySpec = new WeakMap<StyleSpec, number>();
    /** `<index>|<format>` -> the same style with a date format added to it. */
    private readonly dated = new Map<string, number>();

    constructor(
        declared: Readonly<Record<string, StyleSpec>> = {},
        dates: DateFormats = DEFAULT_DATE_FORMATS,
    ) {
        this.declared = declared;
        this.dates = dates;
    }

    private numFmtId(numFmt: string | number | undefined): number {
        if (numFmt === undefined) return 0;
        // A number is one of Excel's own formats, which are not written out.
        if (typeof numFmt === 'number') return numFmt;
        return (
            FIRST_CUSTOM_NUMFMT +
            this.numFmts.indexOf(`<numFmt formatCode="${sanitizeText(numFmt)}"/>`)
        );
    }

    /** One flat spec as one `<cellXfs>` entry, and every table under it. */
    private register(spec: StyleSpec): number {
        const numFmtId = this.numFmtId(spec.numFmt);
        const fontId = this.fonts.indexOf(fontXml(spec));
        const fillId = spec.bg === undefined ? 0 : this.fills.indexOf(fillXml(spec.bg));
        const borderId =
            spec.border === undefined ? 0 : this.borders.indexOf(borderXml(spec.border));
        const alignment = alignmentXml(spec);
        const protection = protectionXml(spec);

        // `applyX` is what tells Excel the entry means the value next to it,
        // rather than the one it would inherit from `xfId`.
        const xf =
            `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
            (numFmtId ? ' applyNumberFormat="1"' : '') +
            (fontId ? ' applyFont="1"' : '') +
            (fillId ? ' applyFill="1"' : '') +
            (borderId ? ' applyBorder="1"' : '') +
            (alignment ? ' applyAlignment="1"' : '') +
            (protection ? ' applyProtection="1"' : '');
        return this.xfs.indexOf(
            alignment || protection ? `${xf}>${alignment}${protection}</xf>` : `${xf}/>`,
        );
    }

    /**
     * A spec with its `base` folded in, so what `register` sees is one flat
     * object. A `base` that leads back to where it started is refused rather
     * than followed forever.
     */
    private resolve(spec: StyleSpec, seen: readonly string[]): StyleSpec {
        if (spec.base === undefined) return spec;
        if (seen.includes(spec.base)) {
            throw new Error(
                `Style "${spec.base}" is based on itself: ${[...seen, spec.base].join(' -> ')}.`,
            );
        }
        const { base, ...own } = spec;
        return { ...this.named(base, seen), ...own };
    }

    /**
     * The declared style called `name`, merged and remembered. A name nobody
     * declared is refused: taken as the default style instead, it would come
     * out as a workbook that is silently missing what it asked for.
     */
    private named(name: string, seen: readonly string[] = []): StyleSpec {
        const known = this.merged.get(name);
        if (known !== undefined) return known;
        const declared = this.declared[name];
        if (declared === undefined) {
            const names = Object.keys(this.declared);
            throw new Error(
                `Unknown style "${name}": declare it in the writer's "styles". ` +
                    (names.length
                        ? `The declared ones are ${names.join(', ')}.`
                        : 'None are declared.'),
            );
        }
        const spec = this.resolve(declared, [...seen, name]);
        this.merged.set(name, spec);
        return spec;
    }

    /** What a ref says, flat: a declared style by name, or the spec itself. */
    spec(ref: StyleRef): StyleSpec {
        return typeof ref === 'string' ? this.named(ref) : this.resolve(ref, []);
    }

    /** The `<cellXfs>` index for a ref. `undefined` is the default style, 0. */
    index(ref: StyleRef | undefined): number {
        if (ref === undefined) return 0;
        if (typeof ref === 'string') {
            const known = this.byName.get(ref);
            if (known !== undefined) return known;
            const index = this.register(this.named(ref));
            this.byName.set(ref, index);
            return index;
        }
        // A spec object the caller holds on to — one declared next to the
        // columns, say — is recognized before anything is rendered at all.
        const known = this.bySpec.get(ref);
        if (known !== undefined) return known;
        const index = this.register(this.resolve(ref, []));
        this.bySpec.set(ref, index);
        return index;
    }

    /**
     * The index a cell holding `value` gets. It is `index(ref)`, except for a
     * `Date`: a date is a number in a sheet, and a number with no format is
     * shown as the five-digit serial it is. So a date whose style says nothing
     * about the format gets one — the date alone, or the date and the time,
     * depending on what the value carries.
     */
    forValue(value: unknown, ref: StyleRef | undefined): number {
        if (!(value instanceof Date)) return this.index(ref);
        const format = this.dates.for(value);
        const base = this.index(ref);
        const key = `${base}|${format}`;
        const known = this.dated.get(key);
        if (known !== undefined) return known;
        const spec = ref === undefined ? {} : this.spec(ref);
        // A style that already says how to show the number is left alone:
        // asking for a format is how a caller says it wants that one.
        const index = spec.numFmt !== undefined ? base : this.register({ ...spec, numFmt: format });
        this.dated.set(key, index);
        return index;
    }

    /** The part, as it stands. Written once, at the end, and never before. */
    xml(): string {
        // The ids are `FIRST_CUSTOM_NUMFMT` plus the position, so each entry
        // has to carry the id it was handed out under.
        const numFmts = this.numFmts.entries.map((numFmt, index) =>
            numFmt.replace('<numFmt ', `<numFmt numFmtId="${FIRST_CUSTOM_NUMFMT + index}" `),
        );
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            section('numFmts', numFmts) +
            section('fonts', this.fonts.entries) +
            section('fills', this.fills.entries) +
            section('borders', this.borders.entries) +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            section('cellXfs', this.xfs.entries) +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '</styleSheet>'
        );
    }
}
