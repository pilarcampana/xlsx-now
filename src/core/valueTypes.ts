// The types a workbook knows how to write, and the map that says so.
//
// A sheet holds four things — a number, a boolean, a string, or nothing — and
// everything else has to become one of them on the way in. `Date` was the
// first value that needed it, and for a while it was written into every module
// that had to ask what a cell held: the type of the `<c>`, the text of the
// `<v>`, the number format the cell falls back to, and the width the column is
// measured to. Four places, one type.
//
// This is that question asked once. A value the writer does not already know
// is looked up by the class it is an instance of, and what comes back says how
// to write it — in terms of what the writer already knew all along. `Date` is
// the first entry of the map rather than a case above it, which is the whole
// test of whether this generalizes: a `Temporal.PlainDate`, a `BigInt`, or a
// class of the caller's own goes in the same way and costs the same.
import { excelSerial } from './cell.js';
import { DEFAULT_DATE_FORMATS, type DateFormats } from './styles.js';
import type { CellType } from './types.js';

/**
 * A value as a cell holds it, with nothing standing in for it: what `cellXml`
 * writes without asking anyone. It is deliberately narrower than `CellValue` —
 * a `Date` is not one of these — because it is what a conversion produces, and
 * a conversion that could produce something needing conversion would be a
 * conversion that never ends.
 */
export type NativeValue = string | number | boolean | null | undefined;

/** What the workbook lends a conversion. */
export interface ConvertContext {
    /**
     * The formats the workbook's `dateFormat` and `dateTimeFormat` settled on.
     * A type that is a date in any sense reads them rather than inventing its
     * own, which is what keeps one workbook writing its dates one way.
     */
    readonly dates: DateFormats;
}

/**
 * A value the writer did not know, said in terms it does.
 *
 * Only `v` is required. The rest is what the value would have lost by becoming
 * a native one, handed back so the cell can carry it:
 *
 * ```js
 * { v: 45306 }                                  // a plain number, nothing more
 * { v: 45306, numFmt: 14, width: 10 }           // a date: a serial, shown as one
 * { v: '9007199254740993', t: 'inlineStr' }     // a number too big to stay one
 * ```
 */
export interface ConvertedValue {
    /** The value itself, as the writer already writes it. */
    v: NativeValue;
    /**
     * What the cell says it holds. Read off `v` when it is left out — and
     * overridden by a `t` the caller wrote on the cell, since asking for a
     * type outright is asking for that one.
     */
    t?: CellType;
    /**
     * The number format the cell falls back to. It applies exactly as a date's
     * does: a style that already says `numFmt` has said what it wants and is
     * left alone.
     */
    numFmt?: string | number;
    /**
     * How many characters this shows, for the column that is sizing itself.
     * Left out, it is worked out from the `numFmt` above, and failing that
     * `v` is measured as it is written — see `shownWidth`.
     */
    width?: number;
}

/**
 * How many characters a converted value shows: what it said, or what its
 * format says for it.
 *
 * A conversion that gave a `numFmt` has already said that `v` is not what the
 * cell shows — asking for a format is what that means — so measuring `v`
 * there is not measuring imprecisely, it is measuring the wrong thing. Half
 * an hour written as a fraction of a day is `0.020833333333333332`, twenty
 * characters of a number nobody will ever see, against the seven of the
 * `0:30:00` the cell actually shows.
 *
 * A format code is a template of what comes out, so its own length is the
 * estimate — which is what `DateFormats` already measures a spelled-out date
 * by. It is an estimate and not a measurement: a format whose output grows
 * with the value, like `#,##0.00` against a million, comes out short. That is
 * what `width` is for, and a conversion that knows its own magnitude says it.
 *
 * A built-in format is an id with no code to measure, so there is nothing to
 * work out from it and the type that used one says `width` itself — which is
 * what `dateValue` does.
 */
export function shownWidth(value: ConvertedValue | undefined): number | undefined {
    if (value === undefined) return undefined;
    if (value.width !== undefined) return value.width;
    return typeof value.numFmt === 'string' ? value.numFmt.length : undefined;
}

/**
 * How one type becomes a value. `convert` runs per value; which handler runs
 * is decided by the type alone, once, and remembered.
 */
export interface TypeHandler<T> {
    convert(value: T, context: ConvertContext): ConvertedValue;
}

/**
 * A type, as the map names it: the class itself — `Date`, `URL`, `BigInt`, a
 * class of the caller's own. Not the name of the class and not its prototype,
 * so nothing here depends on a name a minifier is free to rewrite.
 */
export interface TypeKey {
    readonly prototype: unknown;
}

/**
 * A handler as the map holds it, with the type it was registered for
 * forgotten. `never` is what makes every `TypeHandler<T>` fit in one map;
 * the key a handler was found under is what makes calling it back true.
 */
export type RegisteredHandler = TypeHandler<never>;

/**
 * Every type a workbook knows, in one value.
 *
 * A map and not a registry with a `register` on it, and the difference is the
 * point: there is no moment at which types are added, so there is no order for
 * anything to depend on. A workbook is handed the map it will use and reads it
 * once. Two workbooks written side by side cannot change what the other one
 * knows, which is what a mutable global would have made possible.
 */
export type TypeMap = ReadonlyMap<TypeKey, RegisteredHandler>;

/**
 * A `Date` as a cell holds it: the serial number, under the format the
 * workbook shows its dates in, measured as that format writes them.
 *
 * Exported because it is what a caller's own date-like type delegates to —
 * `convert: (own, context) => dateValue(own.toDate(), context)` is the whole
 * of a handler for a class that wraps a `Date`.
 */
export function dateValue(value: Date, { dates }: ConvertContext): ConvertedValue {
    return { v: excelSerial(value), numFmt: dates.for(value), width: dates.textLength(value) };
}

/**
 * The largest whole number a sheet can hold, as a `bigint`. A cell stores a
 * double, so 15 to 16 digits is all the precision there is: past this, the
 * number that comes back out of the file is not the one that went in.
 */
const MAX_EXACT_INTEGER = BigInt(Number.MAX_SAFE_INTEGER);

/**
 * A `BigInt` as a cell holds it: a number while a number can still hold it,
 * and text once it cannot.
 *
 * The alternative is writing every one of them as a number and letting the
 * long ones come back with their last digits replaced by zeros — a wrong
 * value that looks like a right one, which is the one outcome worth going out
 * of the way to avoid. Text keeps every digit, and a caller who would rather
 * have the rounded number can say `t: 'n'` on the cell.
 */
export function bigintValue(value: bigint): ConvertedValue {
    if (value > MAX_EXACT_INTEGER || value < -MAX_EXACT_INTEGER) {
        return { v: String(value), t: 'inlineStr' };
    }
    return { v: Number(value) };
}

/** A `URL` as its text. */
export function urlValue(value: URL): ConvertedValue {
    return { v: value.href };
}

/**
 * The types a workbook knows when it was told nothing.
 *
 * `Date` is one entry among the others on purpose. What is not here is as much
 * of a decision: a `Map`, a `Set`, a nested array or a class nobody registered
 * has no one right way to be written down, so it is refused by name rather
 * than written out as whatever `String()` makes of it.
 */
export const defaultTypes: TypeMap = new Map<TypeKey, RegisteredHandler>([
    [Date, { convert: dateValue }],
    [BigInt, { convert: bigintValue }],
    [URL, { convert: urlValue }],
]);

/**
 * The same types, and one more.
 *
 * ```js
 * export const appTypes = withType(defaultTypes, HourRange, {
 *     convert: (range) => ({ v: range.toString() }),
 * });
 * ```
 *
 * A new map every time: what it is based on is never touched, so a module can
 * hand `defaultTypes` around without anyone being able to change it. Handing
 * the result to a writer is one `types` option, however many types were added
 * to it — which is what this is for, rather than naming them one by one at
 * every place a workbook is written.
 *
 * `new Map(base).set(Type, handler)` does the same thing; this is the spelling
 * that keeps the type, so `convert` is written against the class it is for and
 * not against `unknown`.
 */
export function withType<T>(
    base: TypeMap,
    type: { readonly prototype: T },
    handler: TypeHandler<T>,
): TypeMap {
    return new Map(base).set(type, handler);
}

/**
 * The classes that stand for a primitive: a `bigint` is not an instance of
 * anything, but `BigInt` is what it would be written as, so that is what the
 * map is keyed on and this is what turns a `typeof` back into it.
 */
const PRIMITIVE_TYPES: ReadonlyMap<string, TypeKey> = new Map([['bigint', BigInt]]);

/**
 * A `TypeMap` as a workbook reads it: the same types, indexed by what a value
 * is actually looked up on.
 *
 * The map is keyed by the class because that is what a caller can write down.
 * A value carries its prototype, not its class, and the two are one step
 * apart — so the index is built once, here, and never per cell. Past that,
 * every prototype is remembered the first time it is seen, which is what keeps
 * a sheet of a million dates to one lookup and 999,999 hits.
 */
export class ValueTypes {
    /** `prototype` -> handler, for every entry but `Object`. */
    private readonly byPrototype = new Map<object, RegisteredHandler>();
    /** `typeof` -> handler, for the entries that stand for a primitive. */
    private readonly byPrimitive = new Map<string, RegisteredHandler>();
    /**
     * The `Object` entry: what an object nobody claimed becomes.
     *
     * It is held apart from the index rather than found by walking to
     * `Object.prototype`, because the walk would reach it before anything had
     * asked whether the object was a cell — and `{ v: 1, s: 'money' }` is an
     * object whose prototype is `Object.prototype` too. So the order is fixed
     * where it can be read: a registered type is a value, then a cell that
     * says more is a cell, and only what is neither gets here.
     */
    readonly objectHandler: RegisteredHandler | undefined;
    /** What the conversions are handed; one per workbook. */
    private readonly context: ConvertContext;
    /** `prototype` -> the handler it resolved to, or `null` for none. */
    private readonly resolved = new WeakMap<object, RegisteredHandler | null>();

    constructor(
        types: TypeMap = defaultTypes,
        context: ConvertContext = { dates: DEFAULT_DATE_FORMATS },
    ) {
        let objectHandler: RegisteredHandler | undefined;
        for (const [type, handler] of types) {
            if (type === Object) {
                objectHandler = handler;
                continue;
            }
            const primitive = primitiveOf(type);
            if (primitive !== undefined) this.byPrimitive.set(primitive, handler);
            const { prototype } = type;
            // A class has an object for a prototype; anything else is a key
            // that no value can ever be an instance of, and the primitive
            // table above is the only reason it would be in the map at all.
            if (typeof prototype === 'object' && prototype !== null) {
                this.byPrototype.set(prototype, handler);
            }
        }
        this.objectHandler = objectHandler;
        this.context = context;
    }

    /**
     * The handler for an object, or `undefined` when nothing claims it.
     *
     * The prototype chain is walked, so a subclass is written as whatever its
     * base was registered as — `class Timestamp extends Date` needs no entry
     * of its own. `Object.prototype` is where the walk stops without being
     * consulted: see `objectHandler`.
     */
    handlerFor(value: object): RegisteredHandler | undefined {
        const prototype: unknown = Object.getPrototypeOf(value);
        if (typeof prototype !== 'object' || prototype === null) return undefined;
        const known = this.resolved.get(prototype);
        if (known !== undefined) return known ?? undefined;
        const handler = this.walk(prototype);
        this.resolved.set(prototype, handler ?? null);
        return handler;
    }

    /** The chain, one step at a time, up to but not including `Object.prototype`. */
    private walk(from: object): RegisteredHandler | undefined {
        let prototype: object | null = from;
        while (prototype !== null && prototype !== Object.prototype) {
            const handler = this.byPrototype.get(prototype);
            if (handler !== undefined) return handler;
            prototype = Object.getPrototypeOf(prototype) as object | null;
        }
        return undefined;
    }

    /**
     * What this value is written as, or `undefined` when the writer already
     * knew — which is every string, every number and every boolean, so the
     * common cell reaches nothing here at all.
     *
     * An object with nothing to claim it falls to the `Object` entry, which by
     * then is the only thing it can be: `isStyledCell` has already refused the
     * objects that were meant to be cells.
     */
    convert(value: unknown): ConvertedValue | undefined {
        if (value === null || value === undefined) return undefined;
        if (typeof value === 'object') {
            const handler = this.handlerFor(value) ?? this.objectHandler;
            return handler && call(handler, value, this.context);
        }
        const handler = this.byPrimitive.get(typeof value);
        return handler && call(handler, value, this.context);
    }
}

/**
 * A handler, called with the value it was found for. The type it was
 * registered under was forgotten to get it into the map, and the lookup that
 * produced it is what stands in for the check here.
 */
function call(
    handler: RegisteredHandler,
    value: unknown,
    context: ConvertContext,
): ConvertedValue {
    return (handler as TypeHandler<unknown>).convert(value, context);
}

/** The `typeof` a key stands for, when it stands for a primitive at all. */
function primitiveOf(type: TypeKey): string | undefined {
    for (const [primitive, key] of PRIMITIVE_TYPES) {
        if (key === type) return primitive;
    }
    return undefined;
}
