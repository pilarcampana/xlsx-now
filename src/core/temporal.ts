// `Temporal`, named without depending on having it.
//
// Two problems, and they are not the same one. At *runtime*, `Temporal` is a
// global that a given engine either has or does not: Node 26 has it, Node 22
// has it behind a flag, and a browser that is a year behind has it through a
// polyfill or not at all. At *compile time*, `Temporal` is a type that a
// given `lib` either declares or does not: TypeScript 6 added
// `lib.esnext.temporal`, and everything before it has never heard of the
// name.
//
// The runtime side is a lookup, and it is at the bottom of this file. The
// compile-time side is what the rest of it is for, and the constraint it has
// to meet is worth stating: this package's `.d.ts` has to keep compiling for
// somebody on TypeScript 5.9, whose `lib` has no `Temporal` in it, while
// still handing somebody on TypeScript 6 or 7 the real `Temporal.PlainDate`
// rather than a look-alike. Writing `Temporal.PlainDate` in a declaration
// fails the first; declaring our own shapes and stopping there fails the
// second, since a value typed as a look-alike cannot be assigned to a
// `Temporal.PlainDate` without a cast.
//
// So the names below are conditional: they resolve to the real types when the
// `lib` compiling them has `Temporal`, and to shapes of our own when it does
// not. Which one it is depends on the consumer's `lib` and not on ours, and
// it is settled where it should be — in the consumer's compilation.

/**
 * The global `Temporal` as the compiling `lib` declares it, or `undefined`
 * when it declares no such thing.
 *
 * `typeof globalThis` is what makes this answerable: TypeScript builds it out
 * of the global declarations actually in scope, so asking whether it has a
 * `Temporal` is asking whether the `lib` in force declared one — a question
 * with an answer at every version, rather than an error at the older ones.
 */
type TemporalGlobal = typeof globalThis extends { Temporal: infer T } ? T : undefined;

/**
 * The instance type of one of `Temporal`'s constructors, when there is a
 * `Temporal` to take it from, and a shape of our own when there is not.
 *
 * The real one is read off the constructor's `prototype` rather than through
 * `InstanceType`, because `lib.esnext.temporal` declares each of these as an
 * *interface* with a `var` constructor next to it, not as a class — so
 * `Temporal.PlainDate` the type and `Temporal.PlainDate` the value are two
 * different declarations, and only the second one is reachable from
 * `typeof globalThis`.
 */
type FromTemporal<Name extends string, Fallback> = TemporalGlobal extends {
    [K in Name]: { readonly prototype: infer Instance };
}
    ? Instance
    : Fallback;

/**
 * What we read off a `Temporal.PlainDate`, for a compilation that has no
 * `Temporal` to read it off.
 *
 * Deliberately only the fields this package touches. It is a fallback for
 * naming a value, not a re-declaration of the proposal: somebody who wants
 * `.add()` and `.until()` wants the real thing, and the real thing is what
 * they get as soon as their `lib` has it.
 */
export interface PlainDateShape {
    readonly year: number;
    readonly month: number;
    readonly day: number;
}

/** The same, for a `Temporal.PlainTime`. */
export interface PlainTimeShape {
    readonly hour: number;
    readonly minute: number;
    readonly second: number;
    readonly millisecond: number;
}

/** The same, for a `Temporal.PlainDateTime`, which is both of the above. */
export interface PlainDateTimeShape extends PlainDateShape, PlainTimeShape {}

/** `Temporal.PlainDate`, or our shape of one where there is no `Temporal`. */
export type PlainDate = FromTemporal<'PlainDate', PlainDateShape>;
/** `Temporal.PlainDateTime`, under the same rule. */
export type PlainDateTime = FromTemporal<'PlainDateTime', PlainDateTimeShape>;
/** `Temporal.PlainTime`, under the same rule. */
export type PlainTime = FromTemporal<'PlainTime', PlainTimeShape>;

/**
 * What a date read as `Temporal` comes back as: the day, the day and the
 * time, or the time on its own. Which of the three it is depends on the
 * value and the format it is under — see `temporalDates`.
 */
export type TemporalDate = PlainDate | PlainDateTime | PlainTime;

/**
 * The part of the `Temporal` namespace this package uses, at runtime.
 *
 * Three constructors and nothing else, because three is all that is called.
 * It is written out rather than taken from `TemporalGlobal` so that the
 * lookup below has something to name on every TypeScript version — and the
 * instance types in it are the conditional ones, so what comes out of a
 * `new` here is a real `Temporal.PlainDate` wherever the compiler knows what
 * that is.
 */
export interface TemporalApi {
    readonly PlainDate: {
        new (year: number, month: number, day: number): PlainDate;
        readonly prototype: PlainDate;
    };
    readonly PlainDateTime: {
        new (
            year: number,
            month: number,
            day: number,
            hour?: number,
            minute?: number,
            second?: number,
            millisecond?: number,
        ): PlainDateTime;
        readonly prototype: PlainDateTime;
    };
    readonly PlainTime: {
        new (
            hour?: number,
            minute?: number,
            second?: number,
            millisecond?: number,
        ): PlainTime;
        readonly prototype: PlainTime;
    };
}

/** The `Temporal` found so far, once there has been one to find. */
let found: TemporalApi | undefined;

/**
 * The `Temporal` this runtime has, or `undefined`.
 *
 * Read out of `globalThis` on every call until it is there, and remembered
 * once it is. That asymmetry is the whole point: a polyfill is a module, a
 * module runs when it is imported, and nothing says it is imported before
 * this file — remembering a *missing* `Temporal` would make the order of two
 * imports decide whether the package works. Remembering a found one costs
 * nothing and cannot be wrong, since no runtime takes `Temporal` away again.
 */
export function temporalApi(): TemporalApi | undefined {
    if (found !== undefined) return found;
    const global = (globalThis as { Temporal?: unknown }).Temporal;
    if (global === null || typeof global !== 'object') return undefined;
    return (found = global as TemporalApi);
}

/**
 * The same, or the error that says what to do about it.
 *
 * Asked once when a package is opened rather than once per cell, so a
 * workbook full of dates fails before its first row instead of somewhere in
 * the middle of it — and so a workbook with no dates in it fails too, which
 * is the point: whether the code runs at all should not depend on whether the
 * file that arrived happened to have a date in it.
 */
export function requireTemporal(): TemporalApi {
    const api = temporalApi();
    if (api === undefined) {
        throw new Error(
            'Reading dates as Temporal needs a Temporal in the runtime, and this one has none. ' +
                'Node 26 and later have it; before that, import a polyfill such as ' +
                '"temporal-polyfill/global" before opening a workbook, or ask for another ' +
                "date type: { dates: 'localDate' } is what a Date was always read as.",
        );
    }
    return api;
}
