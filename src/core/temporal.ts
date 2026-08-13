// `Temporal`, from a package that does not depend on it.
//
// The reader builds `Temporal.PlainDate`, `PlainDateTime` and `PlainTime` when
// it is asked to, and the writer takes the three of them back. Nothing else of
// that API is used: `from` to build one, `toString` to read it, and the classes
// themselves for the type map to be keyed on.
//
// That much of it is written out here rather than imported, because there is
// nothing to import from — `Temporal` is a global, TypeScript's own libraries
// do not declare it yet, and a polyfill is a dependency this package does not
// have and should not force. So the global is read where it is needed, and
// checked before it is: a workbook opened with `dates: 'temporal'` in an
// environment with no `Temporal` fails at `openXlsx`, saying so, rather than
// with a `TypeError` in the middle of a sheet.
//
// The types below are what this package touches of each class, so a real
// Temporal value fits them. They are deliberately not a copy of the whole API:
// what is declared here is what is used here.

/** A date with no time of day and no time zone — what a date cell holds. */
export interface PlainDate {
    readonly year: number;
    readonly month: number;
    readonly day: number;
    /** The ISO date: `2024-01-15`. */
    toString(): string;
}

/** A date and a time of day, with no time zone. */
export interface PlainDateTime extends PlainDate {
    readonly hour: number;
    readonly minute: number;
    readonly second: number;
    readonly millisecond: number;
    /** The ISO date and time: `2024-01-15T10:30:00`. */
    toString(): string;
}

/** A time of day, with no date and no time zone. */
export interface PlainTime {
    readonly hour: number;
    readonly minute: number;
    readonly second: number;
    readonly millisecond: number;
    /** The ISO time: `10:30:00`. */
    toString(): string;
}

/**
 * One of the three classes, as this package uses it: the `from` that builds a
 * value out of its ISO text, and the `prototype` the type map is keyed on.
 */
export interface TemporalClass<T> {
    from(text: string): T;
    readonly prototype: T;
}

/** The corner of `Temporal` this package uses. */
export interface TemporalApi {
    readonly PlainDate: TemporalClass<PlainDate>;
    readonly PlainDateTime: TemporalClass<PlainDateTime>;
    readonly PlainTime: TemporalClass<PlainTime>;
}

/**
 * The environment's `Temporal`, or `undefined` where there is none.
 *
 * Read once, when this module loads, because that is all there is to read: a
 * `Temporal` is native or it is a polyfill, a polyfill is installed by
 * importing it, and an import runs before anything that could use what it
 * installs. So it is there from the start or it is never there, and looking
 * again per date would be looking for something that cannot arrive.
 */
export const temporalApi: TemporalApi | undefined = globalTemporal();

function globalTemporal(): TemporalApi | undefined {
    const temporal = (globalThis as { Temporal?: TemporalApi }).Temporal;
    return temporal?.PlainDate === undefined ? undefined : temporal;
}

/** The same, for whoever cannot go on without it. */
export function requireTemporal(): TemporalApi {
    if (temporalApi === undefined) {
        throw new Error(
            'Dates are read as Temporal values and this environment has no Temporal.PlainDate: ' +
                'run where there is one, install a polyfill, or say how dates should be built ' +
                'with dates: "utcDate", "localDate" or "isoString".',
        );
    }
    return temporalApi;
}
