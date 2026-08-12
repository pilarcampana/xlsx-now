// `Temporal`: the three types a sheet has room for, both ways round.
//
// The runtime this suite runs on may or may not have a `Temporal` of its own
// — Node 26 does, everything before it does not — so `.mocharc.json` loads a
// polyfill before any of this. Which is also the case worth covering: a
// polyfill is what most callers will be on for a while yet, and a package
// that only works on the native one would be a package that works nowhere
// today.
import { execFile } from 'node:child_process';
import assert from 'node:assert/strict';
import { promisify } from 'node:util';
import 'temporal-polyfill/global';
import 'temporal-polyfill/types/global';
import { createXlsxStream } from '../src/core/createXlsxStream.js';
import { localClock, utcClock } from '../src/core/dates.js';
import { readXlsx } from '../src/core/read/readXlsx.js';
import { DateFormats, DEFAULT_DATE_FORMAT, DEFAULT_TIME_FORMAT } from '../src/core/styles.js';
import { temporalApi } from '../src/core/temporal.js';
import type { StyledCell } from '../src/core/types.js';
import {
    BUILTIN_TYPES,
    defaultTypes,
    plainDateTimeValue,
    plainDateValue,
    plainTimeValue,
    withTemporalTypes,
    type ConvertContext,
} from '../src/core/valueTypes.js';
import { collect } from './helpers/streams.js';

const run = promisify(execFile);

const PLAIN: ConvertContext = { dates: new DateFormats(), clock: localClock };

/** A workbook holding one row, as bytes. */
async function written(
    row: readonly unknown[],
    options: Record<string, unknown> = {},
): Promise<Uint8Array> {
    const stream = createXlsxStream({ rows: [row as never[]], ...options });
    return new Uint8Array(await collect(stream));
}

describe('the Temporal types a workbook knows', () => {
    it('are in defaultTypes, and are the only ones added to the built-in three', () => {
        const api = temporalApi();
        assert.ok(api, 'this runtime has no Temporal, and the tests load a polyfill');
        assert.equal(defaultTypes.size, BUILTIN_TYPES.size + 3);
        for (const type of [api.PlainDate, api.PlainDateTime, api.PlainTime]) {
            assert.equal(typeof defaultTypes.get(type)?.convert, 'function');
        }
    });

    it('are not in the built-in three, which is what a runtime without them keeps', () => {
        assert.deepEqual([...BUILTIN_TYPES.keys()], [Date, BigInt, URL]);
    });

    it('leave out the ones a sheet has no room for', () => {
        // A `ZonedDateTime` and an `Instant` carry a zone a cell cannot hold,
        // and a `Duration` is not a point in time at all. Each is a `withType`
        // away for whoever knows which answer they want.
        for (const name of ['ZonedDateTime', 'Instant', 'Duration', 'PlainYearMonth'] as const) {
            assert.equal(defaultTypes.get(Temporal[name]), undefined, name);
        }
    });

    it('can be added again by anyone whose Temporal arrived later than this module', () => {
        const grown = withTemporalTypes(BUILTIN_TYPES);
        assert.equal(grown.size, BUILTIN_TYPES.size + 3);
        // A new map every time: what it was based on is never touched.
        assert.equal(BUILTIN_TYPES.size, 3);
    });
});

describe('writing a Temporal', () => {
    it('writes a PlainDate as a whole serial under the date format', () => {
        const value = plainDateValue(new Temporal.PlainDate(2024, 1, 15), PLAIN);
        assert.equal(value.v, 45306);
        assert.equal(value.numFmt, DEFAULT_DATE_FORMAT);
    });

    it('reads the fields rather than spreading them, which would find none', () => {
        // A `Temporal` keeps its fields as getters on the prototype, so
        // `{ ...plainDate }` is an empty object — and a serial worked out from
        // one would be `NaN` rather than a failure anybody would notice.
        assert.deepEqual({ ...new Temporal.PlainDate(2024, 1, 15) }, {});
        assert.ok(Number.isFinite(plainDateValue(new Temporal.PlainDate(2024, 1, 15), PLAIN).v));
    });

    it('writes a PlainDateTime as the fraction of the day it says', () => {
        const value = plainDateTimeValue(
            new Temporal.PlainDateTime(2024, 1, 15, 10, 30),
            PLAIN,
        );
        assert.equal(value.v, 45306.4375);
        assert.equal(value.numFmt, new DateFormats().dateTime);
    });

    it('writes a PlainDateTime at midnight as the day it amounts to', () => {
        // Which is what makes the round trip close: the whole serial that
        // comes out of it reads back as a `PlainDate`.
        const value = plainDateTimeValue(new Temporal.PlainDateTime(2024, 1, 15), PLAIN);
        assert.equal(value.v, 45306);
        assert.equal(value.numFmt, DEFAULT_DATE_FORMAT);
    });

    it('writes a PlainTime on the day a sheet stores a bare time on', () => {
        const value = plainTimeValue(new Temporal.PlainTime(10, 30), PLAIN);
        assert.equal(value.v, 0.4375);
        assert.equal(value.numFmt, DEFAULT_TIME_FORMAT);
    });

    it('drops the micro- and nanoseconds a serial has no digits left for', () => {
        const value = plainTimeValue(new Temporal.PlainTime(10, 30, 0, 0, 1, 1), PLAIN);
        assert.equal(value.v, 0.4375);
    });

    it('needs no clock, which is the reason to prefer one to a Date', () => {
        const utc: ConvertContext = { dates: new DateFormats(), clock: utcClock };
        const date = new Temporal.PlainDate(2024, 1, 15);
        assert.equal(plainDateValue(date, utc).v, plainDateValue(date, PLAIN).v);
    });
});

describe('a Temporal through a whole workbook', () => {
    it('goes in and comes back as the same value, all three of them', async () => {
        const row = [
            new Temporal.PlainDate(2024, 1, 15),
            new Temporal.PlainDateTime(2024, 7, 1, 10, 30, 45),
            new Temporal.PlainTime(23, 59, 59),
        ];
        const [sheet] = await readXlsx(await written(row));
        assert.deepEqual(sheet?.cells[0]?.map(String), row.map(String));
    });

    it('comes back as the same three types it went in as', async () => {
        const row = [
            new Temporal.PlainDate(2024, 1, 15),
            new Temporal.PlainDateTime(2024, 7, 1, 10, 30),
            new Temporal.PlainTime(23, 59),
        ];
        const [sheet] = await readXlsx(await written(row));
        assert.deepEqual(
            sheet?.cells[0]?.map((value) => Object.getPrototypeOf(value)),
            [
                Temporal.PlainDate.prototype,
                Temporal.PlainDateTime.prototype,
                Temporal.PlainTime.prototype,
            ],
        );
    });

    it('survives a trip through the cells mode, which is what that mode is for', async () => {
        const row = [new Temporal.PlainDate(2024, 1, 15), new Temporal.PlainTime(10, 30)];
        const [read] = await readXlsx(await written(row), { mode: 'cells' });
        const cells = read?.cells[0] as StyledCell[];
        // Straight back into a writer, and out the other side unchanged.
        const [again] = await readXlsx(await written(cells));
        assert.deepEqual(again?.cells[0]?.map(String), row.map(String));
    });

    it('is measured by the format it is shown in, not by the serial under it', async () => {
        // Half past ten as a fraction of a day is twenty characters of a
        // number nobody sees; the column is sized to the `10:30:00` it shows.
        const bytes = await written([new Temporal.PlainTime(10, 30)], { autoWidth: true });
        const text = new TextDecoder().decode(bytes);
        assert.doesNotMatch(text, /width="2[0-9]/);
    });
});

describe('a runtime with no Temporal at all', () => {
    /** The compiled reader, as the child process will import it. */
    const reader = new URL('../src/core/read/readXlsx.js', import.meta.url).href;

    /** Node, with nothing loaded before the script. */
    async function node(script: string): Promise<string> {
        const { stdout } = await run(process.execPath, ['--input-type=module', '-e', script]);
        return stdout.trim();
    }

    before(async function () {
        // On Node 26 and later there is no such runtime to test against.
        if ((await node('process.stdout.write(typeof globalThis.Temporal)')) !== 'undefined') {
            this.skip();
        }
    });

    it('says so when the package is opened, before a single row is read', async () => {
        const message = await node(`
            const { openXlsx } = await import(${JSON.stringify(reader)});
            try {
                // Bytes that are not a package at all: the check comes first,
                // so what fails is the runtime and not the file.
                await openXlsx(new Uint8Array(0));
                process.stdout.write('no error');
            } catch (error) {
                process.stdout.write(error.message);
            }
        `);
        assert.match(message, /needs a Temporal in the runtime/);
        assert.match(message, /temporal-polyfill\/global/);
        assert.match(message, /'localDate'/);
    });

    it('reads dates any other way without complaining', async () => {
        const answer = await node(`
            const { readXlsx } = await import(${JSON.stringify(reader)});
            const { createXlsxStream } = await import(
                ${JSON.stringify(new URL('../src/core/createXlsxStream.js', import.meta.url).href)}
            );
            const chunks = [];
            const stream = createXlsxStream({ rows: [[new Date(2024, 0, 15)]] });
            for await (const chunk of stream) chunks.push(chunk);
            const size = chunks.reduce((total, chunk) => total + chunk.length, 0);
            const bytes = new Uint8Array(size);
            let at = 0;
            for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length; }
            const [sheet] = await readXlsx(bytes, { dates: 'isoString' });
            process.stdout.write(String(sheet.cells[0][0]));
        `);
        assert.equal(answer, '2024-01-15');
    });
});
