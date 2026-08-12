// The public surface. Nothing here exercises behaviour — the point is that
// what the package promises to export is what `src/core/index.ts` re-exports,
// under the names the README uses.
import assert from 'node:assert/strict';
import * as core from '../src/core/index.js';

describe('the public surface', () => {
    it('exports the three ways to drive the writer', () => {
        assert.equal(typeof core.XlsxWriter, 'function');
        assert.equal(typeof core.XlsxStream, 'function');
        assert.equal(typeof core.createXlsxStream, 'function');
    });

    it('exports the style table and the zip writer under it', () => {
        assert.equal(typeof core.StyleTable, 'function');
        assert.equal(typeof core.argb, 'function');
        assert.equal(typeof core.ZipWriter, 'function');
        assert.equal(typeof core.DateFormats, 'function');
        assert.equal(core.DEFAULT_DATE_FORMAT, 14);
        assert.equal(core.DEFAULT_DATETIME_FORMAT, 22);
        assert.equal(core.DEFAULT_TIME_FORMAT, 21);
        assert.equal(core.DEFAULT_COMPRESSION_LEVEL, 6);
    });

    it('exports the date arithmetic, and the two clocks a `Date` is read with', () => {
        assert.equal(typeof core.excelSerial, 'function');
        assert.equal(typeof core.fromExcelSerial, 'function');
        assert.equal(typeof core.serialOfParts, 'function');
        assert.equal(typeof core.partsOfSerial, 'function');
        assert.equal(typeof core.hasTimeOfDay, 'function');
        assert.equal(typeof core.kindOf, 'function');
        assert.equal(core.dateClocks.local, core.localClock);
        assert.equal(core.dateClocks.utc, core.utcClock);
        assert.equal(core.clockOf(undefined), core.DEFAULT_CLOCK);
    });

    it('exports the readers a date can come back as, and the way to pick one', () => {
        assert.equal(core.DEFAULT_DATE_READER, 'temporal');
        assert.equal(core.dateReaders.temporal, core.temporalDates);
        assert.equal(core.dateReaders.localDate, core.localDates);
        assert.equal(core.dateReaders.utcDate, core.utcDates);
        assert.equal(core.dateReaders.isoString, core.isoDates);
        assert.equal(core.dateReaders.serial, core.serialDates);
        assert.equal(core.dateReaderOf('isoString'), core.isoDates);
        assert.equal(core.dateReaderOf(undefined), core.temporalDates);
    });

    it('exports the command keys, and the guards that recognize one', () => {
        assert.equal(core.WORKSHEET, '#worksheet');
        assert.equal(core.LINE, '#line');
        assert.equal(core.isWorksheetCommand({ '#worksheet': 'Sheet2' }), true);
        assert.equal(core.isWorksheetCommand({ name: 'Ana' }), false);
        assert.equal(core.isLineCommand({ '#line': 'empty' }), true);
        assert.equal(core.isLineCommand({ name: 'Ana' }), false);
    });

    it('exports the types a workbook knows, and the way to add one', () => {
        assert.equal(typeof core.withType, 'function');
        assert.equal(typeof core.defaultTypes.get(Date)?.convert, 'function');
        assert.equal(typeof core.dateValue, 'function');
        assert.equal(typeof core.bigintValue, 'function');
        assert.equal(typeof core.urlValue, 'function');
        assert.equal(typeof core.shownWidth, 'function');
        assert.equal(typeof core.partsValue, 'function');
        assert.equal(typeof core.BUILTIN_TYPES.get(Date)?.convert, 'function');
        assert.equal(typeof core.withTemporalTypes, 'function');
    });

    it('exports the `Temporal` conversions, and the way to see if there is one', () => {
        assert.equal(typeof core.temporalApi, 'function');
        assert.equal(typeof core.plainDateValue, 'function');
        assert.equal(typeof core.plainDateTimeValue, 'function');
        assert.equal(typeof core.plainTimeValue, 'function');
    });

    it('exports the reader, and the way to hand it something to seek in', () => {
        assert.equal(typeof core.readXlsx, 'function');
        assert.equal(typeof core.openXlsx, 'function');
        assert.equal(typeof core.bytesAccess, 'function');
    });

    it('exports nothing else', () => {
        assert.deepEqual(Object.keys(core).sort(), [
            'BUILTIN_TYPES',
            'DEFAULT_CLOCK',
            'DEFAULT_COMPRESSION_LEVEL',
            'DEFAULT_DATETIME_FORMAT',
            'DEFAULT_DATE_FORMAT',
            'DEFAULT_DATE_READER',
            'DEFAULT_TIME_FORMAT',
            'DateFormats',
            'LINE',
            'StyleTable',
            'WORKSHEET',
            'XlsxStream',
            'XlsxWriter',
            'ZipWriter',
            'argb',
            'bigintValue',
            'bytesAccess',
            'clockOf',
            'createXlsxStream',
            'dateClocks',
            'dateReaderOf',
            'dateReaders',
            'dateValue',
            'defaultTypes',
            'excelSerial',
            'fromExcelSerial',
            'hasTimeOfDay',
            'isLineCommand',
            'isWorksheetCommand',
            'isoDates',
            'kindOf',
            'localClock',
            'localDates',
            'openXlsx',
            'partsOfSerial',
            'partsValue',
            'plainDateTimeValue',
            'plainDateValue',
            'plainTimeValue',
            'readXlsx',
            'serialDates',
            'serialOfParts',
            'shownWidth',
            'temporalApi',
            'temporalDates',
            'urlValue',
            'utcClock',
            'utcDates',
            'withTemporalTypes',
            'withType',
        ]);
    });
});
