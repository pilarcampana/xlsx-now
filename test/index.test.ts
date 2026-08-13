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
        assert.equal(core.DEFAULT_COMPRESSION_LEVEL, 6);
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
        assert.equal(typeof core.defaultTypes().get(Date)?.convert, 'function');
        assert.equal(typeof core.dateValue, 'function');
        assert.equal(typeof core.bigintValue, 'function');
        assert.equal(typeof core.urlValue, 'function');
        assert.equal(typeof core.shownWidth, 'function');
    });

    it('exports the reader, and the way to hand it something to seek in', () => {
        assert.equal(typeof core.readXlsx, 'function');
        assert.equal(typeof core.openXlsx, 'function');
        assert.equal(typeof core.bytesAccess, 'function');
    });

    it('exports what a date is made of, on both sides of the file', () => {
        assert.equal(core.excelSerial(new Date(1970, 0, 1)), 25569);
        assert.equal(core.fromExcelSerialUtc(25569).toISOString(), '1970-01-01T00:00:00.000Z');
        assert.equal(core.fromExcelSerial(25569).getFullYear(), 1970);
        assert.equal(core.serialKind(45306.5), 'dateTime');
        assert.equal(core.readDate(45306, 'isoString'), '2024-01-15');
        assert.equal(typeof core.temporalApi, 'function');
        assert.equal(typeof core.plainDateValue, 'function');
        assert.equal(typeof core.plainDateTimeValue, 'function');
        assert.equal(typeof core.plainTimeValue, 'function');
        assert.equal(typeof core.serialValue, 'function');
    });

    it('exports nothing else', () => {
        assert.deepEqual(Object.keys(core).sort(), [
            'DEFAULT_COMPRESSION_LEVEL',
            'DEFAULT_DATETIME_FORMAT',
            'DEFAULT_DATE_FORMAT',
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
            'createXlsxStream',
            'dateValue',
            'defaultTypes',
            'excelSerial',
            'fromExcelSerial',
            'fromExcelSerialUtc',
            'isLineCommand',
            'isWorksheetCommand',
            'openXlsx',
            'plainDateTimeValue',
            'plainDateValue',
            'plainTimeValue',
            'readDate',
            'readXlsx',
            'serialKind',
            'serialValue',
            'shownWidth',
            'temporalApi',
            'urlValue',
            'withType',
        ]);
    });
});
