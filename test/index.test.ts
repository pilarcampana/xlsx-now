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
        assert.equal(core.DATE_FORMAT, 'yyyy-mm-dd');
        assert.equal(core.DATETIME_FORMAT, 'yyyy-mm-dd hh:mm:ss');
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

    it('exports nothing else', () => {
        assert.deepEqual(Object.keys(core).sort(), [
            'DATETIME_FORMAT',
            'DATE_FORMAT',
            'DEFAULT_COMPRESSION_LEVEL',
            'LINE',
            'StyleTable',
            'WORKSHEET',
            'XlsxStream',
            'XlsxWriter',
            'ZipWriter',
            'argb',
            'createXlsxStream',
            'isLineCommand',
            'isWorksheetCommand',
        ]);
    });
});
