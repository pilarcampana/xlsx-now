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
        assert.equal(typeof core.styleIndex, 'function');
        assert.equal(typeof core.stylesXml, 'function');
        assert.equal(typeof core.ZipWriter, 'function');
        assert.deepEqual(core.STYLE, { DEFAULT: 0, BOLD: 1, HIGHLIGHT: 2, BOLD_HIGHLIGHT: 3 });
        assert.equal(core.DEFAULT_COMPRESSION_LEVEL, 6);
    });

    it('exports the command key, and the guard that recognizes one', () => {
        assert.equal(core.WORKSHEET, '#worksheet');
        assert.equal(core.isWorksheetCommand({ '#worksheet': 'Sheet2' }), true);
        assert.equal(core.isWorksheetCommand({ name: 'Ana' }), false);
    });

    it('exports nothing else', () => {
        assert.deepEqual(Object.keys(core).sort(), [
            'DEFAULT_COMPRESSION_LEVEL',
            'STYLE',
            'WORKSHEET',
            'XlsxStream',
            'XlsxWriter',
            'ZipWriter',
            'createXlsxStream',
            'isWorksheetCommand',
            'styleIndex',
            'stylesXml',
        ]);
    });
});
