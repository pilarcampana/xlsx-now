import assert from 'node:assert/strict';
import { readSharedStrings } from '../../src/core/read/sharedStrings.js';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sst = (body: string): string => `<sst xmlns="${NS}">${body}</sst>`;

describe('readSharedStrings', () => {
    it('is the table in order, one string per entry', () => {
        assert.deepEqual(readSharedStrings(sst('<si><t>uno</t></si><si><t>dos</t></si>')), [
            'uno',
            'dos',
        ]);
    });

    it('joins the runs of an entry that carries its own formatting', () => {
        assert.deepEqual(
            readSharedStrings(sst('<si><r><t>en </t></r><r><rPr/><t>negrita</t></r></si>')),
            ['en negrita'],
        );
    });

    it('keeps the spaces a run was told to preserve', () => {
        assert.deepEqual(readSharedStrings(sst('<si><t xml:space="preserve">  ab  </t></si>')), [
            '  ab  ',
        ]);
    });

    it('leaves the phonetic guide of an entry out of its text', () => {
        // The reading of a word, which a Japanese entry carries alongside it
        // in `<t>` elements of its own.
        assert.deepEqual(
            readSharedStrings(sst('<si><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh></si>')),
            ['東京'],
        );
    });

    it('reads an empty entry as an empty string', () => {
        assert.deepEqual(readSharedStrings(sst('<si/><si><t></t></si>')), ['', '']);
    });

    it('decodes the entities a string was written with', () => {
        assert.deepEqual(readSharedStrings(sst('<si><t>a &amp; b &lt; c &#233;</t></si>')), [
            'a & b < c é',
        ]);
    });

    it('reads a table with nothing in it', () => {
        assert.deepEqual(readSharedStrings(sst('')), []);
    });

    it('refuses a part that does not parse', () => {
        assert.throws(() => readSharedStrings('<sst><si><t>a</si></sst>'));
    });
});
