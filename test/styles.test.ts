import assert from 'node:assert/strict';
import { STYLE, styleIndex, stylesXml } from '../src/core/styles.js';

/** A `<name count="n">...</name>` section: what it says, and what it holds. */
function section(xml: string, name: string, child: string): { declared: number; found: number } {
    const found = new RegExp(`<${name} count="(\\d+)">(.*)</${name}>`).exec(xml);
    assert.ok(found, `no <${name}> section`);
    const inner = found[2] ?? '';
    return {
        declared: Number(found[1]),
        // Opening tags of `child` only: the lookahead is what keeps <fonts>
        // from counting as a <font> and </font> from counting at all.
        found: (inner.match(new RegExp(`<${child}(?=[ />])`, 'g')) ?? []).length,
    };
}

describe('styleIndex', () => {
    it('is the default without a style', () => {
        assert.equal(styleIndex(undefined), STYLE.DEFAULT);
        assert.equal(styleIndex({}), STYLE.DEFAULT);
    });

    it('is one entry per combination of the two flags', () => {
        assert.equal(styleIndex({ bold: true }), STYLE.BOLD);
        assert.equal(styleIndex({ highlight: true }), STYLE.HIGHLIGHT);
        assert.equal(styleIndex({ bold: true, highlight: true }), STYLE.BOLD_HIGHLIGHT);
    });

    it('reads a false flag as absent', () => {
        assert.equal(styleIndex({ bold: false, highlight: true }), STYLE.HIGHLIGHT);
        assert.equal(styleIndex({ bold: false, highlight: false }), STYLE.DEFAULT);
    });

    it('is the bitmask the table is laid out as', () => {
        assert.deepEqual(
            [STYLE.DEFAULT, STYLE.BOLD, STYLE.HIGHLIGHT, STYLE.BOLD_HIGHLIGHT],
            [0, 1, 2, 3],
        );
    });
});

describe('stylesXml', () => {
    const xml = stylesXml();

    it('declares as many entries as it writes', () => {
        for (const [name, child] of [
            ['fonts', 'font'],
            ['fills', 'fill'],
            ['borders', 'border'],
            ['cellStyleXfs', 'xf'],
            ['cellXfs', 'xf'],
            ['cellStyles', 'cellStyle'],
        ] as const) {
            const { declared, found } = section(xml, name, child);
            assert.equal(found, declared, `<${name}> declares ${declared} and holds ${found}`);
        }
    });

    it('has a cellXfs entry for every index a cell can ask for', () => {
        const entries = section(xml, 'cellXfs', 'xf').declared;
        for (const index of Object.values(STYLE)) {
            assert.ok(
                index < entries,
                `style index ${index} is past the ${entries} entries in cellXfs`,
            );
        }
    });

    it('lays the cellXfs entries out in the order the bitmask implies', () => {
        const cellXfs = /<cellXfs count="\d+">(.*)<\/cellXfs>/.exec(xml)?.[1] ?? '';
        const entries = cellXfs.match(/<xf [^>]*\/>/g) ?? [];
        const bold = (entry: string): boolean => entry.includes('fontId="1"');
        const highlight = (entry: string): boolean => entry.includes('fillId="2"');

        assert.equal(entries.length, 4);
        assert.deepEqual(entries.map(bold), [false, true, false, true]);
        assert.deepEqual(entries.map(highlight), [false, false, true, true]);
    });

    it('carries the bold font and the pk fill the two flags mean', () => {
        assert.ok(xml.includes('<font><b/>'), 'no bold font');
        assert.ok(xml.includes('<fgColor rgb="FFFFE699"/>'), 'no pk fill');
    });

    it('is a standalone styleSheet part', () => {
        assert.ok(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'));
        assert.ok(xml.endsWith('</styleSheet>'));
    });
});
