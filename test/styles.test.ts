import assert from 'node:assert/strict';
import {
    DATETIME_FORMAT,
    DATE_FORMAT,
    StyleTable,
    argb,
    type StyleSpec,
} from '../src/core/styles.js';

/** A `<name count="n">...</name>` section: what it says, and what it holds. */
function section(xml: string, name: string, child: string): { declared: number; found: number } {
    const found = new RegExp(`<${name} count="(\\d+)">(.*?)</${name}>`).exec(xml);
    assert.ok(found, `no <${name}> section`);
    const inner = found[2] ?? '';
    return {
        declared: Number(found[1]),
        // Opening tags of `child` only: the lookahead is what keeps <fonts>
        // from counting as a <font> and </font> from counting at all.
        found: (inner.match(new RegExp(`<${child}(?=[ />])`, 'g')) ?? []).length,
    };
}

/** The `<xf>` entries of `<cellXfs>`, in the order they were handed out. */
function cellXfs(xml: string): string[] {
    const inner = /<cellXfs count="\d+">(.*)<\/cellXfs>/.exec(xml)?.[1] ?? '';
    // The empty form and the one with an <alignment> inside, as two patterns:
    // one `[^>]*` covering both would run past the end of the first entry.
    return inner.match(/<xf [^>]*\/>|<xf [^>]*>.*?<\/xf>/g) ?? [];
}

describe('argb', () => {
    it('takes the four spellings of a colour, with or without the hash', () => {
        assert.equal(argb('#FF0000'), 'FFFF0000');
        assert.equal(argb('FF0000'), 'FFFF0000');
        assert.equal(argb('80FF0000'), '80FF0000');
        assert.equal(argb('#f00'), 'FFFF0000');
    });

    it('upper-cases what it is given', () => {
        assert.equal(argb('#ffe699'), 'FFFFE699');
    });

    it('refuses what is not a colour rather than writing it out', () => {
        // An attribute Excel chokes on is a file that will not open, and by
        // then the rows are already gone.
        assert.throws(() => argb('red'), /not a colour/);
        assert.throws(() => argb('#12345'), /not a colour/);
        assert.throws(() => argb('#GGGGGG'), /not a colour/);
        assert.throws(() => argb(''), /not a colour/);
    });
});

describe('StyleTable: the default style', () => {
    it('is index 0, which is what an unstyled cell asks for', () => {
        const styles = new StyleTable();
        assert.equal(styles.index(undefined), 0);
        assert.equal(styles.index({}), 0);
    });

    it('is the only entry of a table nobody asked anything of', () => {
        assert.deepEqual(cellXfs(new StyleTable().xml()), [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        ]);
    });
});

describe('StyleTable: registering', () => {
    it('hands out one index per distinct style, in the order they arrive', () => {
        const styles = new StyleTable();
        assert.equal(styles.index({ bold: true }), 1);
        assert.equal(styles.index({ italic: true }), 2);
        assert.equal(cellXfs(styles.xml()).length, 3);
    });

    it('gives the same index to the same style, however it was spelled', () => {
        const styles = new StyleTable({ head: { bold: true } });
        assert.equal(styles.index({ bold: true }), styles.index('head'));
        // Two ways of writing one colour are one fill.
        assert.equal(styles.index({ bg: '#f00' }), styles.index({ bg: 'FFFF0000' }));
        assert.equal(cellXfs(styles.xml()).length, 3); // the default, the bold, the red
    });

    it('shares the fonts, fills and borders between the styles that repeat them', () => {
        const styles = new StyleTable();
        styles.index({ bold: true, bg: '#f00' });
        styles.index({ bold: true, bg: '#0f0' });
        styles.index({ italic: true, bg: '#f00' });
        const xml = styles.xml();
        // Two fonts past the default, two fills past the two Excel reserves,
        // and four entries in the table that combines them.
        assert.equal(section(xml, 'fonts', 'font').declared, 3);
        assert.equal(section(xml, 'fills', 'fill').declared, 4);
        assert.equal(cellXfs(xml).length, 4);
    });

    it('declares as many entries as it writes', () => {
        const styles = new StyleTable();
        styles.index({ bold: true, bg: '#f00', border: { all: 'thin' }, numFmt: '0.00' });
        const xml = styles.xml();
        for (const [name, child] of [
            ['numFmts', 'numFmt'],
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

    it('is a standalone styleSheet part', () => {
        const xml = new StyleTable().xml();
        assert.ok(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'));
        assert.ok(xml.endsWith('</styleSheet>'));
    });

    it('leaves out the section it has nothing for', () => {
        // An empty `<numFmts/>` is not a workbook with no formats of its own,
        // it is a workbook Excel refuses to open.
        assert.ok(!new StyleTable().xml().includes('numFmts'));
    });
});

describe('StyleTable: declared styles', () => {
    it('is asked for by name, and registered once', () => {
        const styles = new StyleTable({ money: { numFmt: '#,##0.00' } });
        assert.equal(styles.index('money'), 1);
        assert.equal(styles.index('money'), 1);
        assert.equal(cellXfs(styles.xml()).length, 2);
    });

    it('says what a name nobody declared was, rather than going in unstyled', () => {
        const styles = new StyleTable({ money: {} });
        assert.throws(() => styles.index('monye'), /Unknown style "monye"[\s\S]*money/);
        assert.throws(() => new StyleTable().index('x'), /None are declared/);
    });

    it('is what `base` starts from, under what the style itself says', () => {
        const styles = new StyleTable({ money: { numFmt: '#,##0.00', align: 'right' } });
        assert.deepEqual(styles.spec({ base: 'money', bold: true }), {
            numFmt: '#,##0.00',
            align: 'right',
            bold: true,
        });
    });

    it('lets what is based on it override, field by field', () => {
        const styles = new StyleTable({ money: { numFmt: '#,##0.00', bold: true } });
        assert.deepEqual(styles.spec({ base: 'money', bold: false }), {
            numFmt: '#,##0.00',
            bold: false,
        });
    });

    it('follows a base through as many styles as it takes', () => {
        const styles = new StyleTable({
            a: { size: 8 },
            b: { base: 'a', bold: true },
            c: { base: 'b', italic: true },
        });
        assert.deepEqual(styles.spec('c'), { size: 8, bold: true, italic: true });
    });

    it('refuses a base that leads back to where it started', () => {
        const styles = new StyleTable({ a: { base: 'b' }, b: { base: 'a' } });
        assert.throws(() => styles.index('a'), /based on itself/);
    });
});

describe('StyleTable: what one style renders as', () => {
    /** The whole part, for a table one style was registered in. */
    function xmlOf(spec: StyleSpec): string {
        const styles = new StyleTable();
        styles.index(spec);
        return styles.xml();
    }

    it('carries the font attributes, and says it means them', () => {
        const xml = xmlOf({ bold: true, italic: true, strike: true, underline: true });
        assert.ok(xml.includes('<font><b/><i/><strike/><u/><sz val="11"/>'), xml);
        assert.ok(cellXfs(xml)[1]?.includes('applyFont="1"'));
    });

    it('names the underline that is not the plain one', () => {
        assert.ok(xmlOf({ underline: 'double' }).includes('<u val="double"/>'));
        assert.ok(!xmlOf({ underline: false }).includes('<u'));
    });

    it('spells out the superscript and the subscript', () => {
        assert.ok(xmlOf({ script: 'super' }).includes('<vertAlign val="superscript"/>'));
        assert.ok(xmlOf({ script: 'sub' }).includes('<vertAlign val="subscript"/>'));
    });

    it('drops the theme scheme once a font of its own is asked for', () => {
        // `minor` is the theme's body font, which is only Calibri while
        // nobody has named another one.
        assert.ok(xmlOf({ font: 'Arial', size: 14 }).includes('<sz val="14"/><color theme="1"/><name val="Arial"/><family val="2"/></font>'));
        // Entry 0 is the sheet's own default font, which keeps its scheme.
        assert.equal((xmlOf({ font: 'Arial' }).match(/scheme/g) ?? []).length, 1);
        assert.ok(xmlOf({ bold: true }).includes('<b/><sz val="11"/>'));
    });

    it('writes the text colour, and keeps the theme one without it', () => {
        assert.ok(xmlOf({ color: '#003366' }).includes('<color rgb="FF003366"/>'));
        assert.ok(xmlOf({ bold: true }).includes('<color theme="1"/>'));
    });

    it('writes a background as a solid fill, past the two Excel reserves', () => {
        const xml = xmlOf({ bg: '#FFE699' });
        assert.ok(
            xml.includes(
                '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor indexed="64"/></patternFill></fill>',
            ),
            xml,
        );
        assert.ok(cellXfs(xml)[1]?.includes('fillId="2"'), 'the reserved fills were written over');
    });

    it('registers a number format past the ids Excel keeps for its own', () => {
        const xml = xmlOf({ numFmt: 'dd/mm/yyyy' });
        assert.ok(xml.includes('<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>'), xml);
        assert.ok(cellXfs(xml)[1]?.includes('numFmtId="164"'));
        assert.ok(cellXfs(xml)[1]?.includes('applyNumberFormat="1"'));
    });

    it("takes a number as one of Excel's own formats, and writes none", () => {
        const xml = xmlOf({ numFmt: 14 });
        assert.ok(!xml.includes('numFmts'), xml);
        assert.ok(cellXfs(xml)[1]?.includes('numFmtId="14"'));
    });

    it('escapes a format code that has XML in it', () => {
        assert.ok(xmlOf({ numFmt: '[<100]0.0;0' }).includes('formatCode="[&lt;100]0.0;0"'));
    });

    it('draws every side a border asks for', () => {
        const xml = xmlOf({ border: { all: 'thin', bottom: { style: 'thick', color: '#f00' } } });
        assert.ok(xml.includes('<left style="thin"/>'), xml);
        assert.ok(xml.includes('<bottom style="thick"><color rgb="FFFF0000"/></bottom>'), xml);
        // A border around a cell is not a border across it.
        assert.ok(xml.includes('<diagonal/></border>'), xml);
        assert.ok(cellXfs(xml)[1]?.includes('applyBorder="1"'));
    });

    it('marks the diagonal it was asked to run', () => {
        assert.ok(xmlOf({ border: { diagonal: 'thin', diagonalUp: true } }).includes('<border diagonalUp="1">'));
        assert.ok(xmlOf({ border: { diagonal: 'thin', diagonalDown: true } }).includes('<border diagonalDown="1">'));
        assert.ok(xmlOf({ border: { diagonal: 'thin' } }).includes('<diagonal style="thin"/>'));
    });

    it('draws no side for one that names no line', () => {
        assert.equal(xmlOf({ border: { top: {} } }), xmlOf({ border: {} }));
    });

    it('puts the alignment inside the entry, not in a table of its own', () => {
        const xml = xmlOf({ align: 'center', valign: 'middle', wrap: true, indent: 2 });
        assert.ok(
            cellXfs(xml)[1]?.includes(
                '<alignment horizontal="center" vertical="center" wrapText="1" indent="2"/>',
            ),
            xml,
        );
        assert.ok(cellXfs(xml)[1]?.includes('applyAlignment="1"'));
        assert.ok(xmlOf({ shrink: true }).includes('<alignment shrinkToFit="1"/>'));
    });

    it('counts a rotation the way Excel does', () => {
        // 0-90 turns counterclockwise; past 90 it turns the other way, so -45
        // is 135, and the vertical layout is a value of its own.
        assert.ok(xmlOf({ rotate: 45 }).includes('textRotation="45"'));
        assert.ok(xmlOf({ rotate: -45 }).includes('textRotation="135"'));
        assert.ok(xmlOf({ rotate: 255 }).includes('textRotation="255"'));
        assert.throws(() => xmlOf({ rotate: 120 }), /not between -90 and 90/);
    });

    it('writes the protection only when it is not what Excel already does', () => {
        assert.ok(cellXfs(xmlOf({ locked: false }))[1]?.includes('<protection locked="0"/>'));
        assert.ok(cellXfs(xmlOf({ locked: true }))[1]?.includes('<protection locked="1"/>'));
        assert.ok(cellXfs(xmlOf({ hideFormula: true }))[1]?.includes('<protection hidden="1"/>'));
        assert.ok(!xmlOf({ bold: true }).includes('protection'));
    });
});

describe('StyleTable: what a date is shown as', () => {
    const day = new Date(2024, 0, 15);
    const moment = new Date(2024, 0, 15, 12, 30);

    function table(): StyleTable {
        return new StyleTable({ money: { numFmt: '#,##0.00' }, big: { size: 20 } });
    }

    it('leaves anything that is not a date to the style it was given', () => {
        const styles = table();
        assert.equal(styles.forValue(1, undefined), 0);
        assert.equal(styles.forValue('x', 'money'), styles.index('money'));
    });

    it('formats a date that asked for nothing, so it is not shown as a serial', () => {
        const styles = table();
        assert.notEqual(styles.forValue(day, undefined), 0);
        assert.ok(styles.xml().includes(`formatCode="${DATE_FORMAT}"`));
    });

    it('adds the time only when the date carries one', () => {
        const styles = table();
        assert.notEqual(styles.forValue(moment, undefined), styles.forValue(day, undefined));
        assert.ok(styles.xml().includes(`formatCode="${DATETIME_FORMAT}"`));
    });

    it('keeps the rest of the style it was given, and only adds the format', () => {
        const styles = table();
        const plain = styles.index('big');
        const dated = styles.forValue(day, 'big');
        assert.notEqual(dated, plain);
        assert.deepEqual(styles.spec('big'), { size: 20 }); // the style itself is untouched
        const entries = cellXfs(styles.xml());
        const fontId = (entry: string): string => /fontId="(\d+)"/.exec(entry)?.[1] ?? '';
        assert.equal(fontId(entries[dated] ?? ''), fontId(entries[plain] ?? ''));
        assert.ok(entries[dated]?.includes('applyNumberFormat="1"'), entries[dated] ?? 'no entry');
    });

    it('leaves a style that already says how to show its number alone', () => {
        // Asking for a format is how a caller says it wants that one.
        const styles = table();
        assert.equal(styles.forValue(day, 'money'), styles.index('money'));
    });

    it('asks for the date style once, however many dates go through it', () => {
        const styles = new StyleTable();
        for (let k = 0; k < 100; k++) styles.forValue(new Date(2024, 0, 1 + k), undefined);
        assert.equal(cellXfs(styles.xml()).length, 2); // the default, and the date
    });
});
