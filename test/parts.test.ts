import assert from 'node:assert/strict';
import {
    contentTypesXml,
    rootRelsXml,
    workbookRelsXml,
    workbookXml,
    worksheetPart,
    sheetName,
} from '../src/core/parts.js';

const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

describe('contentTypesXml', () => {
    const xml = contentTypesXml();

    it('overrides the type of the two parts that are not worksheets', () => {
        assert.ok(xml.includes('PartName="/xl/workbook.xml"'));
        assert.ok(xml.includes('PartName="/xl/styles.xml"'));
    });

    it('types every other .xml part as a worksheet, however many arrive', () => {
        // Naming them one by one would mean knowing them before the first
        // row, which is the one thing a streamed workbook cannot do.
        assert.ok(
            xml.includes(
                '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            ),
            xml,
        );
        assert.doesNotMatch(xml, /worksheets/);
    });

    it('defaults the extension the .rels parts need', () => {
        assert.ok(xml.includes('Extension="rels"'));
    });

    it('is a standalone part', () => {
        assert.ok(xml.startsWith(PROLOG));
        assert.ok(xml.endsWith('</Types>'));
    });
});

describe('rootRelsXml', () => {
    it('points the package at the workbook', () => {
        const xml = rootRelsXml();
        assert.ok(xml.startsWith(PROLOG));
        assert.match(xml, /Id="rId1"[^>]*Target="xl\/workbook\.xml"/);
    });
});

describe('worksheetPart', () => {
    it('numbers the parts from one, as Excel numbers the sheets', () => {
        assert.equal(worksheetPart(1), 'xl/worksheets/sheet1.xml');
        assert.equal(worksheetPart(12), 'xl/worksheets/sheet12.xml');
    });
});

describe('workbookXml', () => {
    it('declares one sheet, under the name it was given', () => {
        assert.ok(
            workbookXml(['Sheet1']).includes('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'),
        );
    });

    it('declares as many sheets as it was given, in order', () => {
        const xml = workbookXml(['Enero', 'Febrero', 'Marzo']);
        assert.ok(xml.includes('<sheet name="Enero" sheetId="1" r:id="rId1"/>'), xml);
        assert.ok(xml.includes('<sheet name="Febrero" sheetId="2" r:id="rId2"/>'), xml);
        assert.ok(xml.includes('<sheet name="Marzo" sheetId="3" r:id="rId3"/>'), xml);
    });

    it('escapes the sheet name', () => {
        assert.ok(
            workbookXml(['Q1 & Q2 <"draft">']).includes(
                'name="Q1 &amp; Q2 &lt;&quot;draft&quot;&gt;"',
            ),
        );
    });

    it('is a standalone part', () => {
        const xml = workbookXml(['Sheet1']);
        assert.ok(xml.startsWith(PROLOG));
        assert.ok(xml.endsWith('</workbook>'));
    });
});

describe('workbookRelsXml', () => {
    it('relates the workbook to its worksheet and its styles', () => {
        const xml = workbookRelsXml(1);
        assert.match(xml, /Id="rId1"[^>]*Target="worksheets\/sheet1\.xml"/);
        assert.match(xml, /Id="rId2"[^>]*Target="styles\.xml"/);
    });

    it('relates one worksheet per sheet, and the styles after the last of them', () => {
        const xml = workbookRelsXml(3);
        assert.match(xml, /Id="rId2"[^>]*Target="worksheets\/sheet2\.xml"/);
        assert.match(xml, /Id="rId3"[^>]*Target="worksheets\/sheet3\.xml"/);
        assert.match(xml, /Id="rId4"[^>]*Target="styles\.xml"/);
    });

    it('uses for every worksheet the same id workbook.xml refers to', () => {
        const xml = workbookRelsXml(2);
        for (const id of ['rId1', 'rId2']) {
            assert.ok(workbookXml(['a', 'b']).includes(`r:id="${id}"`));
            assert.match(xml, new RegExp(`Id="${id}"[^>]*relationships/worksheet"`));
        }
    });
});

describe('sheetName', () => {
    it('keeps a name Excel keeps', () => {
        assert.equal(sheetName('Ventas 2024', ['Compras'], 2), 'Ventas 2024');
    });

    it('names a sheet that arrived with nothing to be called', () => {
        assert.equal(sheetName(undefined, [], 1), 'Sheet1');
        assert.equal(sheetName('', ['Sheet1'], 2), 'Sheet2');
        assert.equal(sheetName(7 as unknown as string, [], 3), 'Sheet3');
    });

    it('drops the characters Excel forbids instead of refusing the name', () => {
        assert.equal(sheetName('Ventas/Compras', [], 1), 'VentasCompras');
        assert.equal(sheetName('a\\b?c*d[e]f:g', [], 1), 'abcdefg');
        assert.equal(sheetName("'draft'", [], 1), 'draft');
    });

    it('falls back to the default when nothing survives the cleaning', () => {
        assert.equal(sheetName('///', [], 4), 'Sheet4');
    });

    it('cuts a name longer than the 31 characters Excel allows', () => {
        assert.equal(sheetName('x'.repeat(40), [], 1), 'x'.repeat(31));
    });

    it('numbers a name another sheet already took, whatever its case', () => {
        assert.equal(sheetName('Ventas', ['ventas'], 2), 'Ventas (2)');
        assert.equal(sheetName('Ventas', ['Ventas', 'Ventas (2)'], 3), 'Ventas (3)');
    });

    it('makes room for the number inside the 31 characters', () => {
        const long = 'x'.repeat(31);
        const numbered = sheetName(long, [long], 2);
        assert.equal(numbered.length, 31);
        assert.equal(numbered, `${'x'.repeat(27)} (2)`);
    });
});
