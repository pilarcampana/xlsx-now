import assert from 'node:assert/strict';
import {
    checkSheetName,
    contentTypesXml,
    rootRelsXml,
    workbookRelsXml,
    workbookXml,
    worksheetPart,
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

describe('checkSheetName', () => {
    it('takes a name Excel takes', () => {
        assert.doesNotThrow(() => checkSheetName('Ventas 2024', ['Compras']));
    });

    it('refuses an empty name, and anything that is not one', () => {
        assert.throws(() => checkSheetName('', []), /cannot be empty/);
        assert.throws(() => checkSheetName(undefined as unknown as string, []), /cannot be empty/);
    });

    it('refuses a name longer than the 31 characters Excel allows', () => {
        assert.doesNotThrow(() => checkSheetName('x'.repeat(31), []));
        assert.throws(() => checkSheetName('x'.repeat(32), []), /31 characters/);
    });

    it('refuses the characters Excel forbids', () => {
        for (const name of ['a/b', 'a\\b', 'a?b', 'a*b', 'a[b', 'a]b', 'a:b']) {
            assert.throws(() => checkSheetName(name, []), /forbids/, name);
        }
    });

    it('refuses a name another sheet already took, whatever its case', () => {
        assert.throws(() => checkSheetName('Ventas', ['Compras', 'ventas']), /already taken/);
    });
});
