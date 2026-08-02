import assert from 'node:assert/strict';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from '../src/core/parts.js';

const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

describe('contentTypesXml', () => {
    const xml = contentTypesXml();

    it('declares a type for every part the writer emits', () => {
        assert.ok(xml.includes('PartName="/xl/workbook.xml"'));
        assert.ok(xml.includes('PartName="/xl/styles.xml"'));
        assert.ok(xml.includes('PartName="/xl/worksheets/sheet1.xml"'));
    });

    it('defaults the two extensions the .rels parts need', () => {
        assert.ok(xml.includes('Extension="rels"'));
        assert.ok(xml.includes('Extension="xml"'));
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

describe('workbookXml', () => {
    it('declares one sheet, under the name it was given', () => {
        assert.ok(workbookXml('Sheet1').includes('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'));
    });

    it('escapes the sheet name', () => {
        assert.ok(workbookXml('Q1 & Q2 <"draft">').includes('name="Q1 &amp; Q2 &lt;&quot;draft&quot;&gt;"'));
    });

    it('is a standalone part', () => {
        const xml = workbookXml('Sheet1');
        assert.ok(xml.startsWith(PROLOG));
        assert.ok(xml.endsWith('</workbook>'));
    });
});

describe('workbookRelsXml', () => {
    const xml = workbookRelsXml();

    it('relates the workbook to its worksheet and its styles', () => {
        assert.match(xml, /Id="rId1"[^>]*Target="worksheets\/sheet1\.xml"/);
        assert.match(xml, /Id="rId2"[^>]*Target="styles\.xml"/);
    });

    it('uses for the worksheet the same id workbook.xml refers to', () => {
        assert.ok(workbookXml('Sheet1').includes('r:id="rId1"'));
        assert.match(xml, /Id="rId1"[^>]*relationships\/worksheet"/);
    });
});
