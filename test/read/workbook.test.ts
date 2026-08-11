import assert from 'node:assert/strict';
import {
    PACKAGE_ROOT,
    partOfType,
    readRelationships,
    readWorkbook,
    relsFor,
    resolvePart,
    workbookPart,
} from '../../src/core/read/workbook.js';

const RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(...entries: readonly string[]): string {
    return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join('')}</Relationships>`;
}

function rel(id: string, type: string, target: string, mode = ''): string {
    return `<Relationship Id="${id}" Type="${RELATIONSHIP_TYPE}/${type}" Target="${target}"${mode}/>`;
}

describe('relsFor', () => {
    it('is where a part keeps its relationships', () => {
        assert.equal(relsFor('xl/workbook.xml'), 'xl/_rels/workbook.xml.rels');
        assert.equal(relsFor(PACKAGE_ROOT), '_rels/.rels');
    });
});

describe('resolvePart', () => {
    it('reads a target as relative to the part it belongs to', () => {
        // Not to the `_rels/` directory the relationship is written in, which
        // would put every part of the package one directory too deep.
        assert.equal(resolvePart('xl/workbook.xml', 'worksheets/sheet1.xml'), 'xl/worksheets/sheet1.xml');
        assert.equal(resolvePart(PACKAGE_ROOT, 'xl/workbook.xml'), 'xl/workbook.xml');
    });

    it('takes a target that starts at the root for what it says', () => {
        assert.equal(resolvePart('xl/workbook.xml', '/xl/styles.xml'), 'xl/styles.xml');
    });

    it('lets a target climb out of its own directory', () => {
        assert.equal(resolvePart('xl/workbook.xml', '../docProps/app.xml'), 'docProps/app.xml');
        assert.equal(resolvePart('xl/workbook.xml', './styles.xml'), 'xl/styles.xml');
    });
});

describe('readRelationships', () => {
    it('is every relationship of a part, by id', () => {
        const map = readRelationships(
            rels(rel('rId1', 'worksheet', 'worksheets/sheet1.xml'), rel('rId2', 'styles', 'styles.xml')),
            'xl/workbook.xml',
        );
        assert.deepEqual(map.get('rId1'), { type: 'worksheet', part: 'xl/worksheets/sheet1.xml' });
        assert.deepEqual(map.get('rId2'), { type: 'styles', part: 'xl/styles.xml' });
    });

    it('leaves out a target that is not in the package at all', () => {
        const map = readRelationships(
            rels(rel('rId1', 'hyperlink', 'https://example.com', ' TargetMode="External"')),
            'xl/workbook.xml',
        );
        assert.equal(map.size, 0);
    });
});

describe('partOfType and workbookPart', () => {
    it('finds the singular parts by what they are', () => {
        const map = readRelationships(
            rels(rel('rId1', 'worksheet', 'worksheets/sheet1.xml'), rel('rId2', 'sharedStrings', 'sharedStrings.xml')),
            'xl/workbook.xml',
        );
        assert.equal(partOfType(map, 'sharedStrings'), 'xl/sharedStrings.xml');
        assert.equal(partOfType(map, 'styles'), undefined);
    });

    it('follows the package to wherever the workbook is', () => {
        assert.equal(workbookPart(rels(rel('rId1', 'officeDocument', 'xl/workbook.xml'))), 'xl/workbook.xml');
        // A package is not obliged to call it that, and some do not.
        assert.equal(workbookPart(rels(rel('rId1', 'officeDocument', 'libro/hoja.xml'))), 'libro/hoja.xml');
    });

    it('refuses a package with no workbook in it', () => {
        assert.throws(() => workbookPart(rels(rel('rId1', 'styles', 'styles.xml'))), /no workbook/);
    });
});

describe('readWorkbook', () => {
    it('is the sheets in the order the workbook declares them', () => {
        const workbook = readWorkbook(
            '<workbook xmlns:r="x"><sheets>' +
                '<sheet name="Segunda" sheetId="7" r:id="rIdB"/>' +
                '<sheet name="Primera" sheetId="2" r:id="rIdA"/>' +
                '</sheets></workbook>',
            'xl/workbook.xml',
        );
        assert.deepEqual(workbook.sheets, [
            { name: 'Segunda', relationshipId: 'rIdB' },
            { name: 'Primera', relationshipId: 'rIdA' },
        ]);
    });

    it('counts from 1900 unless the workbook says otherwise', () => {
        const of = (pr: string): boolean =>
            readWorkbook(`<workbook>${pr}<sheets/></workbook>`, 'xl/workbook.xml').date1904;
        assert.equal(of(''), false);
        assert.equal(of('<workbookPr/>'), false);
        assert.equal(of('<workbookPr date1904="1"/>'), true);
        assert.equal(of('<workbookPr date1904="true"/>'), true);
        assert.equal(of('<workbookPr date1904="0"/>'), false);
    });

    it('refuses a sheet that points at no part', () => {
        assert.throws(
            () => readWorkbook('<workbook><sheets><sheet name="H"/></sheets></workbook>', 'xl/workbook.xml'),
            /no relationship/,
        );
    });
});
