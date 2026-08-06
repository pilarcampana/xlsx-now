import assert from 'node:assert/strict';
import { columnsMode } from '../src/core/columns.js';
import type { Column } from '../src/core/types.js';
import { ValueTypes } from '../src/core/valueTypes.js';

/** The types a workbook that said nothing about them knows. */
const TYPES = new ValueTypes();

/** The fill a pk column asks for. It is `columns.ts`'s to name, not a caller's. */
const PK_FILL = '#FFE699';

const ID: Column = { name: 'id', pk: true };
const YEAR: Column = { name: 'year', pk: true };
const NAME: Column = { name: 'name' };

describe('columnsMode: the freeze the columns imply', () => {
    it('fixes the header row and nothing else without pks', () => {
        assert.deepEqual(columnsMode([NAME], TYPES).freeze, { rows: 1, columns: 0 });
    });

    it('fixes the leading pk columns along with it', () => {
        assert.deepEqual(columnsMode([ID, YEAR, NAME], TYPES).freeze, { rows: 1, columns: 2 });
    });

    it('fixes no column when a pk is not one of the first', () => {
        // A freeze is a split at one position: freezing the pk here would
        // drag `name` along with it.
        assert.deepEqual(columnsMode([NAME, ID], TYPES).freeze, { rows: 1, columns: 0 });
        assert.deepEqual(columnsMode([ID, NAME, YEAR], TYPES).freeze, { rows: 1, columns: 1 });
    });

    it('fixes no column when every column is a pk', () => {
        // Freezing all of them would leave nothing to scroll.
        assert.deepEqual(columnsMode([ID, YEAR], TYPES).freeze, { rows: 1, columns: 0 });
    });

    it('fixes the header row of a sheet with no columns at all', () => {
        assert.deepEqual(columnsMode([], TYPES).freeze, { rows: 1, columns: 0 });
    });
});

describe('columnsMode: the header row', () => {
    it('is the column names, bold, and highlighted where they are pks', () => {
        assert.deepEqual(columnsMode([ID, NAME], TYPES).headerRow, [
            { v: 'id', s: { bold: true, bg: PK_FILL } },
            { v: 'name', s: { bold: true } },
        ]);
    });

    it('shows the name, not the key it reads', () => {
        const header = columnsMode([{ name: 'Full name', key: 'full_name' }], TYPES).headerRow;
        assert.deepEqual(header, [{ v: 'Full name', s: { bold: true } }]);
    });
});

describe('columnsMode: one record as one row', () => {
    it('reads each column by name, in the declared order', () => {
        const { toCellRow } = columnsMode([NAME, { name: 'age' }], TYPES);
        assert.deepEqual(toCellRow({ age: 30, name: 'Ana' }), ['Ana', 30]);
    });

    it('reads by key when the column has one', () => {
        const { toCellRow } = columnsMode([{ name: 'Full name', key: 'full_name' }], TYPES);
        assert.deepEqual(toCellRow({ full_name: 'Ana', 'Full name': 'ignored' }), ['Ana']);
    });

    it('highlights the pk columns', () => {
        const { toCellRow } = columnsMode([ID, NAME], TYPES);
        assert.deepEqual(toCellRow({ id: 1, name: 'Ana' }), [{ v: 1, s: { bg: PK_FILL } }, 'Ana']);
    });

    it('leaves a missing property empty rather than failing', () => {
        const { toCellRow } = columnsMode([ID, NAME], TYPES);
        assert.deepEqual(toCellRow({}), [{ v: undefined, s: { bg: PK_FILL } }, undefined]);
    });

    it('leaves a cell that says how it looks saying it', () => {
        // The pk fill is what the column asks for, not what it insists on.
        const { toCellRow } = columnsMode([ID], TYPES);
        assert.deepEqual(toCellRow({ id: { v: 1, s: 'flagged' } }), [{ v: 1, s: 'flagged' }]);
    });

    it('fills a cell that says everything but how it looks', () => {
        const { toCellRow } = columnsMode([ID], TYPES);
        assert.deepEqual(toCellRow({ id: { v: 1, f: 'A1' } }), [
            { v: 1, f: 'A1', s: { bg: PK_FILL } },
        ]);
    });

    it('keeps every kind of value as it is', () => {
        const date = new Date('2024-01-15T12:30:00.000Z');
        const { toCellRow } = columnsMode([NAME, { name: 'when' }, { name: 'ok' }], TYPES);
        assert.deepEqual(toCellRow({ name: null, when: date, ok: false }), [null, date, false]);
    });
});
