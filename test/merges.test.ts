import assert from 'node:assert/strict';
import { MergeTable } from '../src/core/merges.js';

/** The columns a row of the sheet finds already taken, in order. */
function columnsAt(merges: MergeTable, rowNumber: number): number[] {
    return merges.openAt(rowNumber).map((span) => span.column);
}

describe('MergeTable', () => {
    it('writes nothing at all for a sheet that merged nothing', () => {
        // An empty <mergeCells/> is not a sheet without merges: it is a sheet
        // Excel refuses to open.
        assert.equal(new MergeTable().xml(), '');
    });

    it('names a range by the cell that declared it and the one it reaches', () => {
        const merges = new MergeTable();
        assert.equal(merges.add(0, 1, 3, 1, 0), 'A1:C1');
        assert.equal(merges.add(0, 3, 1, 3, 0), 'A3:A5');
        assert.equal(merges.add(4, 7, 2, 2, 0), 'E7:F8');
        assert.equal(
            merges.xml(),
            '<mergeCells count="3"><mergeCell ref="A1:C1"/><mergeCell ref="A3:A5"/>' +
                '<mergeCell ref="E7:F8"/></mergeCells>',
        );
    });

    it('leaves the rows below alone when the merge stays inside its own row', () => {
        const merges = new MergeTable();
        merges.add(0, 1, 3, 1, 0);
        assert.deepEqual(columnsAt(merges, 2), []);
    });

    it('covers every column of the range in the rows below, and only those rows', () => {
        const merges = new MergeTable();
        merges.add(4, 7, 2, 2, 0);
        assert.deepEqual(columnsAt(merges, 8), [4, 5]);
        assert.deepEqual(columnsAt(merges, 9), []);
    });

    it('hands the columns over in order, whatever order the rows opened them in', () => {
        // Two merges started in different rows, the later one to the left:
        // what a row reads has to be ascending by column, since that is the
        // order its cells go out in.
        const merges = new MergeTable();
        merges.add(4, 1, 1, 5, 0);
        merges.add(1, 2, 1, 5, 0);
        assert.deepEqual(columnsAt(merges, 3), [1, 4]);
    });

    it('carries the style of the cell that declared the merge into the rows below', () => {
        const merges = new MergeTable();
        merges.add(0, 1, 1, 2, 7);
        assert.deepEqual(
            merges.openAt(2).map((span) => ({ style: span.style, ref: span.ref })),
            [{ style: 7, ref: 'A1:A2' }],
        );
    });

    it('says which merge is still waiting for a row nobody wrote', () => {
        const merges = new MergeTable();
        merges.add(0, 3, 1, 3, 0);
        // Asked with the row the sheet would have written next: after row 5
        // the range is closed, before it the sheet ends in the middle of one.
        assert.equal(merges.unfinishedAt(5)?.ref, 'A3:A5');
        assert.equal(merges.unfinishedAt(6), undefined);
    });
});
