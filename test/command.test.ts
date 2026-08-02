import assert from 'node:assert/strict';
import {
    LINE,
    WORKSHEET,
    checkRecord,
    isLineCommand,
    isWorksheetCommand,
    lineCells,
    lineRecord,
    noColumnsError,
    sparseCellRow,
    type LineCommand,
} from '../src/core/command.js';

describe('the command keys', () => {
    it('are the two a message can carry', () => {
        assert.equal(WORKSHEET, '#worksheet');
        assert.equal(LINE, '#line');
    });
});

describe('isWorksheetCommand', () => {
    it('recognizes a command by its key, whatever else it carries', () => {
        assert.ok(isWorksheetCommand({ '#worksheet': 'Sheet2' }));
        assert.ok(isWorksheetCommand({ '#worksheet': 'Sheet2', freezeColumns: 2, columns: [] }));
    });

    it('takes a record for what it is', () => {
        assert.ok(!isWorksheetCommand({ id: 1, name: 'Ana' }));
        // The key is the whole of the difference: a column named after it
        // would be read as a command, which is why `#` is reserved.
        assert.ok(!isWorksheetCommand({ '#': 1 }));
    });

    it('takes a row of cells for what it is', () => {
        assert.ok(!isWorksheetCommand(['#worksheet', 'Sheet2']));
        assert.ok(!isWorksheetCommand([]));
    });
});

describe('isLineCommand', () => {
    it('recognizes a line by its key', () => {
        assert.ok(isLineCommand({ '#line': 'empty' }));
        assert.ok(isLineCommand({ '#line': 'array', values: [1] }));
        assert.ok(!isLineCommand({ line: 'empty' }));
        assert.ok(!isLineCommand([1, 2]));
    });
});

describe('checkRecord', () => {
    it('lets an ordinary record through', () => {
        assert.doesNotThrow(() => checkRecord({ id: 1, '#tag': undefined }.id ? {} : { id: 1 }));
        assert.doesNotThrow(() => checkRecord({ id: 1, name: 'Ana' }));
    });

    it('names the command nobody knows, rather than writing a row of blanks', () => {
        // A misspelled command is a record whose key starts with `#`.
        assert.throws(() => checkRecord({ '#worksheets': 'Sheet2' }), /Unknown command/);
        assert.throws(() => checkRecord({ '#lines': 'empty' }), /"#line"/);
    });
});

describe('noColumnsError', () => {
    it('says what a record without columns is missing', () => {
        assert.match(noColumnsError().message, /needs columns/);
    });
});

describe('sparseCellRow', () => {
    it('puts each value at the column its letter names', () => {
        const row = sparseCellRow({ A: 'first', C: 3 });
        assert.equal(row.length, 3);
        assert.equal(row[0], 'first');
        assert.equal(row[1], undefined, 'the gap was filled in');
        assert.equal(row[2], 3);
    });

    it('reads the letters past Z, and reads them in any case', () => {
        assert.equal(sparseCellRow({ AA: 1 }).length, 27);
        assert.equal(sparseCellRow({ aa: 1 }).length, 27);
    });

    it('leaves the gaps as gaps, so a far column costs one cell', () => {
        const row = sparseCellRow({ BZ: 'far' });
        assert.equal(row.length, 78);
        // A hole is not a value: nothing is written for it.
        assert.equal(Object.keys(row).length, 1);
    });

    it('takes a styled cell like anywhere else', () => {
        assert.deepEqual(sparseCellRow({ B: { value: 'x', style: { bold: true } } })[1], {
            value: 'x',
            style: { bold: true },
        });
    });

    it('says what is not a column', () => {
        assert.throws(() => sparseCellRow({ '1': 'x' }), /not a column/);
        assert.throws(() => sparseCellRow({ 'A1': 'x' }), /column letters/);
        assert.throws(() => sparseCellRow({ '': 'x' }), /not a column/);
    });
});

describe('lineCells', () => {
    it('takes an empty line as a row with nothing in it', () => {
        assert.deepEqual(lineCells({ '#line': 'empty' }), []);
    });

    it('takes an array as the row it already is', () => {
        assert.deepEqual(lineCells({ '#line': 'array', values: [1, 'a'] }), [1, 'a']);
    });

    it('spreads a sparse line over its columns', () => {
        assert.equal(lineCells({ '#line': 'sparse', values: { B: 2 } })?.length, 2);
    });

    it('leaves a record to the caller, who has the columns', () => {
        assert.equal(lineCells({ '#line': 'row', values: { id: 1 } }), undefined);
    });

    it('takes a missing `values` as an empty line rather than a failure', () => {
        assert.deepEqual(lineCells({ '#line': 'array' } as unknown as LineCommand), []);
        assert.deepEqual(lineCells({ '#line': 'sparse' } as unknown as LineCommand), []);
    });

    it('says what a line nobody knows was asked to be', () => {
        assert.throws(
            () => lineCells({ '#line': 'colum' } as unknown as LineCommand),
            /Unknown line "colum"/,
        );
    });
});

describe('lineRecord', () => {
    it('is the record the command carries', () => {
        assert.deepEqual(lineRecord({ '#line': 'row', values: { id: 1 } }), { id: 1 });
    });

    it('is an empty record when the command carries none', () => {
        assert.deepEqual(lineRecord({ '#line': 'row' } as unknown as LineCommand), {});
    });

    it('is checked for commands like a bare record is', () => {
        assert.throws(
            () => lineRecord({ '#line': 'row', values: { '#worksheet': 'x' } }),
            /Unknown command/,
        );
    });
});
