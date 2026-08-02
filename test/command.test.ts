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

describe('lineCells', () => {
    it('takes an empty line as a row with nothing in it', () => {
        assert.deepEqual(lineCells({ '#line': 'empty' }), []);
    });

    it('takes an array as the row it already is', () => {
        assert.deepEqual(lineCells({ '#line': 'array', values: [1, 'a'] }), [1, 'a']);
    });

    it('leaves a record to the caller, who has the columns', () => {
        assert.equal(lineCells({ '#line': 'row', values: { id: 1 } }), undefined);
    });

    it('takes a missing `values` as an empty line rather than a failure', () => {
        assert.deepEqual(lineCells({ '#line': 'array' } as unknown as LineCommand), []);
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
