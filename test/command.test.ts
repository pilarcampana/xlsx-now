import assert from 'node:assert/strict';
import { WORKSHEET, isWorksheetCommand, recordError } from '../src/core/command.js';

describe('WORKSHEET', () => {
    it('is the key a command carries', () => {
        assert.equal(WORKSHEET, '#worksheet');
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

describe('recordError', () => {
    it('says what a record without columns is missing', () => {
        assert.match(recordError({ id: 1 }).message, /needs columns/);
    });

    it('names the command nobody knows, rather than blaming the columns', () => {
        // A misspelled command is a record whose key starts with `#`, and
        // would otherwise go in as a blank row.
        assert.match(recordError({ '#worksheets': 'Sheet2' }).message, /Unknown command/);
        assert.match(recordError({ '#worksheets': 'Sheet2' }).message, /"#worksheet"/);
    });
});
