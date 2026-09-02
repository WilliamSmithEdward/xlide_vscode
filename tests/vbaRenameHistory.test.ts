import { beforeEach, describe, expect, it } from 'vitest';
import {
	describeUndo,
	invalidateRenameHistory,
	noteModuleWrite,
	pendingUndo,
	recordRename,
	takeRenameForUndo,
	type RenameSnapshot,
} from '../src/vbaRenameHistory';

// Issue #9 rule 10. A rename edits several modules and an editor's undo stack
// is per document, so undoing in the file you are looking at would reverse that
// file's share and leave the rest renamed.
function snapshot(overrides: Partial<RenameSnapshot> = {}): RenameSnapshot {
	return {
		projectPath: 'C:/w/Book.xlsm',
		oldName: 'Alpha',
		newName: 'Beta',
		modules: [
			{ moduleName: 'Helpers', before: 'Public Sub Alpha()\r\nEnd Sub\r\n' },
			{ moduleName: 'Consumer', before: 'Public Sub Drive()\r\n    Alpha\r\nEnd Sub\r\n' },
		],
		...overrides,
	};
}

beforeEach(() => {
	invalidateRenameHistory();
});

describe('the rename that can be put back', () => {
	it('remembers what every touched module said before', () => {
		recordRename(snapshot());
		expect(pendingUndo()?.modules.map((m) => m.moduleName)).toEqual(['Helpers', 'Consumer']);
		expect(pendingUndo()?.modules[0].before).toContain('Public Sub Alpha()');
	});

	it('can only be taken once', () => {
		// A second restore would write stale text over whatever the first
		// produced.
		recordRename(snapshot());
		expect(takeRenameForUndo()).toBeDefined();
		expect(takeRenameForUndo()).toBeUndefined();
		expect(pendingUndo()).toBeUndefined();
	});

	it('keeps only the most recent rename', () => {
		// A stack would invite undoing a rename from three renames ago, by which
		// time its before-images describe modules that have moved on.
		recordRename(snapshot());
		recordRename(snapshot({ oldName: 'Gamma', newName: 'Delta' }));
		expect(pendingUndo()?.oldName).toBe('Gamma');
	});

	it('records nothing when the rename touched nothing', () => {
		recordRename(snapshot({ modules: [] }));
		expect(pendingUndo()).toBeUndefined();
	});

	it('still records a module rename that edited no text', () => {
		recordRename(snapshot({ modules: [], renamedModule: { from: 'Old', to: 'New' } }));
		expect(pendingUndo()?.renamedModule).toEqual({ from: 'Old', to: 'New' });
	});
});

describe('the history is dropped when it would be wrong to use', () => {
	it('is dropped for the project that was written to', () => {
		recordRename(snapshot());
		invalidateRenameHistory('c:/W/BOOK.XLSM');
		expect(pendingUndo()).toBeUndefined();
	});

	it('survives a write to a different project', () => {
		recordRename(snapshot());
		invalidateRenameHistory('C:/w/Other.xlsm');
		expect(pendingUndo()).toBeDefined();
	});

	it('is dropped entirely when no project is named', () => {
		recordRename(snapshot());
		invalidateRenameHistory();
		expect(pendingUndo()).toBeUndefined();
	});
});

describe('what the undo says it will do', () => {
	it('names the rename and how far it reached', () => {
		expect(describeUndo(snapshot())).toBe(
			"Undo renaming 'Alpha' to 'Beta' across 2 modules.",
		);
	});

	it('mentions the module rename when there was one', () => {
		const text = describeUndo(snapshot({ renamedModule: { from: 'Old', to: 'New' } }));
		expect(text).toContain('rename the module back to Old');
	});

	it('uses the singular for one module', () => {
		const text = describeUndo(snapshot({ modules: [{ moduleName: 'M', before: 'x' }] }));
		expect(text).toContain('across 1 module.');
	});
});

describe('telling the rename own writes from everyone else', () => {
    // Applying a rename's edits and a developer pressing Save arrive at the
    // same file-system path. Without the distinction the rename would either
    // clear its own history and never offer an undo, or an undo would restore
    // over a save the developer made afterwards.
    it('survives the writes the rename itself causes', () => {
        recordRename(snapshot());
        noteModuleWrite('C:/w/Book.xlsm', 'Helpers');
        noteModuleWrite('C:/w/Book.xlsm', 'Consumer');
        expect(pendingUndo()).toBeDefined();
    });

    it('is dropped by a save to a module it touched', () => {
        recordRename(snapshot());
        noteModuleWrite('C:/w/Book.xlsm', 'Helpers');   // the rename's own
        noteModuleWrite('C:/w/Book.xlsm', 'Helpers');   // then the developer saves
        expect(pendingUndo()).toBeUndefined();
    });

    it('is dropped by a save to a module it never touched', () => {
        recordRename(snapshot());
        noteModuleWrite('C:/w/Book.xlsm', 'Unrelated');
        expect(pendingUndo()).toBeUndefined();
    });

    it('ignores a write to a different project', () => {
        recordRename(snapshot());
        noteModuleWrite('C:/w/Other.xlsm', 'Helpers');
        expect(pendingUndo()).toBeDefined();
    });

    it('matches module and project case-insensitively', () => {
        recordRename(snapshot());
        noteModuleWrite('c:/W/BOOK.XLSM', 'helpers');
        noteModuleWrite('c:/W/BOOK.XLSM', 'CONSUMER');
        expect(pendingUndo()).toBeDefined();
    });
});
