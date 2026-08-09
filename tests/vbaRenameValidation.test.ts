import { describe, expect, it } from 'vitest';
import {
	checkRenameName,
	describeRenameCollision,
	findRenameCollision,
} from '../src/vbaRenameValidation';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';

// Rule 7 of issue #9: the new name is checked before anything is written. A
// rename edits several modules at once, so writing first and discovering the
// project no longer compiles is the worst outcome available.
function symbolsFor(source: string, moduleName = 'M') {
	const project = buildVbaProjectIndex([{ moduleName, source }]);
	return project.documentSymbols(moduleName);
}

describe('the new name must be a usable VBA identifier', () => {
	it.each(['Beta', 'ProcessOrder', '_leading', 'a1', 'Проверка', 'ค่าใหม่', 'नामनया'])(
		'accepts %s', (name) => {
			expect(checkRenameName(name)).toBeUndefined();
		});

	it.each(['1Bad', 'has space', 'has-dash', '', 'has.dot'])('rejects %s', (name) => {
		expect(checkRenameName(name)?.reason).toBe('not-an-identifier');
	});

	it.each(['Sub', 'End', 'If', 'Function', 'Dim', 'Next', 'sub', 'iF'])(
		'refuses the reserved word %s', (name) => {
			expect(checkRenameName(name)?.reason).toBe('reserved');
		});

	it('refuses a name longer than VBA allows', () => {
		expect(checkRenameName('A'.repeat(256))?.reason).toBe('too-long');
	});
});

describe('the new name must be free where the old one is declared', () => {
	it('refuses a rename that would declare a procedure twice', () => {
		// The issue's case: two Public Sub Beta() in one module compiles nowhere.
		const root = symbolsFor([
			'Option Explicit',
			'',
			'Public Sub Alpha()',
			'End Sub',
			'',
			'Public Sub Beta()',
			'End Sub',
			'',
		].join('\r\n'), 'Helpers');
		const collision = findRenameCollision(root, 'Helpers', 'Alpha', 'Beta');
		expect(collision).toBeDefined();
		expect(collision?.container).toBe('module Helpers');
		expect(describeRenameCollision(collision!, 'Beta')).toContain('does not compile');
	});

	it('allows a rename to a name nothing else declares', () => {
		const root = symbolsFor([
			'Option Explicit', '', 'Public Sub Alpha()', 'End Sub', '',
		].join('\r\n'), 'Helpers');
		expect(findRenameCollision(root, 'Helpers', 'Alpha', 'Gamma')).toBeUndefined();
	});

	it('allows a case-only change of the same name', () => {
		const root = symbolsFor('Public Sub alpha()\r\nEnd Sub\r\n', 'Helpers');
		expect(findRenameCollision(root, 'Helpers', 'alpha', 'Alpha')).toBeUndefined();
	});

	it('refuses a duplicate module-level variable', () => {
		const root = symbolsFor([
			'Option Explicit', '', 'Private mAlpha As Long', 'Private mBeta As Long', '',
		].join('\r\n'), 'Helpers');
		expect(findRenameCollision(root, 'Helpers', 'mAlpha', 'mBeta')).toBeDefined();
	});

	it('refuses a duplicate local inside one procedure', () => {
		const root = symbolsFor([
			'Option Explicit',
			'',
			'Public Sub Drive()',
			'    Dim first As Long',
			'    Dim second As Long',
			'End Sub',
			'',
		].join('\r\n'), 'Helpers');
		const collision = findRenameCollision(root, 'Helpers', 'first', 'second');
		expect(collision).toBeDefined();
		expect(collision?.container).toContain('Drive');
	});

	it('lets a local take a name the module also declares', () => {
		// VBA lets a local shadow a module-level name, so this is legal and must
		// not be refused - scope is the declaration's siblings, not the module.
		const root = symbolsFor([
			'Option Explicit',
			'',
			'Public Sub Shared()',
			'End Sub',
			'',
			'Public Sub Drive()',
			'    Dim temp As Long',
			'End Sub',
			'',
		].join('\r\n'), 'Helpers');
		expect(findRenameCollision(root, 'Helpers', 'temp', 'Shared')).toBeUndefined();
	});

	it('is silent when the module has no symbols at all', () => {
		expect(findRenameCollision(undefined, 'Helpers', 'Alpha', 'Beta')).toBeUndefined();
	});
});
