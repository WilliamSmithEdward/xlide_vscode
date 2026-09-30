import { describe, expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';

// resolveBareIdentifier keeps the visible symbols of the last module that
// asked until the indexed modules change. These pin that the kept list
// follows the asking module and every edit, and that it is built once for a
// run of lookups from one module, as Find References makes.

const ALPHA = [
	'Private Hidden As Long',
	'Public Shared As Long',
	'Private Sub Internal()',
	'End Sub',
	'Public Sub Caller()',
	'    Hidden = Shared',
	'    Internal',
	'    Other',
	'End Sub',
	'',
].join('\r\n');

const BETA = [
	'Private Hidden As String',
	'Public Sub Other()',
	'    Hidden = "b"',
	'    Shared = 2',
	'    Internal',
	'End Sub',
	'',
].join('\r\n');

function project(): ProjectIndex {
	const index = new ProjectIndex();
	index.setModule({ moduleName: 'Alpha', moduleKind: 'standard', source: ALPHA });
	index.setModule({ moduleName: 'Beta', moduleKind: 'standard', source: BETA });
	return index;
}

function bindsTo(index: ProjectIndex, moduleName: string, source: string, word: string): string[] {
	const offset = source.indexOf(`    ${word}`) + 4;
	return index.resolveBareIdentifier(moduleName, word, offset, 'expression').definitions
		.map((symbol) => `${symbol.moduleName}.${symbol.name}`);
}

describe('ProjectIndex bare identifier resolution', () => {
	it('binds private names in their own module, whichever module asks in turn', () => {
		const index = project();
		for (let round = 0; round < 2; round += 1) {
			expect(bindsTo(index, 'Alpha', ALPHA, 'Hidden')).toEqual(['Alpha.Hidden']);
			expect(bindsTo(index, 'Beta', BETA, 'Hidden')).toEqual(['Beta.Hidden']);
			expect(bindsTo(index, 'Alpha', ALPHA, 'Internal')).toEqual(['Alpha.Internal']);
			// Alpha's kept list holds its private declarations; Beta must not see them.
			expect(bindsTo(index, 'Beta', BETA, 'Internal')).toEqual([]);
			expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual(['Alpha.Shared']);
			expect(bindsTo(index, 'Alpha', ALPHA, 'Other')).toEqual(['Beta.Other']);
		}
	});

	it('follows setModule and removeModule', () => {
		const index = project();
		expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual(['Alpha.Shared']);

		index.setModule({ moduleName: 'Alpha', moduleKind: 'standard', source: ALPHA.replace('Public Shared', 'Private Shared') });
		expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual([]);

		index.setModule({ moduleName: 'Gamma', moduleKind: 'standard', source: 'Public Shared As Long\r\n' });
		expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual(['Gamma.Shared']);

		index.removeModule('Gamma');
		expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual([]);
	});

	it('does not let a caller change a later answer', () => {
		const index = project();
		const offset = BETA.indexOf('    Shared') + 4;
		const first = index.resolveBareIdentifier('Beta', 'Shared', offset, 'expression');
		(first.definitions as unknown[]).length = 0;
		expect(bindsTo(index, 'Beta', BETA, 'Shared')).toEqual(['Alpha.Shared']);
	});

	it('builds the visible symbols once for a run of lookups from one module', () => {
		const index = project();
		const build = vi.spyOn(index, 'visibleIdentifierSymbols');
		for (const word of ['Hidden', 'Shared', 'Internal', 'Other']) {
			bindsTo(index, 'Alpha', ALPHA, word);
		}
		expect(build).toHaveBeenCalledTimes(1);

		bindsTo(index, 'Beta', BETA, 'Shared');
		bindsTo(index, 'Beta', BETA, 'Hidden');
		expect(build).toHaveBeenCalledTimes(2);

		index.setModule({ moduleName: 'Beta', moduleKind: 'standard', source: BETA });
		bindsTo(index, 'Beta', BETA, 'Shared');
		expect(build).toHaveBeenCalledTimes(3);
	});
});
