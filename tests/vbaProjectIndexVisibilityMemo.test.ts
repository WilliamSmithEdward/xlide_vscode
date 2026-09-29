import { describe, expect, it } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';

// Each module's part of a visibility query is memoized twice: what it shows
// its own queries (private declarations included) and what it shows every
// other module's. These pin that the two never mix, whichever module asks
// first, and that setModule drops both.

const ALPHA = [
	'Public Const Shared As Long = 1',
	'Private Const Hidden As Long = 2',
	'Public Type OpenType',
	'    A As Long',
	'End Type',
	'Private Type ClosedType',
	'    B As Long',
	'End Type',
	'Public Enum OpenEnum',
	'    OpenValue',
	'End Enum',
	'Private Enum ClosedEnum',
	'    ClosedValue',
	'End Enum',
	'Public Sub Exported()',
	'End Sub',
	'Private Sub Internal()',
	'End Sub',
	'',
].join('\r\n');

const BETA = 'Public Sub Other()\r\nEnd Sub\r\n';

function project(alpha = ALPHA): ProjectIndex {
	const index = new ProjectIndex();
	index.setModule({ moduleName: 'Alpha', moduleKind: 'standard', source: alpha });
	index.setModule({ moduleName: 'Beta', moduleKind: 'standard', source: BETA });
	return index;
}

function view(index: ProjectIndex, moduleName: string) {
	return {
		procedures: [...index.visibleProcedureNames(moduleName)],
		signatures: index.visibleProcedureSignatures(moduleName).map((sig) => sig.name.toLowerCase()),
		identifiers: [...index.visibleIdentifierNames(moduleName)],
		symbols: index.visibleIdentifierSymbols(moduleName).map((symbol) => symbol.name.toLowerCase()),
		nonTypes: [...index.visibleNonTypeNames(moduleName)],
		types: index.visibleTypeNames(moduleName).map((type) => type.name.toLowerCase()),
		constants: [...index.visibleExternalIntegerConstantExpressions(moduleName).keys()],
		surfaces: index.projectMemberSurfaces(moduleName).map((surface) => surface.name.toLowerCase()),
		alphaMembers: index.projectStandardModuleMembers(moduleName)
			.find((surface) => surface.name === 'Alpha')?.members.map((member) => member.name.toLowerCase()),
	};
}

const PRIVATE_NAMES = ['hidden', 'internal', 'closedtype', 'closedenum', 'closedvalue'];

function privateNamesIn(answer: ReturnType<typeof view>): string[] {
	return Object.values(answer).flatMap((names) => names ?? [])
		.filter((name) => PRIVATE_NAMES.includes(name));
}

describe('ProjectIndex per-module visibility memo', () => {
	for (const order of [['Alpha', 'Beta'], ['Beta', 'Alpha']]) {
		it(`keeps private declarations to their own module when ${order[0]} asks first`, () => {
			const index = project();
			const answers = Object.fromEntries(order.map((name) => [name, view(index, name)]));

			expect(privateNamesIn(answers.Beta)).toEqual([]);
			expect(answers.Beta.procedures).toContain('exported');
			expect(answers.Beta.identifiers).toEqual(expect.arrayContaining(['shared', 'openvalue']));
			expect(answers.Beta.types).toEqual(expect.arrayContaining(['opentype', 'openenum']));
			expect(answers.Beta.constants).toEqual(expect.arrayContaining(['shared', 'alpha.shared']));

			expect(answers.Alpha.procedures).toEqual(expect.arrayContaining(['exported', 'internal', 'other']));
			expect(answers.Alpha.signatures).toEqual(expect.arrayContaining(['internal', 'other']));
			expect(answers.Alpha.identifiers).toEqual(expect.arrayContaining(['hidden', 'closedvalue']));
			expect(answers.Alpha.symbols).toEqual(expect.arrayContaining(['hidden', 'internal']));
			expect(answers.Alpha.types).toEqual(expect.arrayContaining(['closedtype', 'closedenum']));
			expect(answers.Alpha.surfaces).toEqual(expect.arrayContaining(['closedtype', 'closedenum']));
			expect(answers.Alpha.alphaMembers).toEqual(expect.arrayContaining(['hidden', 'internal']));
			// A module never counts its own constants as external.
			expect(answers.Alpha.constants).not.toContain('shared');
		});
	}

	it('answers the same on a warm memo as on a cold one', () => {
		const index = project();
		const cold = { alpha: view(index, 'Alpha'), beta: view(index, 'Beta') };
		expect({ beta: view(index, 'Beta'), alpha: view(index, 'Alpha') })
			.toEqual({ beta: cold.beta, alpha: cold.alpha });
	});

	it('drops both parts of a module when setModule replaces it', () => {
		const index = project();
		view(index, 'Alpha');
		view(index, 'Beta');

		index.setModule({
			moduleName: 'Alpha',
			moduleKind: 'standard',
			source: ALPHA.replace('Public Sub Exported', 'Private Sub Exported') + 'Public Sub AddedLater()\r\nEnd Sub\r\n',
		});

		const beta = view(index, 'Beta');
		expect(beta.procedures).not.toContain('exported');
		expect(beta.procedures).toContain('addedlater');
		expect(beta.alphaMembers).not.toContain('exported');
		const alpha = view(index, 'Alpha');
		expect(alpha.procedures).toEqual(expect.arrayContaining(['exported', 'addedlater']));
	});

	it('updates an exported constant for other modules after an edit', () => {
		const index = project();
		expect(index.visibleExternalIntegerConstantExpressions('Beta').get('shared')).toBe('1');
		index.setModule({
			moduleName: 'Alpha',
			moduleKind: 'standard',
			source: ALPHA.replace('Shared As Long = 1', 'Shared As Long = 7'),
		});
		expect(index.visibleExternalIntegerConstantExpressions('Beta').get('shared')).toBe('7');
	});
});
