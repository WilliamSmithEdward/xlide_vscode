import { describe, it, expect } from 'vitest';
import {
	ProjectIndex,
	resolveProjectTypeSemanticTokens,
	type VbaProjectTypeName,
} from '../src/analyzer';

function tokenTexts(
	source: string,
	projectTypes: readonly VbaProjectTypeName[],
): Array<{ text: string; type: string }> {
	return resolveProjectTypeSemanticTokens(source, projectTypes).map((t) => ({
		text: source.slice(t.span.start, t.span.end),
		type: t.tokenType,
	}));
}

describe('project type semantic tokens', () => {
	it('marks resolved class module names in As type positions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Person\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, projectTypes)).toEqual([
			{ text: 'Person', type: 'class' },
		]);
	});

	it('marks UDTs, enums, parameters, returns, fields, and nested locals', () => {
		const source = [
			'Public Type Wrapper',
			'    Item As Person',
			'End Type',
			'Public Function Make(ByVal mode As Color) As Person',
			'    If True Then',
			'        Dim inner As Wrapper',
			'    End If',
			'End Function',
		].join('\n');
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
			{ name: 'Color', kind: 'enum', moduleName: 'Types' },
			{ name: 'Wrapper', kind: 'userType', moduleName: 'Types' },
		];
		expect(tokenTexts(source, projectTypes)).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'Color', type: 'enum' },
			{ text: 'Person', type: 'class' },
			{ text: 'Wrapper', type: 'struct' },
		]);
	});

	it('does not mark unresolved names or non-type positions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim value As Missing\n' +
			'    Person = 1\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, projectTypes)).toEqual([]);
	});

	it('uses the project binder visible type names as input', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim customer As Person\n' +
			'    Dim state As Status\n' +
			'End Sub\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source,
		});
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: '',
		});
		index.setModule({
			moduleName: 'SharedTypes',
			moduleKind: 'standard',
			source: 'Public Enum Status\n    Active\nEnd Enum\n',
		});

		expect(tokenTexts(source, index.visibleTypeNames('Caller'))).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'Status', type: 'enum' },
		]);
	});
});
