import { describe, it, expect } from 'vitest';
import {
	ProjectIndex,
	resolveTypeSemanticTokens,
	type TypeCompletionContext,
	type VbaProjectTypeName,
} from '../src/analyzer';

function tokenTexts(
	source: string,
	ctx: TypeCompletionContext = {},
): Array<{ text: string; type: string }> {
	return resolveTypeSemanticTokens(source, ctx).map((t) => ({
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
		expect(tokenTexts(source, { projectTypes })).toEqual([
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
		expect(tokenTexts(source, { projectTypes })).toEqual([
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
		expect(tokenTexts(source, { projectTypes })).toEqual([]);
	});

	it('marks project classes in New expressions', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New person\n' +
			'    If TypeOf p Is Person Then Debug.Print "ok"\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'person', type: 'class' },
			{ text: 'Person', type: 'class' },
		]);
	});

	it('marks project classes in Implements statements', () => {
		const source = 'Implements Person\nImplements Excel.Worksheet\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Person', type: 'class' },
		]);
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

		expect(tokenTexts(source, { projectTypes: index.visibleTypeNames('Caller') })).toEqual([
			{ text: 'Person', type: 'class' },
			{ text: 'Status', type: 'enum' },
		]);
	});

	it('marks qualified project type names in As, New, and TypeOf positions', () => {
		const source = [
			'Public Sub T()',
			'    Dim point As Geometry.TPoint',
			'    Dim state As Workflow.Status',
			'    Set point = New Geometry.TPoint',
			'    If TypeOf point Is Geometry.TPoint Then Debug.Print "ok"',
			'End Sub',
		].join('\n');
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'TPoint', kind: 'class', moduleName: 'Geometry' },
			{ name: 'Status', kind: 'enum', moduleName: 'Workflow' },
			{ name: 'TPoint', kind: 'class', moduleName: 'OtherGeometry' },
		];

		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'TPoint', type: 'class' },
			{ text: 'Status', type: 'enum' },
			{ text: 'TPoint', type: 'class' },
			{ text: 'TPoint', type: 'class' },
		]);
	});
});

describe('type semantic tokens', () => {
	it('marks primitive, host, and project types with distinct token categories', () => {
		const source =
			'Public Sub T()\n' +
			'    Dim amount As Currency\n' +
			'    Dim p As Person\n' +
			'    Dim ws As Worksheet\n' +
			'    Dim state As Status\n' +
			'    Dim point As TPoint\n' +
			'End Sub\n';
		const projectTypes: VbaProjectTypeName[] = [
			{ name: 'Person', kind: 'class', moduleName: 'Person' },
			{ name: 'Status', kind: 'enum', moduleName: 'Types' },
			{ name: 'TPoint', kind: 'userType', moduleName: 'Types' },
		];
		expect(tokenTexts(source, { projectTypes })).toEqual([
			{ text: 'Currency', type: 'type' },
			{ text: 'Person', type: 'class' },
			{ text: 'Worksheet', type: 'class' },
			{ text: 'Status', type: 'enum' },
			{ text: 'TPoint', type: 'struct' },
		]);
	});

	it('lets project types shadow primitive names for coloring', () => {
		const source = 'Public Sub T()\n    Dim value As Long\nEnd Sub\n';
		expect(
			tokenTexts(source, {
				projectTypes: [{ name: 'Long', kind: 'class', moduleName: 'Long' }],
			}),
		).toEqual([{ text: 'Long', type: 'class' }]);
	});

	it('colors colliding project type names generically', () => {
		const source = 'Public Sub T()\n    Dim value As Status\nEnd Sub\n';
		expect(
			tokenTexts(source, {
				projectTypes: [
					{ name: 'Status', kind: 'class', moduleName: 'StatusClass' },
					{ name: 'Status', kind: 'enum', moduleName: 'SharedTypes' },
				],
			}),
		).toEqual([{ text: 'Status', type: 'type' }]);
	});
});
