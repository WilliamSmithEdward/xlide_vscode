import { describe, expect, it, vi } from 'vitest';
import type * as VscodeType from 'vscode';
import * as path from 'path';

vi.mock('vscode', () => {
	class Position {
		constructor(
			public line: number,
			public character: number,
		) {}
	}
	class Range {
		constructor(
			public start: Position,
			public end: Position,
		) {}
	}
	class Location {
		constructor(
			public uri: VscodeType.Uri,
			public range: Range,
		) {}
	}
	return {
		workspace: { textDocuments: [] },
		Position,
		Range,
		Location,
		Uri: {
			parse: (value: string) => ({
				scheme: value.split(':', 1)[0],
				path: value.replace(/^xlide-vba:/, ''),
				toString: () => value,
			}),
		},
	};
});

import {
	analyzeModule,
	resolveMemberCompletions,
	resolveTypeSemanticTokens,
} from '../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../src/vbaProjectAnalysis';
import {
	projectClassModuleDefinition,
	projectClassReferenceLocations,
	typeReferenceLocations,
	type VbaNavigationModule,
} from '../src/vbaNavigation';
import {
	applyOpenDocumentSources,
	type VbaOpenDocumentLike,
} from '../src/vbaOpenDocuments';
import { projectStandardModuleReferenceLocations } from '../src/vbaStandardModuleRename';

function doc(path: string, source: string): VbaOpenDocumentLike {
	return {
		uri: {
			scheme: 'xlide-vba',
			path,
			toString: () => path,
		} as VscodeType.Uri,
		getText: () => source,
	};
}

function dotOffset(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

describe('multi-workbook isolation', () => {
	it('feeds completion, diagnostics, and semantic coloring from only the requested workbook', () => {
		const workbookOne = path.join(path.sep, 'one', 'book.xlsm');
		const caller = [
			'Public Sub Main()',
			'    Dim p As Person',
			'    Debug.Print p.OneName',
			'    Debug.Print p.TwoName',
			'    SharedProc 1',
			'End Sub',
		].join('\n');
		const modules: VbaProjectModuleInput[] = [
			{ moduleName: 'Caller', type: 'standard', source: caller },
			{
				moduleName: 'Person',
				type: 'class',
				source: 'Public Property Get SavedName() As String\nEnd Property\n',
			},
			{
				moduleName: 'Helpers',
				type: 'standard',
				source: 'Public Sub SharedProc(ByVal value As Long)\nEnd Sub\n',
			},
			{
				moduleName: 'Sheet1',
				type: 'document',
				source: 'Public Sub SavedSheetMember()\nEnd Sub\n',
			},
		];
		const docs = [
			doc('/one/book.xlsm/Person.bas', 'Public Property Get OneName() As String\nEnd Property\n'),
			doc('/two/book.xlsm/Person.bas', 'Public Property Get TwoName() As String\nEnd Property\n'),
			doc(
				'/one/book.xlsm/Helpers.bas',
				'Public Sub SharedProc(ByVal value As Long, ByVal caption As String)\nEnd Sub\n',
			),
			doc(
				'/two/book.xlsm/Helpers.bas',
				'Public Sub SharedProc(ByVal value As Long, ByVal caption As String, ByVal extra As String)\nEnd Sub\n',
			),
			doc('/one/book.xlsm/Sheet1.bas', 'Public Sub OneSheetMember()\nEnd Sub\n'),
			doc('/two/book.xlsm/Sheet1.bas', 'Public Sub TwoSheetMember()\nEnd Sub\n'),
		];

		const overlaid = applyOpenDocumentSources(modules, workbookOne, docs);
		const project = buildVbaProjectIndex(overlaid);
		const options = projectAnalysisOptionsForModule(project, 'Caller');
		const memberSurfaces = options.projectClassMembers ?? [];
		const person = memberSurfaces.find((surface) => surface.name === 'Person');
		const sheet = memberSurfaces.find((surface) => surface.name === 'Sheet1');
		const sharedProc = projectProcedureSignatures(project)?.get('sharedproc');

		expect(person?.members.map((member) => member.name)).toEqual(['OneName']);
		expect(sheet?.members.map((member) => member.name)).toEqual(['OneSheetMember']);
		expect(sharedProc).toHaveLength(1);
		expect(sharedProc?.[0].params.map((param) => param.name)).toEqual(['value', 'caption']);

		const completions = resolveMemberCompletions(
			caller,
			dotOffset(caller, 'p.'),
			{ projectClassMembers: memberSurfaces },
		).map((member) => member.name);
		expect(completions).toContain('OneName');
		expect(completions).not.toContain('TwoName');
		expect(completions).not.toContain('SavedName');

		const diagnostics = analyzeModule(caller, {
			moduleName: 'Caller',
			...options,
		});
		expect(diagnostics.filter((hit) => hit.code === 'member-not-found')).toHaveLength(1);
		expect(diagnostics.find((hit) => hit.code === 'member-not-found')?.message).toContain(
			'Person.TwoName',
		);
		expect(diagnostics.find((hit) => hit.code === 'argument-count')?.message).toContain(
			'expected 2 arguments',
		);

		const typeTokens = resolveTypeSemanticTokens(caller, {
			projectTypes: options.projectTypes,
		}).map((token) => caller.slice(token.span.start, token.span.end));
		expect(typeTokens).toEqual(['Person']);
	});

	it('keeps class reference locations bound to the requested workbook overlay', () => {
		const workbookOne = path.join(path.sep, 'one', 'book.xlsm');
		const caller = [
			'Public Sub Main()',
			'    Dim p As Person',
			'    Set p = New Person',
			'End Sub',
		].join('\n');
		const modules: VbaProjectModuleInput[] = [
			{ moduleName: 'Caller', type: 'standard', source: 'Public Sub Saved()\nEnd Sub\n' },
			{ moduleName: 'Person', type: 'class', source: '' },
		];
		const docs = [
			doc('/one/book.xlsm/Caller.bas', caller),
			doc(
				'/two/book.xlsm/Caller.bas',
				'Public Sub Main()\n    Dim p As Person\n    Set p = New Person\n    Dim q As Person\nEnd Sub\n',
			),
		];

		const overlaid = applyOpenDocumentSources(modules, workbookOne, docs);
		const project = buildVbaProjectIndex(overlaid);
		const byModule = new Map<string, VbaNavigationModule>(
			overlaid.map((mod) => [mod.moduleName.toLowerCase(), mod]),
		);
		const definition = projectClassModuleDefinition(project, 'Caller', 'Person');
		expect(definition).toBeDefined();

		const references = projectClassReferenceLocations(
			workbookOne,
			byModule,
			project,
			'Person',
			definition!,
		);

		expect(references).toHaveLength(2);
		expect(references.map((ref) => ref.uri.toString())).toEqual([
			'xlide-vba:/one/book.xlsm/Caller.bas',
			'xlide-vba:/one/book.xlsm/Caller.bas',
		]);
		expect(references.map((ref) => ref.range.start.line)).toEqual([1, 2]);
	});

	it('keeps qualified type reference locations bound to the named module', () => {
		const workbookOne = path.join(path.sep, 'one', 'book.xlsm');
		const caller = [
			'Public Sub Main()',
			'    Dim a As Geometry.TPoint',
			'    Dim b As OtherGeometry.TPoint',
			'    Dim c As Geometry.TPoint',
			'End Sub',
		].join('\n');
		const modules: VbaProjectModuleInput[] = [
			{ moduleName: 'Caller', type: 'standard', source: caller },
			{
				moduleName: 'Geometry',
				type: 'standard',
				source: 'Public Type TPoint\n    X As Long\nEnd Type\n',
			},
			{
				moduleName: 'OtherGeometry',
				type: 'standard',
				source: 'Public Type TPoint\n    Y As Long\nEnd Type\n',
			},
		];
		const project = buildVbaProjectIndex(modules);
		const byModule = new Map<string, VbaNavigationModule>(
			modules.map((mod) => [mod.moduleName.toLowerCase(), mod]),
		);
		const definitions = project.resolveTypeDefinitions('Caller', 'TPoint')
			.filter((definition) => definition.moduleName === 'Geometry');

		const references = typeReferenceLocations(
			workbookOne,
			byModule,
			project,
			'TPoint',
			definitions,
			false,
		);

		expect(references).toHaveLength(2);
		expect(references.map((ref) => ref.range.start.line)).toEqual([1, 3]);
	});

	it('finds bound standard-module qualifier references for tree rename edits', () => {
		const workbookOne = path.join(path.sep, 'one', 'book.xlsm');
		const caller = [
			'Public Sub Main()',
			'    Helpers.PrintTotal 100',
			'    Alternate.PrintTotal 100',
			'    Debug.Print Helpers.DefaultTaxRate',
			'    Dim p As Helpers.TPoint',
			'    Dim mode As Helpers.SharedMode',
			'    Debug.Print "Helpers.PrintTotal"',
			"    ' Helpers.PrintTotal",
			'End Sub',
		].join('\n');
		const helpers = [
			'Public Const DefaultTaxRate As Double = 0.08',
			'Public Enum SharedMode',
			'    SharedOnly',
			'End Enum',
			'Public Type TPoint',
			'    X As Long',
			'End Type',
			'Public Sub PrintTotal(ByVal amount As Currency)',
			'End Sub',
			'Public Sub SelfCall()',
			'    Helpers.PrintTotal 1',
			'End Sub',
		].join('\n');
		const modules: VbaProjectModuleInput[] = [
			{ moduleName: 'Caller', type: 'standard', source: caller },
			{ moduleName: 'Helpers', type: 'standard', source: helpers },
			{
				moduleName: 'Alternate',
				type: 'standard',
				source: 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n',
			},
		];
		const project = buildVbaProjectIndex(modules);
		const byModule = new Map<string, VbaNavigationModule>(
			modules.map((mod) => [mod.moduleName.toLowerCase(), mod]),
		);

		const references = projectStandardModuleReferenceLocations(
			workbookOne,
			byModule,
			project,
			'Helpers',
			'RenamedHelpers',
		);

		expect(references.map((ref) =>
			`${ref.uri.toString()}:${ref.range.start.line}:${ref.range.start.character}`,
		)).toEqual([
			'xlide-vba:/one/book.xlsm/Caller.bas:1:4',
			'xlide-vba:/one/book.xlsm/Caller.bas:3:16',
			'xlide-vba:/one/book.xlsm/Caller.bas:4:13',
			'xlide-vba:/one/book.xlsm/Caller.bas:5:16',
			'xlide-vba:/one/book.xlsm/RenamedHelpers.bas:10:4',
		]);
	});
});
