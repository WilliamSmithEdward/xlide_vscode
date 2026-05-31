import { describe, it, expect } from 'vitest';
import {
	buildModuleSymbols,
	ProjectIndex,
	type VbaSymbol,
} from '../src/analyzer';

/** Offset of the first character of the first occurrence of `marker`. */
function offsetOf(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx;
}

/** The text of a symbol's nameSpan inside `src` (proves the span is right). */
function nameText(src: string, symbol: VbaSymbol): string {
	return src.slice(symbol.nameSpan.start, symbol.nameSpan.end);
}

describe('buildModuleSymbols', () => {
	it('extracts procedures with parameters and locals', () => {
		const src = [
			'Option Explicit',
			'',
			'Public Sub DoWork(ByVal count As Long, name As String)',
			'    Dim total As Long',
			'    Dim ws As Worksheet',
			'    If count > 0 Then',
			'        Dim inner As String',
			'    End If',
			'End Sub',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);

		const proc = mod.root.children?.find((c) => c.name === 'DoWork');
		expect(proc).toBeDefined();
		expect(proc?.kind).toBe('sub');
		expect(proc?.visibility).toBe('Public');

		const childNames = (proc?.children ?? []).map((c) => c.name);
		expect(childNames).toEqual(['count', 'name', 'total', 'ws', 'inner']);

		const count = proc?.children?.find((c) => c.name === 'count');
		expect(count?.kind).toBe('parameter');
		expect(count?.asType).toBe('Long');

		const ws = proc?.children?.find((c) => c.name === 'ws');
		expect(ws?.kind).toBe('localVariable');
		expect(ws?.asType).toBe('Worksheet');

		// block-nested local is captured
		const inner = proc?.children?.find((c) => c.name === 'inner');
		expect(inner?.kind).toBe('localVariable');
	});

	it('points nameSpan at the declared identifier, not the whole declaration', () => {
		const src = 'Public Sub DoWork()\nEnd Sub\n';
		const mod = buildModuleSymbols('Module1', 'standard', src);
		const proc = mod.root.children?.[0] as VbaSymbol;
		expect(nameText(src, proc)).toBe('DoWork');
	});

	it('extracts module variables and constants with visibility', () => {
		const src = [
			'Private mState As Long',
			'Public Const MaxItems As Long = 10',
			'Dim untyped',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);

		const state = mod.all.find((s) => s.name === 'mState');
		expect(state?.kind).toBe('moduleVariable');
		expect(state?.visibility).toBe('Private');
		expect(state?.asType).toBe('Long');

		const max = mod.all.find((s) => s.name === 'MaxItems');
		expect(max?.kind).toBe('constant');
		expect(max?.visibility).toBe('Public');

		const untyped = mod.all.find((s) => s.name === 'untyped');
		expect(untyped?.kind).toBe('moduleVariable');
		expect(untyped?.visibility).toBe('Dim');
	});

	it('extracts Type with fields and Enum with members', () => {
		const src = [
			'Public Type TPoint',
			'    X As Double',
			'    Y As Double',
			'End Type',
			'',
			'Public Enum Color',
			'    Red',
			'    Green',
			'End Enum',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);

		const type = mod.root.children?.find((c) => c.name === 'TPoint');
		expect(type?.kind).toBe('type');
		expect((type?.children ?? []).map((c) => c.name)).toEqual(['X', 'Y']);
		expect(type?.children?.[0].kind).toBe('typeField');
		expect(type?.children?.[0].asType).toBe('Double');

		const en = mod.root.children?.find((c) => c.name === 'Color');
		expect(en?.kind).toBe('enum');
		expect((en?.children ?? []).map((c) => c.name)).toEqual(['Red', 'Green']);
		expect(en?.children?.[0].kind).toBe('enumMember');
	});

	it('distinguishes the five procedure kinds', () => {
		const src = [
			'Sub S()',
			'End Sub',
			'Function F() As Long',
			'End Function',
			'Property Get P() As Long',
			'End Property',
			'Property Let P(v As Long)',
			'End Property',
			'Property Set Q(v As Object)',
			'End Property',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);
		const kinds = (mod.root.children ?? []).map((c) => `${c.name}:${c.kind}`);
		expect(kinds).toEqual([
			'S:sub',
			'F:function',
			'P:propertyGet',
			'P:propertyLet',
			'Q:propertySet',
		]);
	});

	it('records the module root kind', () => {
		const mod = buildModuleSymbols('Sheet1', 'document', 'Sub A()\nEnd Sub\n');
		expect(mod.moduleKind).toBe('document');
		expect(mod.root.kind).toBe('module');
		expect(mod.root.name).toBe('Sheet1');
	});
});

describe('ProjectIndex document and workspace symbols', () => {
	it('returns hierarchical document symbols per module', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Public Sub A(x As Long)\n    Dim y As Long\nEnd Sub\n',
		});
		const root = index.documentSymbols('Module1');
		expect(root?.name).toBe('Module1');
		const proc = root?.children?.[0];
		expect(proc?.name).toBe('A');
		expect((proc?.children ?? []).map((c) => c.name)).toEqual(['x', 'y']);
	});

	it('filters workspace symbols by case-insensitive substring', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Public Sub Alpha()\nEnd Sub\nPublic Sub Beta()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Module2',
			moduleKind: 'standard',
			source: 'Public Function AlphaHelper() As Long\nEnd Function\n',
		});
		const names = index.workspaceSymbols('alpha').map((s) => s.name).sort();
		expect(names).toEqual(['Alpha', 'AlphaHelper']);
	});

	it('lists indexed module names and supports removal', () => {
		const index = new ProjectIndex();
		index.setModule({ moduleName: 'A', moduleKind: 'standard', source: '' });
		index.setModule({ moduleName: 'B', moduleKind: 'class', source: '' });
		expect(index.moduleNames().sort()).toEqual(['A', 'B']);
		index.removeModule('A');
		expect(index.moduleNames()).toEqual(['B']);
	});
});

describe('ProjectIndex name resolution (go-to-definition)', () => {
	const index = new ProjectIndex();
	const mod1 = [
		'Public gShared As Long',
		'Private mPrivate As Long',
		'',
		'Public Sub Caller()',
		'    Dim local As Long',
		'    local = Helper(gShared)',
		'End Sub',
		'',
		'Public Function Helper(value As Long) As Long',
		'    Helper = value',
		'End Function',
	].join('\n');
	const mod2 = [
		'Public Sub OtherEntry()',
		'End Sub',
		'',
		'Private Sub Secret()',
		'End Sub',
	].join('\n');
	index.setModule({ moduleName: 'Module1', moduleKind: 'standard', source: mod1 });
	index.setModule({ moduleName: 'Module2', moduleKind: 'standard', source: mod2 });

	it('resolves a local variable to its declaration in the enclosing procedure', () => {
		const useOffset = mod1.lastIndexOf('local');
		const hits = index.resolveDefinition('Module1', 'local', useOffset);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('localVariable');
		expect(hits[0].containerName).toBe('Caller');
	});

	it('resolves a parameter ahead of a same-named module symbol', () => {
		const useOffset = mod1.lastIndexOf('value');
		const hits = index.resolveDefinition('Module1', 'value', useOffset);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('parameter');
		expect(hits[0].containerName).toBe('Helper');
	});

	it('resolves a module-level variable used inside a procedure', () => {
		const useOffset = mod1.indexOf('gShared)');
		const hits = index.resolveDefinition('Module1', 'gShared', useOffset);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('moduleVariable');
	});

	it('resolves a call to a procedure in the same module', () => {
		const useOffset = mod1.indexOf('Helper(gShared');
		const hits = index.resolveDefinition('Module1', 'Helper', useOffset);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('function');
		expect(hits[0].moduleName).toBe('Module1');
	});

	it('resolves a public procedure declared in another module', () => {
		const hits = index.resolveDefinition('Module1', 'OtherEntry', 0);
		expect(hits).toHaveLength(1);
		expect(hits[0].moduleName).toBe('Module2');
		expect(hits[0].kind).toBe('sub');
	});

	it('does not resolve a Private procedure across modules', () => {
		const hits = index.resolveDefinition('Module1', 'Secret', 0);
		expect(hits).toHaveLength(0);
	});

	it('does not resolve a Private module variable across modules', () => {
		const hits = index.resolveDefinition('Module2', 'mPrivate', 0);
		expect(hits).toHaveLength(0);
	});

	it('returns an empty array for unknown identifiers', () => {
		expect(index.resolveDefinition('Module1', 'Nonexistent', 0)).toEqual([]);
	});
});

describe('ProjectIndex property and enum resolution', () => {
	it('resolves both Property Get and Let sharing a name', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Class1',
			moduleKind: 'class',
			source: [
				'Public Property Get Value() As Long',
				'End Property',
				'Public Property Let Value(v As Long)',
				'End Property',
			].join('\n'),
		});
		const hits = index.resolveDefinition('Class1', 'Value', 0);
		expect(hits.map((h) => h.kind).sort()).toEqual([
			'propertyGet',
			'propertyLet',
		]);
	});

	it('resolves an enum member by its bare name at module scope', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Public Enum Color\n    Red\n    Green\nEnd Enum\n',
		});
		const hits = index.resolveDefinition('Module1', 'Green', 0);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('enumMember');
		expect(hits[0].containerName).toBe('Color');
	});
});

describe('ProjectIndex duplicate procedure detection', () => {
	it('flags procedures declared twice in the same module', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Sub A()\nEnd Sub\nSub A()\nEnd Sub\nSub B()\nEnd Sub\n',
		});
		const dupes = index.duplicateProcedures('Module1');
		expect(dupes).toHaveLength(2);
		expect(dupes.every((d) => d.name === 'A')).toBe(true);
	});

	it('reports no duplicates for unique procedure names', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Sub A()\nEnd Sub\nSub B()\nEnd Sub\n',
		});
		expect(index.duplicateProcedures('Module1')).toEqual([]);
	});
});

describe('ProjectIndex procedure signatures', () => {
	it('collects exported Sub and Function signatures for project diagnostics', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Helpers',
			moduleKind: 'standard',
			source:
				'Public Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double) As Currency\n' +
				'End Function\n' +
				'Private Sub Hidden(ByVal value As String)\nEnd Sub\n',
		});
		const signatures = index.procedureSignatures();
		expect(signatures.get('hidden')).toBeUndefined();
		const invoice = signatures.get('invoicetotal');
		expect(invoice).toHaveLength(1);
		expect(invoice?.[0].moduleName).toBe('Helpers');
		expect(invoice?.[0].returnType).toBe('Currency');
		expect(signatures.get('helpers.invoicetotal')).toEqual(invoice);
		expect(invoice?.[0].params).toEqual([
			{
				name: 'Subtotal',
				type: 'Currency',
				optional: false,
				paramArray: false,
				isArray: false,
			},
			{
				name: 'TaxRate',
				type: 'Double',
				optional: true,
				paramArray: false,
				isArray: false,
			},
		]);
	});

	it('keeps duplicate exported signatures grouped for ambiguity checks', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'First',
			moduleKind: 'standard',
			source: 'Public Sub DoWork(ByVal value As Long)\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Second',
			moduleKind: 'standard',
			source: 'Public Sub DoWork(ByVal value As Long)\nEnd Sub\n',
		});
		expect(index.procedureSignatures().get('dowork')).toHaveLength(2);
		expect(index.procedureSignatures().get('first.dowork')).toHaveLength(1);
		expect(index.procedureSignatures().get('second.dowork')).toHaveLength(1);
	});

	it('does not expose class-module members through standard-module procedure signatures', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: 'Public Sub Save(ByVal caption As String)\nEnd Sub\n',
		});
		expect(index.procedureSignatures().get('save')).toBeUndefined();
		expect(index.procedureSignatures().get('customer.save')).toBeUndefined();
	});
});

describe('ProjectIndex visible procedure names', () => {
	it('includes same-module procedures and exported standard-module procedures', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: 'Private Sub LocalOnly()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Helpers',
			moduleKind: 'standard',
			source:
				'Sub DefaultPublic()\nEnd Sub\n' +
				'Public Sub ExplicitPublic()\nEnd Sub\n' +
				'Private Sub Hidden()\nEnd Sub\n',
		});

		const names = [...index.visibleProcedureNames('Caller')].sort();
		expect(names).toEqual(['defaultpublic', 'explicitpublic', 'localonly']);
	});

	it('does not expose object-module members as bare cross-module procedures', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: '',
		});
		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: 'Public Sub Save()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Sheet1',
			moduleKind: 'document',
			source: 'Public Sub ActivateSheet()\nEnd Sub\n',
		});
		expect(index.visibleProcedureNames('Caller')).toEqual(new Set<string>());
	});
});

describe('ProjectIndex visible type names', () => {
	it('includes object modules and visible project Type/Enum declarations', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: [
				'Private Type LocalRecord',
				'    Value As Long',
				'End Type',
				'Private Enum LocalMode',
				'    LocalOnly',
				'End Enum',
			].join('\n'),
		});
		index.setModule({
			moduleName: 'SharedTypes',
			moduleKind: 'standard',
			source: [
				'Public Type ExportedRecord',
				'    Value As Long',
				'End Type',
				'Enum ExportedMode',
				'    First',
				'End Enum',
				'Private Type HiddenRecord',
				'    Value As Long',
				'End Type',
			].join('\n'),
		});
		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: 'Public Sub Save()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Sheet1',
			moduleKind: 'document',
			source: '',
		});
		index.setModule({
			moduleName: 'OrderForm',
			moduleKind: 'userform',
			source: '',
		});

		const visible = index.visibleTypeNames('Module1');
		const labels = visible
			.map((t) => `${t.name}:${t.kind}:${t.moduleName}`)
			.sort();
		expect(labels).toEqual([
			'Customer:class:Customer',
			'ExportedMode:enum:SharedTypes',
			'ExportedRecord:userType:SharedTypes',
			'LocalMode:enum:Module1',
			'LocalRecord:userType:Module1',
			'OrderForm:userform:OrderForm',
			'Sheet1:document:Sheet1',
		]);
	});

	it('does not expose Private Type/Enum declarations from other modules', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: '',
		});
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: [
				'Private Type HiddenRecord',
				'    Value As Long',
				'End Type',
				'Private Enum HiddenMode',
				'    Hidden',
				'End Enum',
			].join('\n'),
		});
		expect(index.visibleTypeNames('Caller').map((t) => t.name)).toEqual([]);
	});

	it('preserves duplicate visible type names for future ambiguity handling', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'First',
			moduleKind: 'standard',
			source: 'Public Type Customer\n    Id As Long\nEnd Type\n',
		});
		index.setModule({
			moduleName: 'Second',
			moduleKind: 'class',
			source: '',
		});
		index.setModule({
			moduleName: 'Consumer',
			moduleKind: 'standard',
			source: '',
		});
		const names = index.visibleTypeNames('Consumer').filter(
			(t) => t.name.toLowerCase() === 'customer',
		);
		expect(names).toHaveLength(1);

		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: '',
		});
		const duplicates = index.visibleTypeNames('Consumer').filter(
			(t) => t.name.toLowerCase() === 'customer',
		);
		expect(duplicates).toHaveLength(2);
		expect(duplicates.map((t) => t.kind).sort()).toEqual(['class', 'userType']);
	});
});

describe('ProjectIndex project class members', () => {
	it('exposes source-declared public/default-public class members', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: [
				'Public Name As String',
				'Private Secret As String',
				"''' <summary>Age in whole years.</summary>",
				'Public Property Get Age() As Long',
				'End Property',
				'Public Property Let Age(ByVal value As Long)',
				'End Property',
				'Sub Save()',
				'End Sub',
				'Private Sub Hidden()',
				'End Sub',
				'Public Function Manager() As Person',
				'End Function',
			].join('\n'),
		});
		const person = index.projectClassMembers().find((t) => t.name === 'Person');
		expect(person?.members.map((m) => `${m.name}:${m.kind}:${m.returns ?? ''}`)).toEqual([
			'Name:property:String',
			'Age:property:Long',
			'Save:method:',
			'Manager:method:Person',
		]);
		const age = person?.members.find((m) => m.name === 'Age');
		expect(age?.doc?.summary).toBe('Age in whole years.');
		expect(age?.writable).toBe(true);
		expect(age?.writeType).toBe('Long');
		const save = person?.members.find((m) => m.name === 'Save');
		expect(save?.writable).toBeUndefined();
	});

	it('marks Property Get-only members and constants as read-only', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: [
				'Public Const Species As String = "Human"',
				'Public Property Get Age() As Long',
				'End Property',
			].join('\n'),
		});
		const person = index.projectClassMembers().find((t) => t.name === 'Person');
		expect(person?.members.find((m) => m.name === 'Age')?.writable).toBe(false);
		expect(person?.members.find((m) => m.name === 'Species')?.writable).toBe(false);
	});
});

describe('ProjectIndex resolveQualifiedDefinition', () => {
	const index = new ProjectIndex();
	index.setModule({
		moduleName: 'Module1',
		moduleKind: 'standard',
		source: 'Public Sub DoWork()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\n',
	});
	index.setModule({
		moduleName: 'Module2',
		moduleKind: 'standard',
		source: 'Public Sub DoWork()\nEnd Sub\n',
	});

	it('resolves an exported member of the named module only', () => {
		const hits = index.resolveQualifiedDefinition('Module1', 'DoWork');
		expect(hits).toHaveLength(1);
		expect(hits[0].moduleName).toBe('Module1');
	});

	it('does not resolve a Private member through a qualifier', () => {
		expect(index.resolveQualifiedDefinition('Module1', 'Hidden')).toEqual([]);
	});

	it('returns an empty array for an unknown qualifier module', () => {
		expect(index.resolveQualifiedDefinition('Nope', 'DoWork')).toEqual([]);
	});
});

describe('ProjectIndex referenceScope', () => {
	it('limits a local variable to its enclosing procedure', () => {
		const index = new ProjectIndex();
		const src = [
			'Sub A()',
			'    Dim total As Long',
			'    total = 1',
			'End Sub',
			'Sub B()',
			'    Dim total As Long',
			'    total = 2',
			'End Sub',
		].join('\n');
		index.setModule({ moduleName: 'Module1', moduleKind: 'standard', source: src });

		const offset = src.indexOf('total = 1');
		const scope = index.referenceScope('Module1', 'total', offset);
		expect(scope.kind).toBe('local');
		expect(scope.searchModules).toEqual(['Module1']);
		expect(scope.procedureSpan).toBeDefined();
		// The procedure span covers Sub A but not Sub B.
		expect(scope.procedureSpan!.start).toBeLessThan(src.indexOf('total = 1'));
		expect(scope.procedureSpan!.end).toBeLessThan(src.indexOf('Sub B'));
		expect(scope.definitions).toHaveLength(1);
		expect(scope.definitions[0].containerName).toBe('A');
	});

	it('limits a Private module variable to its own module', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Private mState As Long\nSub A()\n    mState = 1\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Module2',
			moduleKind: 'standard',
			source: 'Sub B()\n    Dim mState As Long\nEnd Sub\n',
		});
		const scope = index.referenceScope('Module1', 'mState', 0);
		expect(scope.kind).toBe('module');
		expect(scope.searchModules).toEqual(['Module1']);
	});

	it('spans every module for an exported procedure', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Public Sub Shared()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Module2',
			moduleKind: 'standard',
			source: 'Sub Caller()\n    Shared\nEnd Sub\n',
		});
		const scope = index.referenceScope('Module2', 'Shared', 0);
		expect(scope.kind).toBe('project');
		expect(scope.searchModules.sort()).toEqual(['Module1', 'Module2']);
	});

	it('excludes a module that re-declares the name privately', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Public gValue As Long\n',
		});
		index.setModule({
			moduleName: 'Module2',
			moduleKind: 'standard',
			source: 'Private gValue As Long\n',
		});
		index.setModule({
			moduleName: 'Module3',
			moduleKind: 'standard',
			source: 'Sub Use()\n    gValue = 1\nEnd Sub\n',
		});
		const scope = index.referenceScope('Module1', 'gValue', 0);
		expect(scope.kind).toBe('project');
		expect(scope.searchModules).toContain('Module1');
		expect(scope.searchModules).toContain('Module3');
		expect(scope.searchModules).not.toContain('Module2');
	});

	it('records procedure spans that shadow an exported name with a local', () => {
		const index = new ProjectIndex();
		const src = [
			'Public gValue As Long',
			'Sub Shadower()',
			'    Dim gValue As Long',
			'    gValue = 1',
			'End Sub',
		].join('\n');
		index.setModule({ moduleName: 'Module1', moduleKind: 'standard', source: src });
		const scope = index.referenceScope('Module1', 'gValue', 0);
		expect(scope.kind).toBe('project');
		expect(scope.shadowedSpans).toHaveLength(1);
		expect(scope.shadowedSpans[0].moduleName).toBe('Module1');
	});

	it('keeps an unresolved name inside the home module', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Module1',
			moduleKind: 'standard',
			source: 'Sub A()\n    Range("A1").Select\nEnd Sub\n',
		});
		const scope = index.referenceScope('Module1', 'Range', 0);
		expect(scope.kind).toBe('module');
		expect(scope.searchModules).toEqual(['Module1']);
		expect(scope.definitions).toEqual([]);
	});
});
