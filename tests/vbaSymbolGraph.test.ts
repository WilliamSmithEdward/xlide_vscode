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
			'Private mBuffer As String * 20',
			'Public Const MaxItems As Long = 10',
			'Dim untyped',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);

		const state = mod.all.find((s) => s.name === 'mState');
		expect(state?.kind).toBe('moduleVariable');
		expect(state?.visibility).toBe('Private');
		expect(state?.asType).toBe('Long');

		const buffer = mod.all.find((s) => s.name === 'mBuffer');
		expect(buffer?.kind).toBe('moduleVariable');
		expect(buffer?.asType).toBe('String');
		expect(buffer?.fixedLength).toBe('20');

		const max = mod.all.find((s) => s.name === 'MaxItems');
		expect(max?.kind).toBe('constant');
		expect(max?.visibility).toBe('Public');
		expect(max?.asType).toBe('Long');
		expect(max?.defaultRaw).toBe('10');

		const untyped = mod.all.find((s) => s.name === 'untyped');
		expect(untyped?.kind).toBe('moduleVariable');
		expect(untyped?.visibility).toBe('Dim');
	});

	it('extracts Type with fields and Enum with members', () => {
		const src = [
			'Public Type TPoint',
			'    X As Double',
			'    Label As String * 12',
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
		expect((type?.children ?? []).map((c) => c.name)).toEqual(['X', 'Label', 'Y']);
		expect(type?.children?.[0].kind).toBe('typeField');
		expect(type?.children?.[0].asType).toBe('Double');
		expect(type?.children?.[1].asType).toBe('String');
		expect(type?.children?.[1].fixedLength).toBe('12');

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

	it('extracts Event declarations with parameters', () => {
		const src = 'Public Event Changed(ByVal value As Long, ByRef cancel As Boolean)\n';
		const mod = buildModuleSymbols('Notifier', 'class', src);
		const event = mod.root.children?.find((child) => child.name === 'Changed');

		expect(event?.kind).toBe('event');
		expect(event?.visibility).toBe('Public');
		expect((event?.children ?? []).map((child) => `${child.name}:${child.asType}:${child.byVal}:${child.byRef}`))
			.toEqual(['value:Long:true:false', 'cancel:Boolean:false:true']);
	});

	it('records the module root kind', () => {
		const mod = buildModuleSymbols('Sheet1', 'document', 'Sub A()\nEnd Sub\n');
		expect(mod.moduleKind).toBe('document');
		expect(mod.root.kind).toBe('module');
		expect(mod.root.name).toBe('Sheet1');
	});

	it('attaches a leading doc block before Option as module documentation', () => {
		const mod = buildModuleSymbols(
			'Person',
			'class',
			"''' <summary>Represents a person.</summary>\nOption Explicit\nPublic Sub Save()\nEnd Sub\n",
		);
		expect(mod.root.doc?.summary).toBe('Represents a person.');
		expect(mod.root.children?.find((child) => child.name === 'Save')?.doc).toBeUndefined();
	});

	it('extracts external Declare signatures with library metadata', () => {
		const src =
			"''' <summary>Finds a top-level window.</summary>\n" +
			'Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal ClassName As String, ByVal WindowName As String) As LongPtr\n';
		const mod = buildModuleSymbols('NativeApi', 'standard', src);
		const decl = mod.root.children?.find((child) => child.name === 'FindWindow');
		expect(decl?.kind).toBe('declare');
		expect(decl?.declareKind).toBe('Function');
		expect(decl?.ptrSafe).toBe(true);
		expect(decl?.libName).toBe('user32');
		expect(decl?.aliasName).toBe('FindWindowA');
		expect(decl?.asType).toBe('LongPtr');
		expect(decl?.doc?.summary).toBe('Finds a top-level window.');
		expect((decl?.children ?? []).map((child) => `${child.name}:${child.asType}:${child.byVal}`)).toEqual([
			'ClassName:String:true',
			'WindowName:String:true',
		]);
	});

	it('filters only proven-inactive conditional declarations from the module symbol graph', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n';

		const defaultVba7 = buildModuleSymbols('NativeApi', 'standard', src);
		const defaultSleep = defaultVba7.root.children?.filter((sym) => sym.name === 'Sleep') ?? [];
		expect(defaultSleep).toHaveLength(1);
		expect(defaultSleep[0].ptrSafe).toBe(true);

		const vba7 = buildModuleSymbols('NativeApi', 'standard', src, {
			conditionalCompilation: { compilerConstants: { VBA7: true } },
		});
		const sleep = vba7.root.children?.filter((sym) => sym.name === 'Sleep') ?? [];
		expect(sleep).toHaveLength(1);
		expect(sleep[0].ptrSafe).toBe(true);
		expect(sleep[0].children?.[0].asType).toBe('LongPtr');

		const legacy = buildModuleSymbols('NativeApi', 'standard', src, {
			conditionalCompilation: { compilerConstants: { VBA7: false } },
		});
		const legacySleep = legacy.root.children?.filter((sym) => sym.name === 'Sleep') ?? [];
		expect(legacySleep).toHaveLength(1);
		expect(legacySleep[0].ptrSafe).toBe(false);
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

	it('explains project-level ambiguous bare identifier bindings', () => {
		const index = new ProjectIndex();
		const first = 'Public Const SharedValue As Long = 1\n';
		const second = 'Public Const SharedValue As Long = 2\n';
		const caller = 'Sub Use()\n    Debug.Print SharedValue\nEnd Sub\n';
		index.setModule({ moduleName: 'First', moduleKind: 'standard', source: first });
		index.setModule({ moduleName: 'Second', moduleKind: 'standard', source: second });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: caller });

		const resolved = index.resolveBareIdentifier(
			'Caller',
			'SharedValue',
			offsetOf(caller, 'SharedValue'),
			'expression',
		);

		expect(resolved.scope).toBe('ambiguous');
		expect(resolved.tier).toBe('project');
		expect(resolved.definitions.map((symbol) => symbol.moduleName).sort()).toEqual([
			'First',
			'Second',
		]);
	});

	it('models type-name context separately from local value shadowing', () => {
		const index = new ProjectIndex();
		const types = 'Public Type Customer\n    Id As Long\nEnd Type\n';
		const caller =
			'Sub Use()\n' +
			'    Dim Customer As Long\n' +
			'    Dim item As Customer\n' +
			'    Customer = 1\n' +
			'End Sub\n';
		index.setModule({ moduleName: 'Types', moduleKind: 'standard', source: types });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: caller });

		const expression = index.resolveBareIdentifier(
			'Caller',
			'Customer',
			offsetOf(caller, 'Customer = 1'),
			'expression',
		);
		const typeName = index.resolveBareIdentifier(
			'Caller',
			'Customer',
			offsetOf(caller, 'As Customer') + 'As '.length,
			'typeName',
		);

		expect(expression.scope).toBe('local');
		expect(expression.definitions[0].kind).toBe('localVariable');
		expect(typeName.scope).toBe('project');
		expect(typeName.definitions[0].kind).toBe('type');
		expect(typeName.definitions[0].moduleName).toBe('Types');
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

	it('resolves exported enum members across standard modules', () => {
		const index = new ProjectIndex();
		const globals = 'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n';
		const caller = 'Sub Use()\n    value = SharedOnly\nEnd Sub\n';
		index.setModule({ moduleName: 'Globals', moduleKind: 'standard', source: globals });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: caller });

		const hits = index.resolveDefinition(
			'Caller',
			'SharedOnly',
			offsetOf(caller, 'SharedOnly'),
		);
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('enumMember');
		expect(hits[0].moduleName).toBe('Globals');
		expect(hits[0].containerName).toBe('SharedMode');
		expect(nameText(globals, hits[0])).toBe('SharedOnly');
	});

	it('keeps private enum members private to their module', () => {
		const index = new ProjectIndex();
		const globals = 'Private Enum HiddenMode\n    HiddenOnly\nEnd Enum\n';
		const caller = 'Sub Use()\n    value = HiddenOnly\nEnd Sub\n';
		index.setModule({ moduleName: 'Globals', moduleKind: 'standard', source: globals });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: caller });

		expect(
			index.resolveDefinition('Caller', 'HiddenOnly', offsetOf(caller, 'HiddenOnly')),
		).toEqual([]);

		const localHits = index.resolveDefinition(
			'Globals',
			'HiddenOnly',
			offsetOf(globals, 'HiddenOnly'),
		);
		expect(localHits).toHaveLength(1);
		expect(localHits[0].kind).toBe('enumMember');
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
				"''' <summary>Calculates the invoice total.</summary>\n" +
				'Public Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double = 0.08) As Currency\n' +
				'End Function\n' +
				'Private Sub Hidden(ByVal value As String)\nEnd Sub\n',
		});
		const signatures = index.procedureSignatures();
		expect(signatures.get('hidden')).toBeUndefined();
		const invoice = signatures.get('invoicetotal');
		expect(invoice).toHaveLength(1);
		expect(invoice?.[0].moduleName).toBe('Helpers');
		expect(invoice?.[0].returnType).toBe('Currency');
		expect(invoice?.[0].signature).toBe(
			'InvoiceTotal(Subtotal As Currency, [TaxRate As Double = 0.08]) As Currency',
		);
		expect(invoice?.[0].doc?.summary).toBe('Calculates the invoice total.');
		expect(signatures.get('helpers.invoicetotal')).toEqual(invoice);
		expect(invoice?.[0].params).toEqual([
			{
				name: 'Subtotal',
				type: 'Currency',
				optional: false,
				paramArray: false,
				isArray: false,
				byVal: true,
			},
			{
				name: 'TaxRate',
				type: 'Double',
				optional: true,
				paramArray: false,
				isArray: false,
				defaultRaw: '0.08',
				byVal: true,
			},
		]);
	});

	it('collects exported Declare signatures for project diagnostics', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'NativeApi',
			moduleKind: 'standard',
			source:
				'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n',
		});

		const sleep = index.procedureSignatures().get('sleep');
		expect(sleep).toHaveLength(1);
		expect(sleep?.[0]).toMatchObject({
			name: 'Sleep',
			moduleName: 'NativeApi',
			kind: 'sub',
			external: true,
			ptrSafe: true,
			libName: 'kernel32',
			signature: 'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)',
		});
		expect(sleep?.[0].params).toEqual([
			{
				name: 'Milliseconds',
				type: 'LongPtr',
				optional: false,
				paramArray: false,
				isArray: false,
				byVal: true,
			},
		]);
		expect(index.procedureSignatures().get('nativeapi.sleep')).toEqual(sleep);
	});

	it('uses the same conditional branch filtering for project signatures', () => {
		const index = new ProjectIndex({
			conditionalCompilation: { compilerConstants: { VBA7: false } },
		});
		index.setModule({
			moduleName: 'NativeApi',
			moduleKind: 'standard',
			source:
				'#If VBA7 Then\n' +
				'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
				'#Else\n' +
				'Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
				'#End If\n',
		});

		const sleep = index.procedureSignatures().get('sleep');
		expect(sleep).toHaveLength(1);
		expect(sleep?.[0].ptrSafe).toBe(false);
		expect(sleep?.[0].params[0].type).toBe('Long');
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

describe('ProjectIndex visible procedure signatures', () => {
	it('returns callable same-module and exported standard-module Sub/Function/Declare signatures', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source:
				'Private Sub LocalOnly(ByVal label As String)\nEnd Sub\n' +
				'Function LocalTotal() As Currency\nEnd Function\n',
		});
		index.setModule({
			moduleName: 'Helpers',
			moduleKind: 'standard',
			source:
				'Sub DefaultPublic()\nEnd Sub\n' +
				"''' <summary>Calculates the invoice total.</summary>\n" +
				'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\nEnd Function\n' +
				'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
				'Private Sub Hidden()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: 'Public Sub Save()\nEnd Sub\n',
		});

		const got = index.visibleProcedureSignatures('Caller');
		expect(got.map((sig) => `${sig.moduleName}.${sig.name}`).sort()).toEqual([
			'Caller.LocalOnly',
			'Caller.LocalTotal',
			'Helpers.DefaultPublic',
			'Helpers.InvoiceTotal',
			'Helpers.Sleep',
		]);
		expect(got.find((sig) => sig.name === 'InvoiceTotal')?.returnType).toBe('Currency');
		expect(got.find((sig) => sig.name === 'InvoiceTotal')?.params[0]).toEqual({
			name: 'Subtotal',
			type: 'Currency',
			optional: false,
			paramArray: false,
			isArray: false,
			byVal: true,
		});
		expect(got.find((sig) => sig.name === 'InvoiceTotal')?.signature).toBe(
			'InvoiceTotal(Subtotal As Currency) As Currency',
		);
		expect(got.find((sig) => sig.name === 'InvoiceTotal')?.doc?.summary).toBe(
			'Calculates the invoice total.',
		);
		expect(got.map((sig) => sig.name)).not.toContain('Hidden');
		expect(got.map((sig) => sig.name)).not.toContain('Save');
		expect(got.find((sig) => sig.name === 'Sleep')?.external).toBe(true);
	});
});

describe('ProjectIndex visible identifier names', () => {
	it('includes same-module names, exported standard-module globals, and visible enum members', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source:
				'Private localState As Long\n' +
				'Private Enum LocalMode\n    LocalOnly\nEnd Enum\n',
		});
		index.setModule({
			moduleName: 'Globals',
			moduleKind: 'standard',
			source:
				'Public sharedValue As Long\n' +
				'Private hiddenValue As Long\n' +
				'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n',
		});
		index.setModule({
			moduleName: 'Sheet1',
			moduleKind: 'document',
			source: 'Public Sub Change()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'UserForm1',
			moduleKind: 'userform',
			source: '',
		});

		const names = [...index.visibleIdentifierNames('Caller')].sort();
		expect(names).toEqual([
			'localmode',
			'localonly',
			'localstate',
			'sharedmode',
			'sharedonly',
			'sharedvalue',
			'sheet1',
			'userform1',
		]);
	});

	it('does not expose private standard-module variables or object members as bare globals', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: '',
		});
		index.setModule({
			moduleName: 'Globals',
			moduleKind: 'standard',
			source: 'Private hiddenValue As Long\n',
		});
		index.setModule({
			moduleName: 'Customer',
			moduleKind: 'class',
			source: 'Public Sub Save()\nEnd Sub\n',
		});

		expect(index.visibleIdentifierNames('Caller')).toEqual(new Set<string>());
	});

	it('exposes visible external integer constant expressions without guessing duplicates', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: 'Public Const LocalBadLength As Long = -1\n',
		});
		index.setModule({
			moduleName: 'Globals',
			moduleKind: 'standard',
			source:
				'Public Const SharedBadLength As Long = -1\n' +
				'Public Const SharedGoodLength As Long = 2\n' +
				'Private Const HiddenBadLength As Long = -1\n' +
				'Private Const HiddenBaseLength As Long = -2\n' +
				'Public Const SharedHiddenLength As Long = HiddenBaseLength + 1\n' +
				'Public Enum SharedStart\n' +
				'    SharedBadStart = 0\n' +
				'    SharedNextStart\n' +
				'End Enum\n',
		});
		index.setModule({
			moduleName: 'MoreGlobals',
			moduleKind: 'standard',
			source: 'Public Const SharedBadLength As Long = 0\n',
		});
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: 'Public Const ClassValue As Long = -1\n',
		});

		const constants = index.visibleExternalIntegerConstantExpressions('Caller');
		expect(constants.get('localbadlength')).toBeUndefined();
		expect(constants.get('hiddenbadlength')).toBeUndefined();
		expect(constants.get('hiddenbaselength')).toBeUndefined();
		expect(constants.get('globals.hiddenbaselength')).toBeUndefined();
		expect(constants.get('classvalue')).toBeUndefined();
		expect(constants.get('sharedbadlength')).toBeUndefined();
		expect(constants.get('sharedgoodlength')).toBe('2');
		expect(constants.get('sharedhiddenlength')).toBe('-1');
		expect(constants.get('sharedbadstart')).toBe('0');
		expect(constants.get('sharednextstart')).toBe('1');
		expect(constants.get('globals.sharedbadlength')).toBe('-1');
		expect(constants.get('moreglobals.sharedbadlength')).toBe('0');
		expect(constants.get('globals.sharedhiddenlength')).toBe('-1');
		expect(constants.get('globals.sharedbadstart')).toBe('0');
		expect(constants.get('globals.sharednextstart')).toBe('1');
	});
});

describe('ProjectIndex visible non-type names', () => {
	it('includes visible procedures, globals, and enum members but excludes type names', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source:
				'Private localValue As Long\n' +
				'Private Type LocalRecord\n    Value As Long\nEnd Type\n' +
				'Private Enum LocalMode\n    LocalOnly\nEnd Enum\n' +
				'Private Sub LocalProc()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Globals',
			moduleKind: 'standard',
			source:
				'Public sharedValue As Long\n' +
				'Public Sub SharedSub()\nEnd Sub\n' +
				'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n' +
				'Public Type SharedRecord\n    Value As Long\nEnd Type\n' +
				'Private hiddenValue As Long\n' +
				'Private Sub HiddenSub()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: 'Public Sub Save()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Sheet1',
			moduleKind: 'document',
			source: '',
		});

		const names = [...index.visibleNonTypeNames('Caller')].sort();
		expect(names).toEqual([
			'localonly',
			'localproc',
			'localvalue',
			'sharedonly',
			'sharedsub',
			'sharedvalue',
		]);
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

	it('preserves duplicate visible type names for ambiguity diagnostics', () => {
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

	it('carries module and type docs into visible type names', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: "''' <summary>Represents a person.</summary>\nOption Explicit\n",
		});
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: [
				"''' <summary>Shared status values.</summary>",
				'Public Enum Status',
				'    Active',
				'End Enum',
			].join('\n'),
		});

		const byName = new Map(index.visibleTypeNames('Consumer').map((t) => [t.name, t]));
		expect(byName.get('Person')?.doc?.summary).toBe('Represents a person.');
		expect(byName.get('Status')?.doc?.summary).toBe('Shared status values.');
	});

	it('resolves source locations for visible project type names', () => {
		const index = new ProjectIndex();
		const typeSource = [
			'Public Type Status',
			'    Value As Long',
			'End Type',
		].join('\n');
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: 'Option Explicit\n',
		});
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: typeSource,
		});
		index.setModule({
			moduleName: 'Consumer',
			moduleKind: 'standard',
			source: '',
		});

		const person = index.resolveTypeDefinitions('Consumer', 'Person');
		expect(person).toHaveLength(1);
		expect(person[0].moduleName).toBe('Person');
		expect(person[0].nameSpan).toEqual({ start: 0, end: 0 });

		const status = index.resolveTypeDefinitions('Consumer', 'Status');
		expect(status).toHaveLength(1);
		expect(status[0].moduleName).toBe('Types');
		expect(status[0].nameSpan).toBeDefined();
		expect(typeSource.slice(status[0].nameSpan!.start, status[0].nameSpan!.end)).toBe('Status');
	});
});

describe('ProjectIndex project class members', () => {
	it('exposes source-declared public/default-public class members', () => {
		const index = new ProjectIndex();
		const source = [
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
		].join('\n');
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source,
		});
		const person = index.projectClassMembers().find((t) => t.name === 'Person');
		expect(person?.doc).toBeUndefined();
		expect(person?.exhaustive).toBe(true);
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
		expect(
			age?.definitions?.map((def) => source.slice(def.nameSpan.start, def.nameSpan.end)),
		).toEqual(['Age', 'Age']);
		const save = person?.members.find((m) => m.name === 'Save');
		expect(save?.writable).toBeUndefined();
		expect(save?.signature).toBe('Save()');
		expect(
			save?.definitions?.map((def) => source.slice(def.nameSpan.start, def.nameSpan.end)),
		).toEqual(['Save']);
		expect(person?.members.find((m) => m.name === 'Manager')?.signature).toBe(
			'Manager() As Person',
		);
	});

	it('carries module docs into project class member surfaces', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: [
				"''' <summary>Represents a person.</summary>",
				'Option Explicit',
				'Public Sub Save()',
				'End Sub',
			].join('\n'),
		});

		const person = index.projectClassMembers().find((t) => t.name === 'Person');
		expect(person?.doc?.summary).toBe('Represents a person.');
	});

	it('exposes Event declarations on project class member surfaces', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Notifier',
			moduleKind: 'class',
			source: 'Public Event Changed(ByVal value As Long)\nPrivate Event Hidden()\n',
		});

		const notifier = index.projectClassMembers().find((t) => t.name === 'Notifier');
		expect(notifier?.members.map((member) => `${member.name}:${member.kind}:${member.signature}`))
			.toEqual(['Changed:event:Changed(value As Long)']);
	});

	it('records module-level Implements statements on project class member surfaces', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Class1',
			moduleKind: 'class',
			source: ['Implements Person', 'Implements Excel.Worksheet'].join('\n'),
		});

		const class1 = index.projectClassMembers().find((t) => t.name === 'Class1');
		expect(class1?.implements).toEqual(['Person', 'Excel.Worksheet']);
	});

	it('marks Property Get-only members as read-only and excludes public constants', () => {
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
		expect(person?.members.find((m) => m.name === 'Species')).toBeUndefined();
	});

	it('marks exported VB_UserMemId zero members as default', () => {
		const index = new ProjectIndex();
		const source = [
			'Attribute VB_Name = "Person"',
			'Public Property Get Value() As String',
			'Attribute Value.VB_UserMemId = 0',
			'End Property',
			'Public Property Let Value(ByVal value As String)',
			'End Property',
			'Public Property Get Caption() As String',
			'Attribute Caption.VB_UserMemId = -4',
			'End Property',
		].join('\n');
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source,
		});
		const person = index.projectClassMembers().find((t) => t.name === 'Person');
		const value = person?.members.find((m) => m.name === 'Value');
		const caption = person?.members.find((m) => m.name === 'Caption');
		expect(value?.defaultMember).toBe(true);
		expect(caption?.defaultMember).toBeUndefined();
		expect(
			value?.attributes?.map(
				(attr) => `${attr.targetName}.${attr.name}=${attr.valueRaw}`,
			),
		).toEqual(['Value.VB_UserMemId=0']);
		expect(
			value?.attributes?.map((attr) =>
				source.slice(attr.nameSpan.start, attr.nameSpan.end),
			),
		).toEqual(['Value.VB_UserMemId']);
	});

	it('marks document and UserForm source member surfaces as non-exhaustive', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'ThisWorkbook',
			moduleKind: 'document',
			source: 'Public Sub Hello()\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'UserForm1',
			moduleKind: 'userform',
			source: 'Public Sub ShowStatus()\nEnd Sub\n',
		});
		const members = index.projectClassMembers();
		expect(members.find((t) => t.name === 'ThisWorkbook')?.exhaustive).toBe(false);
		expect(members.find((t) => t.name === 'UserForm1')?.exhaustive).toBe(false);
	});

	it('exposes visible UDT fields as source-backed member surfaces', () => {
		const index = new ProjectIndex();
		const shared = [
			"''' <summary>Shared point payload.</summary>",
			'Public Type TPoint',
			'    X As Long',
			'    Label As String',
			'End Type',
			'Private Type THidden',
			'    Secret As String',
			'End Type',
		].join('\n');
		index.setModule({ moduleName: 'Types', moduleKind: 'standard', source: shared });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });

		const surfaces = index.projectMemberSurfaces('Caller');
		const point = surfaces.find((surface) => surface.name === 'TPoint');
		expect(point?.kind).toBe('userType');
		expect(point?.moduleName).toBe('Types');
		expect(point?.doc?.summary).toBe('Shared point payload.');
		expect(point?.exhaustive).toBe(true);
		expect(point?.members.map((member) => `${member.name}:${member.returns}`)).toEqual([
			'X:Long',
			'Label:String',
		]);
		expect(point?.members.every((member) => member.writable)).toBe(true);
		expect(
			point?.members.map((member) =>
				shared.slice(
					member.definitions?.[0]?.nameSpan.start ?? 0,
					member.definitions?.[0]?.nameSpan.end ?? 0,
				),
			),
		).toEqual(['X', 'Label']);
		expect(surfaces.find((surface) => surface.name === 'THidden')).toBeUndefined();
		expect(index.projectMemberSurfaces('Types').find((surface) => surface.name === 'THidden')).toBeDefined();
	});

	it('exposes exported standard-module members as module-qualified surfaces', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'XlideAssert',
			moduleKind: 'standard',
			source: [
				'Public Sub AreEqual(expected As Variant, actual As Variant)',
				'End Sub',
				'Sub IsTrue(condition As Boolean)',
				'End Sub',
				'Public Const DefaultTimeout As Long = 1000',
				'Public Enum SharedMode',
				'    SharedOnly',
				'End Enum',
				'Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
				'Private Sub Hidden()',
				'End Sub',
				'Private Const HiddenValue As Long = 0',
				'Private Enum HiddenMode',
				'    HiddenOnly',
				'End Enum',
			].join('\n'),
		});
		index.setModule({ moduleName: 'Tests', moduleKind: 'standard', source: '' });

		const surface = index.projectMemberSurfaces('Tests')
			.find((item) => item.name === 'XlideAssert');
		expect(surface?.kind).toBe('standardModule');
		expect(surface?.exhaustive).toBe(true);
		expect(surface?.members.map((member) => member.name)).toEqual([
			'AreEqual',
			'IsTrue',
			'DefaultTimeout',
			'SharedMode',
			'SharedOnly',
			'GetTickCount',
		]);
		expect(surface?.members.find((member) => member.name === 'AreEqual')?.signature)
			.toBe('AreEqual(expected As Variant, actual As Variant)');
		expect(surface?.members.find((member) => member.name === 'DefaultTimeout'))
			.toMatchObject({ kind: 'property', returns: 'Long', writable: false });
		expect(surface?.members.find((member) => member.name === 'SharedOnly'))
			.toMatchObject({ kind: 'property', returns: 'SharedMode', writable: false });
		expect(surface?.members.find((member) => member.name === 'GetTickCount')?.signature)
			.toBe('GetTickCount() As Long');
		expect(surface?.members.find((member) => member.name === 'Hidden')).toBeUndefined();
		expect(surface?.members.find((member) => member.name === 'HiddenValue')).toBeUndefined();
		expect(surface?.members.find((member) => member.name === 'HiddenOnly')).toBeUndefined();
		expect(index.resolveTypeDefinitions('Tests', 'XlideAssert')).toEqual([]);
	});
});

describe('ProjectIndex resolveQualifiedDefinition', () => {
	const index = new ProjectIndex();
	index.setModule({
		moduleName: 'Module1',
		moduleKind: 'standard',
		source: [
			'Public Sub DoWork()',
			'End Sub',
			'Private Sub Hidden()',
			'End Sub',
			'Public Enum SharedMode',
			'    SharedOnly',
			'End Enum',
		].join('\n'),
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

	it('resolves an exported enum member through a module qualifier', () => {
		const hits = index.resolveQualifiedDefinition('Module1', 'SharedOnly');
		expect(hits).toHaveLength(1);
		expect(hits[0].kind).toBe('enumMember');
		expect(hits[0].containerName).toBe('SharedMode');
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

	it('spans visible modules for an exported enum member', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Globals',
			moduleKind: 'standard',
			source: 'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n',
		});
		index.setModule({
			moduleName: 'Caller',
			moduleKind: 'standard',
			source: 'Sub Use()\n    value = SharedOnly\nEnd Sub\n',
		});
		index.setModule({
			moduleName: 'Shadow',
			moduleKind: 'standard',
			source: 'Private Enum LocalMode\n    SharedOnly\nEnd Enum\n',
		});

		const scope = index.referenceScope('Caller', 'SharedOnly', 0);
		expect(scope.kind).toBe('project');
		expect(scope.definitions).toHaveLength(1);
		expect(scope.definitions[0].kind).toBe('enumMember');
		expect(scope.definitions[0].moduleName).toBe('Globals');
		expect(scope.searchModules.sort()).toEqual(['Caller', 'Globals']);
		expect(scope.searchModules).not.toContain('Shadow');
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
