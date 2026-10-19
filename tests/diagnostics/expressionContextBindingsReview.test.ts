import { describe, expect, it, vi } from 'vitest';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { parseExpression } from '../../src/analyzer/parser/parseExpression';
import { rawExpressionTokens } from '../../src/analyzer/diagnostics/walker';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { expressionShapeResolver, type ExpressionShapeContext } from '../../src/analyzer/diagnostics/expressionShape';
import * as runtime from '../../src/analyzer/runtime/vbaRuntime';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';
import type { ProcedureNode } from '../../src/analyzer/parser/nodes';

// Analyzer binding contracts, pending the parent's Excel compile measurements.
// These expectations do not claim measured Office diagnostics or selections.
function resolver(body = '', declarations = '', context?: ExpressionShapeContext, parameters = '') {
	const source = `${declarations}\nSub Main(${parameters})\n${body}\nEnd Sub\n`;
	const mod = parseModule(source);
	const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule: mod });
	const proc = mod.members.find((node): node is ProcedureNode => node.kind === 'Procedure' && node.name === 'Main')!;
	const factory = expressionShapeResolver(symbols, undefined, context);
	const shape = factory(proc);
	return { symbols, proc, factory, shape: (text: string) => shape(parseExpression(rawExpressionTokens(text)).expr!) };
}
const designer = (name: string): ExpressionShapeContext => ({ implicitMembers: [{ name, type: 'MSForms.ComboBox' }] });

describe('expression binding review contracts (Excel measurement pending)', () => {
	it.each(['Timer', 'Rnd', 'Len', 'InputBox'])('preserves designer binding for %s before runtime fallback', (name) => {
		const { shape } = resolver('', '', designer(name));
		expect(shape(name)).toBeUndefined();
		expect(shape(`${name}(1)`)).toBeUndefined();
		expect(shape('VBA.Timer')?.asType).toBe('Single');
	});
	it.each(['Picker = 1', 'Picker.Value = 1', 'Me.Picker = 1'])('excludes assigned designer names from DefType: %s', (assignment) => {
		const { shape } = resolver(assignment, 'DefLng A-Z', designer('Picker'));
		expect(shape('Picker')).toBeUndefined();
		expect(shape('Unbound')?.asType).toBe('Long');
	});
	it.each(['Long', 'Object', 'Variant'])('preserves declared and parameter %s shadows', (type) => {
		for (const name of ['Picker', 'Timer', 'VBA']) {
			expect(resolver(`Dim ${name} As ${type}`, 'DefLng A-Z', designer(name)).shape(name)?.asType).toBe(type);
			expect(resolver('', 'DefLng A-Z', designer(name), `${name} As ${type}`).shape(name)?.asType).toBe(type);
		}
		expect(resolver(`Dim VBA As ${type}`).shape('VBA.Timer')).toBeUndefined();
	});
	it.each([
		['class', true], ['class', false], ['class', undefined], ['document', undefined], ['userform', undefined],
	] as const)('does not invent an implicit local or an object for %s / %s', (kind, predeclaredId) => {
		const { shape } = resolver('Picker = 1', 'DefLng A-Z', { projectClassMembers: [{ name: 'Picker', moduleName: 'Picker', kind, predeclaredId, members: [] }] });
		expect(shape('Picker')).toBeUndefined();
	});
	it.each(['class', 'document', 'userform'] as const)('preserves usable %s instance binding before intrinsic Timer', (kind) => {
		const { shape } = resolver('', '', { projectClassMembers: [{ name: 'Timer', moduleName: 'Timer', kind, predeclaredId: true, members: [] }] });
		expect(shape('Timer')).toBeUndefined();
	});
	it('keeps unread class predeclared metadata conservative on a runtime collision', () => {
		expect(resolver('', '', { projectClassMembers: [{ name: 'Timer', moduleName: 'Timer', kind: 'class', members: [] }] }).shape('Timer')).toBeUndefined();
	});
	it('does not promote an ordinary class into a usable predeclared instance', () => {
		const context: ExpressionShapeContext = { projectClassMembers: [{ name: 'Timer', moduleName: 'Timer', kind: 'class', predeclaredId: false, members: [] }] };
		expect(resolver('', '', context).shape('Timer')?.asType).toBe('Single');
	});
	it.each(['', 'VBA.'])('validates runtime slots through %s qualification', (prefix) => {
		const { shape } = resolver();
		for (const call of ['CLng()', 'CLng(1, 2)', 'CLng(missing:=1)', 'CLng(,)', 'CLng(expression:=1, expression:=2)']) {
			expect(shape(prefix + call), call).toBeUndefined();
		}
		expect(shape(prefix + 'CLng(1)')?.asType).toBe('Long');
		expect(shape(prefix + 'CLng(expression:=1)')?.asType).toBe('Long');
	});
	it('does not interpret unavailable runtime arity as zero arguments', () => {
		const original = runtime.resolveRuntimeFunction;
		const spy = vi.spyOn(runtime, 'resolveRuntimeFunction').mockImplementation((name) => name.toLowerCase() === 'mystery'
			? { ...original('CLng')!, name: 'Mystery', signature: 'Mystery', params: undefined } : original(name));
		try {
			const { shape } = resolver();
			expect(shape('Mystery(1, 2)')?.asType).toBe('Long');
			expect(shape('VBA.Mystery(1, 2)')).toBeUndefined();
		} finally { spy.mockRestore(); }
	});
	it.each([
		['Function Foo()', 'Long'], ['Function Foo() As Variant', 'Variant'],
		['Function Foo() As Long', 'Long'], ['Function Foo$()', 'String'],
	])('uses effective function return for %s', (header, type) => {
		const { shape } = resolver('', `DefLng A-Z\n${header}\nEnd Function`);
		for (const call of ['Foo', 'Foo()', 'Module1.Foo', 'Module1.Foo()']) { expect(shape(call)?.asType, call).toBe(type); }
	});
	it('keeps imported unknown returns out of the caller DefType fallback', () => {
		const { symbols, proc } = resolver('', 'DefLng A-Z');
		const foreign = buildModuleSymbols('Other', 'standard', 'Public Function Foo()\nEnd Function\n');
		const shape = expressionShapeResolver(symbols, foreign.root.children)(proc);
		expect(shape(parseExpression(rawExpressionTokens('Foo')).expr!)).toBeUndefined();
	});
	it('distinguishes function arguments from subsequent array result indexing (#609)', () => {
		const { shape } = resolver('', 'Function Values(n As Long) As Long()\nEnd Function');
		expect(shape('Values(1)')?.isArray).toBe(true);
		// Subsequent indexing is left to runtime; do not promote it to a known scalar.
		expect(shape('Values(1)(0)')).toMatchObject({ isArray: false });
		expect(shape('Values()')).toBeUndefined();
	});
	it('indexes designer and project names once per factory across repeated operands', () => {
		let reads = 0;
		const context: ExpressionShapeContext = {
			implicitMembers: [{ get name() { reads++; return 'Picker'; }, type: 'MSForms.ComboBox' }],
			projectClassMembers: [{ get name() { reads++; return 'Factory'; }, moduleName: 'Factory', kind: 'class', predeclaredId: true, members: [] }],
		};
		const { shape, factory, symbols, proc } = resolver('Picker = 1', 'DefLng A-Z', context);
		shape('Picker');
		const initial = reads;
		expect(initial).toBeGreaterThan(0);
		for (let i = 0; i < 100; i++) {
			expect(expressionShapeResolver(symbols, undefined, context)).toBe(factory);
			factory(proc);
			shape('Picker'); shape('Factory'); shape('Unknown');
		}
		expect(reads).toBe(initial);
	});
	it.each(['Timer', 'Rnd', 'Len', 'InputBox', 'Picker'])('keeps designer %s conservative through the diagnostic pipeline', (name) => {
		const source = `DefLng A-Z\nSub Main()\n${name} = 1\nDebug.Print ${name} Is Nothing\nIf TypeOf ${name} Is Object Then Exit Sub\nEnd Sub`;
		const project = buildVbaProjectIndex([{ moduleName: 'Form1', type: 'userform', source, implicitMembers: [{ name, type: 'MSForms.ComboBox' }] }]);
		const diagnostics = analyzeVbaModuleSource({ source, moduleName: 'Form1', moduleKind: 'userform', host: 'excel',
			...projectAnalysisOptionsForModule(project, 'Form1', projectProcedureSignatures(project)),
		}).diagnostics;
		expect(diagnostics.filter((d) => ['is-operator-non-object', 'typeof-invalid-expression'].includes(d.code ?? ''))).toEqual([]);
	});
	it.each(['CLng()', 'VBA.CLng()', 'CLng(1, 2)', 'VBA.CLng(1, 2)'])('avoids secondary Is/TypeOf errors for invalid %s', (call) => {
		for (const body of [`Debug.Print ${call} Is Nothing`, `If TypeOf ${call} Is Object Then Exit Sub`]) {
			const diagnostics = analyzeVbaModuleSource({ source: `Option Explicit\nSub Main()\n${body}\nEnd Sub`, moduleName: 'Module1', moduleKind: 'standard', host: 'excel' }).diagnostics;
			expect(diagnostics.filter((d) => ['is-operator-non-object', 'typeof-invalid-expression'].includes(d.code ?? ''))).toEqual([]);
		}
	});
});
