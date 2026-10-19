import { describe, expect, it, vi } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { parseExpression } from '../../src/analyzer/parser/parseExpression';
import { rawExpressionTokens } from '../../src/analyzer/diagnostics/walker';
import { expressionShapeResolver, type ExpressionShapeContext } from '../../src/analyzer/diagnostics/expressionShape';
import { buildVbaProjectIndex, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';
import { checkArgumentCount } from '../../src/analyzer/diagnostics/rules/callArity';
import { callableTypeSignaturesFor, sameModuleCallableSignatures, uniqueProjectTypeSignatures, moduleSignatureFromSymbol } from '../../src/analyzer/diagnostics/typeInference';
import type { MemberCompletionContext } from '../../src/analyzer/completion/memberAccess';
import type { VbaProcedureSignature } from '../../src/analyzer/symbols/symbolModel';
import type { ProcedureNode, LeafStatementNode } from '../../src/analyzer/parser/nodes';

function diagnostics(body: string, declarations = '', parameters = '') {
	return analyzeVbaModuleSource({ source: `Option Explicit\n${declarations}\nSub Main(${parameters})\n${body}\nEnd Sub`, moduleName: 'Module1', moduleKind: 'standard', host: 'excel' }).diagnostics.filter((d) => d.severity === 'error');
}
function resolver(declarations = '', body = '', context?: ExpressionShapeContext) {
	const source = `${declarations}\nSub Main()\n${body}\nEnd Sub`;
	const parsedModule = parseModule(source);
	const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule });
	const proc = parsedModule.members.find((p): p is ProcedureNode => p.kind === 'Procedure' && p.name === 'Main')!;
	const factory = expressionShapeResolver(symbols, undefined, context);
	const shape = factory(proc);
	return { symbols, factory, shape: (text: string) => {
		const expr = parseExpression(rawExpressionTokens(text)).expr!;
		return { expr, shape: shape(expr) };
	} };
}
const secondary = ['is-operator-non-object', 'typeof-invalid-expression'];
const contexts = (call: string) => [`Debug.Print ${call}`, `Debug.Print ${call} Is Nothing`, `If TypeOf ${call} Is Collection Then Exit Sub`];

describe('qualified call context review', () => {
	// CLng() message is measured in Excel 16.0.20430.20182, 2026-10-19.
	// Analyzer spans are not asserted to match post-dialog VBE selections.
	it.each(contexts('VBA.CLng()'))('reports the measured primary error without a secondary: %s', (body) => {
		const hits = diagnostics(body);
		const primary = hits.filter((d) => d.code === 'argument-count');
		expect(primary).toHaveLength(1);
		expect(primary[0].message).toContain('Argument not optional');
		expect(primary[0].message).not.toContain('Wrong number of arguments');
		const source = `Option Explicit\n\nSub Main()\n${body}\nEnd Sub`;
		expect(source.slice(primary[0].span.start, primary[0].span.end)).toBe('.CLng');
		expect(hits.filter((d) => secondary.includes(d.code ?? ''))).toEqual([]);
	});
	it.each(contexts('CLng()'))('preserves the measured bare intrinsic syntax error: %s', (body) => {
		const hits = diagnostics(body);
		const primary = hits.filter((d) => d.code === 'argument-count');
		expect(primary).toHaveLength(1);
		expect(primary[0].message).toContain('Syntax error');
		expect(hits.filter((d) => secondary.includes(d.code ?? ''))).toEqual([]);
	});

	it.each([
		['CInt()', 'Syntax error'], ['Len()', 'Syntax error'], ['CLng(1, 2)', 'Syntax error'],
		['VBA.CInt()', 'Argument not optional'], ['VBA.Len()', 'Argument not optional'],
		['Left$()', 'Argument not optional'], ['VBA.Left$()', 'Argument not optional'],
		['VBA.CLng(1, 2)', 'Wrong number of arguments or invalid property assignment'],
	])('preserves the measured neighboring arity message for %s', (call, message) => {
		for (const body of contexts(call)) {
			const hits = diagnostics(body);
			const primary = hits.filter((d) => d.code === 'argument-count');
			expect(primary, body).toHaveLength(1);
			expect(primary[0].message, body).toContain(message);
			expect(hits.filter((d) => secondary.includes(d.code ?? '')), body).toEqual([]);
		}
	});

	// The remaining invocation contracts are analyzer regressions, not new
	// claims about exact Excel messages, declarations or selections.
	it.each(['VBA.CLng(1, 2)', 'Module1.Foo()', 'Module1.Foo(1, 2)'])('owns invalid %s with a primary argument diagnostic', (call) => {
		for (const body of contexts(call)) {
			const hits = diagnostics(body, 'Function Foo(n As Long) As Long\nEnd Function');
			expect(hits.filter((d) => d.code === 'argument-count'), `${body}: ${JSON.stringify(hits)}`).toHaveLength(1);
			expect(hits.filter((d) => secondary.includes(d.code ?? '')), body).toEqual([]);
		}
	});
	it.each(['VBA.CLng(1)', 'VBA.CLng(expression:=1)', 'Module1.Foo(1)'])('retains valid neighboring %s', (call) => {
		expect(diagnostics(`Debug.Print ${call}`, 'Function Foo(n As Long) As Long\nEnd Function')).toEqual([]);
		expect(diagnostics(`Debug.Print ${call} Is Nothing`, 'Function Foo(n As Long) As Long\nEnd Function').some((d) => d.code === 'is-operator-non-object')).toBe(true);
	});
	it.each(['VBA', 'Module1'])('respects local, parameter and module variable %s shadows', (name) => {
		const call = `${name}.${name === 'VBA' ? 'CLng' : 'Foo'}()`;
		for (const hits of [
			diagnostics(`Dim ${name} As Object\nDebug.Print ${call}`, 'Function Foo(n As Long) As Long\nEnd Function'),
			diagnostics(`Debug.Print ${call}`, 'Function Foo(n As Long) As Long\nEnd Function', `${name} As Object`),
			diagnostics(`Debug.Print ${call}`, `Dim ${name} As Object\nFunction Foo(n As Long) As Long\nEnd Function`),
		]) { expect(hits.filter((d) => d.code === 'argument-count')).toEqual([]); }
	});
	it.each(['obj.VBA.CLng()', 'obj.Module1.Foo()', 'VBA.Strings.CLng()'])('does not treat an unproven chain as a runtime/module call: %s', (call) => {
		expect(diagnostics(`Dim obj As Object\nDebug.Print ${call}`, 'Function Foo(n As Long) As Long\nEnd Function').filter((d) => d.code === 'argument-count')).toEqual([]);
	});
	it('normalizes source suffix lookup without changing member spelling or span', () => {
		const { shape } = resolver('DefLng A-Z\nFunction Foo$()\nEnd Function');
		for (const text of ['Foo$()', 'Module1.Foo$()', 'Module1.Foo$']) {
			expect(shape(text).shape?.asType, text).toBe('String');
		}
		const { expr } = shape('Module1.Foo$()');
		expect(expr.exprKind).toBe('IndexExpr');
		if (expr.exprKind === 'IndexExpr' && expr.callee.exprKind === 'MemberAccessExpr') {
			expect(expr.callee.member).toBe('Foo$');
			expect(expr.callee.memberSpan).toEqual({ start: 8, end: 12 });
		}
		expect(diagnostics('Debug.Print Module1.Foo$() Is Nothing', 'Function Foo$()\nEnd Function').some((d) => d.code === 'is-operator-non-object')).toBe(true);
	});
	it('preserves subsequent array indexing and invalid initial invocations (#609)', () => {
		const { shape } = resolver('Function Numbers(n As Long) As Long()\nEnd Function');
		for (const prefix of ['', 'Module1.']) {
			expect(shape(`${prefix}Numbers(1)`).shape?.isArray).toBe(true);
			expect(shape(`${prefix}Numbers(1)(0)`).shape?.isArray).toBe(false);
			expect(shape(`${prefix}Numbers()`).shape).toBeUndefined();
			expect(diagnostics(`Debug.Print ${prefix}Numbers(1)(0)`, 'Function Numbers(n As Long) As Long()\nEnd Function').filter((d) => d.code === 'argument-count')).toEqual([]);
		}
	});
	it('uses exported signatures from the declaring module, guarding ambiguity, arity and local shadows', () => {
		const project = buildVbaProjectIndex([
			{ moduleName: 'Other', type: 'standard', source: 'DefLng A-Z\nPublic Function Foo(n As Long)\nEnd Function' },
			{ moduleName: 'Module1', type: 'standard', source: 'DefStr A-Z\nSub Main()\nEnd Sub' },
		]);
		// The parent owns exporting effective signatures from ProjectIndex and
		// wiring them into analysis. Exercise this resolver's input contract.
		const declaring = buildModuleSymbols('Other', 'standard', 'DefLng A-Z\nPublic Function Foo(n As Long)\nEnd Function');
		const effective = moduleSignatureFromSymbol(declaring, declaring.root.children!.find((s) => s.name === 'Foo')!);
		const projectProcedures = new Map([...projectProcedureSignatures(project)!].map(([key, signatures]) => [key, signatures.map((sig) => ({ ...sig, returnType: effective.returnType }))]));
		const context = { projectProcedures };
		const { shape, symbols, factory } = resolver('DefStr A-Z', '', context);
		expect(expressionShapeResolver(symbols, undefined, context)).toBe(factory);
		for (const call of ['Foo(1)', 'Other.Foo(1)']) { expect(shape(call).shape?.asType, call).toBe('Long'); }
		for (const call of ['Foo()', 'Other.Foo()', 'Foo(1, 2)', 'Other.Foo(1, 2)']) { expect(shape(call).shape, call).toBeUndefined(); }
		const shadow = resolver('', 'Dim Foo As Object\nDim Other As Object', context);
		for (const call of ['Foo(1)', 'Other.Foo(1)']) { expect(shadow.shape(call).shape, call).toBeUndefined(); }
		const ambiguous = new Map(projectProcedures);
		ambiguous.set('foo', [...projectProcedures.get('foo')!, ...projectProcedures.get('foo')!]);
		expect(resolver('', '', { projectProcedures: ambiguous }).shape('Foo(1)').shape).toBeUndefined();
		expect(resolver('', '', { projectProcedures: ambiguous }).shape('Other.Foo(1)').shape?.asType).toBe('Long');
	});
	// All nine combinations compiled in the parent's 2026-10-19 Excel batch.
	it.each(['Abs', 'Int', 'Fix'])('keeps measured %s Null/Variant/Object calls clear and numeric constraints intact', (name) => {
		for (const argument of ['Null', 'v', 'obj']) {
			expect(diagnostics(`Dim v As Variant\nDim obj As Object\nDebug.Print ${name}(${argument}) Is Nothing`)).toEqual([]);
		}
		for (const argument of ['1', '-1', 'n', '"1"', 's']) {
			expect(diagnostics(`Dim n As Long\nDim s As String\nDebug.Print ${name}(${argument}) Is Nothing`).some((d) => d.code === 'is-operator-non-object'), argument).toBe(true);
		}
	});
	it.each(['MsgBox("x")', 'VarType(1)'])('preserves the unconditional enum constraint for %s', (call) => {
		expect(diagnostics(`Debug.Print ${call} Is Nothing`).some((d) => d.code === 'is-operator-non-object')).toBe(true);
	});
	it.each(['Strings.Left$("a", 1)', 'VBA.Strings.Left$("a", 1)'])('resolves measured library membership for %s', (call) => {
		expect(diagnostics(`Debug.Print ${call}`)).toEqual([]);
		for (const body of contexts(call).slice(1)) {
			expect(diagnostics(body).some((d) => secondary.includes(d.code ?? ''))).toBe(true);
		}
	});
	it.each(['Strings.Left$()', 'VBA.Strings.Left$()'])('reports primary arity without secondary for %s', (call) => {
		for (const body of contexts(call)) {
			const hits = diagnostics(body);
			expect(hits.filter((d) => d.code === 'argument-count'), `${body}: ${JSON.stringify(hits)}`).toHaveLength(1);
			expect(hits.filter((d) => secondary.includes(d.code ?? '')), body).toEqual([]);
		}
	});
	it.each(['VBA.Math.Left$("a", 1)', 'Math.Left$("a", 1)', 'obj.Strings.Left$("a", 1)', 'obj.VBA.Strings.Left$("a", 1)'])('leaves wrong containers and receiver chains unresolved: %s', (call) => {
		expect(resolver('', 'Dim obj As Object').shape(call).shape).toBeUndefined();
		expect(diagnostics(`Dim obj As Object\nDebug.Print ${call} Is Nothing`).filter((d) => secondary.includes(d.code ?? ''))).toEqual([]);
	});
	it('keeps qualified library roots and bare containers subject to the right shadow', () => {
		for (const name of ['VBA', 'Strings']) {
			const { shape } = resolver('', `Dim ${name} As Object`);
			expect(shape(`${name === 'VBA' ? 'VBA.Strings' : 'Strings'}.Left$("a", 1)`).shape).toBeUndefined();
		}
		expect(resolver('', 'Dim Strings As Object').shape('VBA.Strings.Left$("a", 1)').shape?.asType).toBe('String');
		for (const name of ['VBA', 'Strings']) {
			const context = { implicitMembers: [{ name, type: 'MSForms.ComboBox' }] };
			expect(resolver('', '', context).shape(`${name === 'VBA' ? 'VBA.Strings' : 'Strings'}.Left$("a", 1)`).shape).toBeUndefined();
		}
	});
	it.each(['Strings', 'VBA'])('does not borrow runtime signatures for members of source module %s', (moduleName) => {
		const source = `Option Explicit\nSub Main()\nDebug.Print ${moduleName}.${moduleName === 'Strings' ? 'Left$' : 'CLng'}()\nEnd Sub`;
		const hits = analyzeVbaModuleSource({ source, moduleName, moduleKind: 'standard', host: 'excel' }).diagnostics;
		expect(hits.filter((d) => d.code === 'argument-count')).toEqual([]);
	});

	it('does no qualified-table or class work for unrelated statements; reuses tables and invalidates changed identities', () => {
		const source = 'Sub Main()\nDebug.Print 1\nDebug.Print VBA.CLng()\nEnd Sub';
		const parsedModule = parseModule(source);
		const proc = parsedModule.members[0] as ProcedureNode;
		const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule });
		const project = new Map<string, VbaProcedureSignature[]>(Array.from({ length: 1000 }, (_, i) => [`other.f${i}`, [{ name: `F${i}`, moduleName: 'Other', kind: 'function', params: [], returnType: 'Long' }]]));
		let copied = 0;
		const instrument = (map: ReadonlyMap<string, unknown>) => {
			const original = map[Symbol.iterator].bind(map);
			return vi.spyOn(map, Symbol.iterator).mockImplementation(function* () { for (const entry of original()) { copied++; yield entry; } });
		};
		const spies = [instrument(callableTypeSignaturesFor(symbols, project)), instrument(sameModuleCallableSignatures(symbols)), instrument(uniqueProjectTypeSignatures(project))];
		let kinds = 0;
		const classes: NonNullable<MemberCompletionContext['projectClassMembers']> = Array.from({ length: 1000 }, (_, i) => ({ name: `Class${i}`, moduleName: `Class${i}`, get kind() { kinds++; return 'class' as const; }, predeclaredId: false, members: [] }));
		const context: MemberCompletionContext = { parsedModule, projectClassMembers: classes };
		const run = (qualified: boolean, ctx = context, signatures = project) => {
			const visitor = checkArgumentCount(source, symbols, signatures, undefined, ctx, () => {});
			visitor(proc)!((proc.body.filter((node) => node.kind === 'Statement')[qualified ? 1 : 0]) as LeafStatementNode);
		};
		try {
			run(false); run(false);
			expect({ copied, kinds }).toEqual({ copied: 0, kinds: 0 });
			run(true);
			expect(copied).toBeGreaterThan(1000); expect(kinds).toBe(1000);
			const first = { copied, kinds };
			run(true, { ...context });
			expect({ copied, kinds }).toEqual(first);
			run(true, { ...context, projectClassMembers: [...classes] });
			expect(kinds).toBeGreaterThan(first.kinds);
			const changed = new Map(project);
			const nextMap = callableTypeSignaturesFor(symbols, changed);
			const spy = instrument(nextMap);
			try { const before = copied; run(true, context, changed); expect(copied).toBeGreaterThan(before); } finally { spy.mockRestore(); }
		} finally { for (const spy of spies) { spy.mockRestore(); } }
	});

});
