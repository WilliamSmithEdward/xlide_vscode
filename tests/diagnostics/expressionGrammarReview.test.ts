import { describe, expect, it, vi } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { forEachExpressionInBody, statementValueExpressions } from '../../src/analyzer/diagnostics/exprWalk';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { tokenize } from '../../src/analyzer/lexer/tokenize';
import { parseExpression } from '../../src/analyzer/parser/parseExpression';
import * as expressions from '../../src/analyzer/parser/parseExpression';
import { analyzeProjectModule } from './helpers';

// Grammar recovery / analyzer regressions, not new Excel message or selection evidence.
function roots(source: string) {
	return statementValueExpressions(source, { kind: 'Statement', raw: source, span: { start: 0, end: source.length } });
}
const wrap = (body: string) => `Option Explicit\nSub Main()\n${body}\nEnd Sub\n`;

describe('expression grammar review recovery', () => {
	it.each(['VBA.Left$("a", 1)', 'VBA.Time$', 'VBA.Strings.Left$("a", 1)', '.Left$("a", 1)'])('preserves suffixed member %s and exact spans', (value) => {
		const source = `${value} Is Nothing`;
		const tokens = tokenize(source);
		const result = parseExpression(tokens);
		expect(result.diagnostics).toEqual([]);
		expect(result.endIndex).toBe(tokens.length);
		expect(result.expr?.exprKind).toBe('BinaryExpr');
		if (result.expr?.exprKind !== 'BinaryExpr') { throw new Error('missing Is'); }
		const left = result.expr.left;
		const member = left.exprKind === 'IndexExpr' ? left.callee : left;
		expect(member.exprKind).toBe('MemberAccessExpr');
		if (member.exprKind !== 'MemberAccessExpr') { throw new Error('missing member'); }
		expect(member.member).toMatch(/\$$/);
		expect(source.slice(member.memberSpan.start, member.memberSpan.end)).toBe(member.member);
		expect(source.slice(left.span.start, left.span.end)).toBe(value);
		expect(roots(`Debug.Print ${source}`)).toHaveLength(1);
	});

	it.each(['result = EXPR', 'Take EXPR', 'Debug.Print EXPR', 'If EXPR Then Exit Sub'])('retains suffix expression spans in %s', (statement) => {
		for (const value of ['VBA.Left$("a", 1)', 'VBA.Time$', 'VBA.Strings.Left$("a", 1)']) {
			const source = wrap(statement.replace('EXPR', `${value} Is Nothing`));
			const proc = parseModule(source).members.find((member) => member.kind === 'Procedure');
			if (proc?.kind !== 'Procedure') { throw new Error('missing procedure'); }
			const found: string[] = [];
			forEachExpressionInBody(proc.body, undefined, (expr) => {
				if (expr.exprKind === 'BinaryExpr' && expr.operator === 'Is') { found.push(source.slice(expr.left.span.start, expr.left.span.end)); }
			}, source);
			expect(found).toEqual([value]);
		}
	});

	it.each(['1', '-1', '1 + 2', '"x"', 'True', 'Nothing', 'Empty', 'Null'])('recovers TypeOf %s without asserting an unmeasured literal error', (operand) => {
		const source = `TypeOf ${operand} Is Collection`;
		const tokens = tokenize(source);
		const result = parseExpression(tokens);
		expect(result.endIndex).toBe(tokens.length);
		expect(result.diagnostics).toEqual([]);
		expect(result.expr?.exprKind).toBe('TypeOfIsExpr');
		expect(analyzeModule(wrap(`Debug.Print ${source}`)).filter((d) => d.code === 'typeof-invalid-expression')).toEqual([]);
	});

	it.each([
		['Print #(a + 1), Tab(b + 2); Spc(c + 3); f(1, 2)', ['(a + 1)', 'b + 2', 'c + 3', 'f(1, 2)']],
		['Write #(a + 1), f(1, 2), n Is Nothing', ['(a + 1)', 'f(1, 2)', 'n Is Nothing']],
		['Seek #f(1, 2), a + 1', ['f(1, 2)', 'a + 1']],
		['Open f("x", 1) For Random Access Read Write Lock Read As #(a + 1) Len = b + 2', ['f("x", 1)', '(a + 1)', 'b + 2']],
		['Put #1, a + 1, v', ['1', 'a + 1', 'v']],
		['Put #1, , v', ['1', 'v']],
		['Get #1, a + 1, v(b + 2)', ['1', 'a + 1', 'b + 2']],
		['Get #1, , v', ['1']],
		['Line Input #(a + 1), s', ['(a + 1)']],
		['Input #f(1, 2), s, v(a + 1)', ['f(1, 2)', 'a + 1']],
		['Debug.Print Tab(a + 1); Spc(b + 2)', ['a + 1', 'b + 2']],
	] as const)('extracts grammar slots from %s', (source, expected) => {
		expect(roots(source).map((root) => source.slice(root.span.start, root.span.end))).toEqual(expected);
	});

	it('parses each file slot once with linear total expression input', () => {
		const source = 'Print #1, ' + Array.from({ length: 200 }, (_, i) => `Tab(f(${i}, 1) + 1)`).join('; ');
		const spy = vi.spyOn(expressions, 'parseExpression');
		try {
			const parsed = roots(source);
			expect(parsed).toHaveLength(201);
			expect(spy).toHaveBeenCalledTimes(201);
			const visitedTokens = spy.mock.calls.reduce((sum, [tokens, from = 0, to = tokens.length]) => sum + to - from, 0);
			expect(visitedTokens).toBeLessThan(tokenize(source).length);
			expect(roots(source)).toBe(parsed);
			expect(spy).toHaveBeenCalledTimes(201);
		} finally { spy.mockRestore(); }
	});

	it('keeps root identity for an unchanged snapshot and changes it for changed source', () => {
		const source = 'Print #1, Tab(a + 1); 2';
		const first = roots(source);
		for (let i = 0; i < 100; i++) { expect(roots(source)).toBe(first); }
		expect(roots(source.replace('a', 'b'))).not.toBe(first);
	});
});

describe('imported Enum operand identity', () => {
	const imported = { moduleName: 'Types', source: 'Public Enum E\nFirst = 1\nEnd Enum\n' };
	it.each(['E', 'Types.E'])('classifies imported %s in Is and TypeOf', (type) => {
		const source = wrap(`Dim n As ${type}\nDebug.Print n Is Nothing\nDebug.Print TypeOf n Is Collection`);
		const hits = analyzeProjectModule(source, [imported], 'MainModule');
		expect(hits.filter((d) => d.code === 'is-operator-non-object')).toHaveLength(1);
		expect(hits.filter((d) => d.code === 'typeof-invalid-expression')).toHaveLength(1);
	});
	it('preserves same-module Enum variables and enum members', () => {
		const source = imported.source + wrap('Dim n As E\nDebug.Print n Is Nothing, First Is Nothing\nDebug.Print TypeOf n Is Collection');
		const hits = analyzeProjectModule(source, [], 'MainModule');
		expect(hits.filter((d) => d.code === 'is-operator-non-object')).toHaveLength(2);
		expect(hits.filter((d) => d.code === 'typeof-invalid-expression')).toHaveLength(1);
	});

	it('does not choose an ambiguous imported Enum', () => {
		const source = wrap('Dim n As E\nDebug.Print n Is Nothing\nDebug.Print TypeOf n Is Collection');
		const hits = analyzeProjectModule(source, [imported, { ...imported, moduleName: 'Other' }], 'MainModule');
		expect(hits.filter((d) => ['is-operator-non-object', 'typeof-invalid-expression'].includes(d.code))).toEqual([]);
	});
	it('uses module qualification to disambiguate imported Enums', () => {
		const source = wrap('Dim n As Types.E\nDebug.Print n Is Nothing\nDebug.Print TypeOf n Is Collection');
		const hits = analyzeProjectModule(source, [imported, { ...imported, moduleName: 'Other' }], 'MainModule');
		expect(hits.filter((d) => ['is-operator-non-object', 'typeof-invalid-expression'].includes(d.code))).toHaveLength(2);
	});
	it('keeps a local UDT type binding ahead of an imported Enum', () => {
		const source = 'Private Type E\nvalue As Long\nEnd Type\n' + wrap('Dim n As E\nDebug.Print n Is Nothing\nDebug.Print TypeOf n Is Collection');
		const hits = analyzeProjectModule(source, [imported], 'MainModule');
		expect(hits.filter((d) => ['is-operator-non-object', 'typeof-invalid-expression'].includes(d.code))).toEqual([]);
	});

	it('retains a local Object value shadow over an imported Enum member', () => {
		const source = wrap('Dim First As Object\nDebug.Print First Is Nothing');
		expect(analyzeProjectModule(source, [imported], 'MainModule').filter((d) => d.code === 'is-operator-non-object')).toEqual([]);
	});
});
