import { afterEach, describe, expect, it, vi } from 'vitest';
import type { BodyNode, ExprNode, ProcedureNode } from '../../src/analyzer/parser/nodes';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import * as parser from '../../src/analyzer/parser/parseExpression';
import { rawExpressionTokens, statementTokens, forEachStatementWithHeaders, walkProcedureStatements } from '../../src/analyzer/diagnostics/walker';
import { blockHeaderStatements, ifBranchStatement } from '../../src/analyzer/diagnostics/blockHeaders';
import { forEachExpressionInBody, statementValueExpressions, walkProcedureExpressions } from '../../src/analyzer/diagnostics/exprWalk';
import { DIAGNOSTIC_RULE_REGISTRY } from '../../src/analyzer/diagnostics/registry';
import * as expressionWalk from '../../src/analyzer/diagnostics/exprWalk';
import { checkStatementTypes } from '../../src/analyzer/diagnostics/rules/statementTypes';
import { checkStatementForms } from '../../src/analyzer/diagnostics/rules/statementForms';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';

afterEach(() => vi.restoreAllMocks());
function fixture(body: string) {
	const source = `Sub Main()\n${body}\nEnd Sub\n`;
	const mod = parseModule(source);
	return { source, mod, proc: mod.members[0] as ProcedureNode };
}

describe('expression traversal work', () => {
	it.each([10, 1000])('checks %i bare Case and For array values without expression parsing', (count) => {
		const source = 'Option Explicit\nSub Main()\nDim a(1) As Long, i As Long\n'
			+ Array(count).fill('Select Case 1\nCase a\nCase Is > a\nEnd Select\nFor i = a To 1\nNext').join('\n') + '\nEnd Sub\n';
		const mod = parseModule(source);
		const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule: mod });
		const parse = vi.spyOn(parser, 'parseExpression');
		const push = vi.fn();
		checkStatementTypes(source, mod, symbols, undefined, undefined, undefined, push);
		expect(push.mock.calls.map(([rule]) => rule)).toEqual(Array(count * 3).fill('nonScalarBinaryOperand'));
		expect(parse).not.toHaveBeenCalled();
	});

	// `Debug.Print 1` names no value, so the class-Sub lookup that a value
	// reference needs reads no project class's members.
	it('reads no class members for a statement callee that is no value', () => {
		const source = 'Sub Main()\nDebug.Print 1\nEnd Sub\n';
		const mod = parseModule(source);
		const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule: mod });
		let reads = 0;
		const classes = Array.from({ length: 1000 }, (_, i) => ({
			name: 'C' + i, moduleName: 'C' + i, kind: 'class' as const,
			members: Array.from({ length: 5 }, (_, k) => {
				const member: Record<string, unknown> = { name: 'M' + k, kind: 'method', moduleName: 'C' + i };
				Object.defineProperty(member, 'sub', { enumerable: true, get() { reads++; return true; } });
				return member;
			}),
		}));
		checkStatementForms(source, mod, symbols, undefined, undefined, () => undefined, { projectClassMembers: classes as never });
		expect(reads).toBe(0);
	});

	it.each([10, 1000])('indexes value membership once for %i values in a single statement', (count) => {
		const source = `Private Enum Choice\none = 1\nEnd Enum\nSub Main()\nDebug.Print ${Array(count).fill('Choice').join(', ')}\nEnd Sub\n`;
		const mod = parseModule(source);
		const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule: mod });
		const original = expressionWalk.statementValueExpressions;
		let spanReads = 0;
		const roots = vi.spyOn(expressionWalk, 'statementValueExpressions').mockImplementation((text, stmt) => {
			// Count root interval inspections: a per-candidate .some scan
			// reads quadratically even though expression parsing is cached.
			return original(text, stmt).map((expr) => {
				const span = expr.span;
				return { ...expr, get span() { spanReads++; return span; } };
			});
		});
		const push = vi.fn();
		checkStatementForms(source, mod, symbols, undefined, undefined, push);
		expect(push.mock.calls.map(([rule]) => rule)).toEqual(Array(count).fill('malformedStatement'));
		expect(roots).toHaveBeenCalledTimes(1);
		expect(spanReads).toBeLessThanOrEqual(count * 3);
	});

	it('reuses the parser-owned If and ElseIf conditions without reparsing either header', () => {
		const { source, proc } = fixture('If n Is Nothing Then\nElseIf n Is Nothing Then\nEnd If');
		const block = proc.body[0];
		if (block.kind !== 'IfBlock') { throw new Error('missing If'); }
		const parse = vi.spyOn(parser, 'parseExpression');
		const seen: ExprNode[] = [];
		forEachExpressionInBody(proc.body, undefined, (expr) => seen.push(expr), source);
		expect(parse).not.toHaveBeenCalled();
		for (const branch of block.branches) { expect(seen.filter((expr) => expr === branch.condition)).toHaveLength(1); }
	});

	it('passes the cached relative tokens directly to fallback parsing and preserves absolute spans', () => {
		const { source, proc } = fixture('Debug.Print (n Is Nothing), CStr(n + 1)');
		const stmt = proc.body[0];
		if (stmt.kind !== 'Statement') { throw new Error('missing print'); }
		const tokens = statementTokens(source, stmt.span);
		for (const token of tokens) { Object.freeze(token); }
		Object.freeze(tokens);
		const parse = vi.spyOn(parser, 'parseExpression');
		const roots = statementValueExpressions(source, stmt);
		expect(parse).toHaveBeenCalledTimes(2);
		for (const call of parse.mock.calls) { expect(call[0]).toBe(tokens); expect(call[3]).toBe(stmt.span.start); }
		expect(roots.map((expr) => source.slice(expr.span.start, expr.span.end))).toEqual(['(n Is Nothing)', 'CStr(n + 1)']);
	});

	it.each([10, 1000])('skips parsing %i plain print items but retains nested operators', (count) => {
		const { source, mod } = fixture(Array(count).fill('Debug.Print n').join('\n') + '\nDebug.Print CStr(a + 1)');
		const parse = vi.spyOn(parser, 'parseExpression');
		const visit = vi.fn();
		walkProcedureExpressions(mod, undefined, [() => visit], source, undefined, [['BinaryExpr']]);
		expect(parse).toHaveBeenCalledTimes(1);
		expect(visit).toHaveBeenCalledTimes(1);
		expect(visit.mock.calls[0][0]).toMatchObject({ exprKind: 'BinaryExpr', operator: '+' });
	});

	it.each([
		['Print #1, a + 1; CStr(n + 1), a(0)', ['a + 1', 'CStr(n + 1)', 'a(0)']],
		['Write #1, a + 1, CStr(n + 1)', ['a + 1', 'CStr(n + 1)']],
		['ReDim Preserve a(0 To n Is Nothing, F(1, 2)), b(1 To n + 1) As Long', ['0', 'n Is Nothing', 'F(1, 2)', '1', 'n + 1']],
	] as const)('parses only value roots once per statement snapshot: %s', (body, values) => {
		const { source, proc } = fixture(body);
		const stmt = proc.body[0];
		if (stmt.kind !== 'Statement') { throw new Error('missing statement'); }
		const parse = vi.spyOn(parser, 'parseExpression');
		const roots = statementValueExpressions(source, stmt);
		expect(roots.map((expr) => source.slice(expr.span.start, expr.span.end))).toEqual(values);
		for (let i = 0; i < 100; i++) { expect(statementValueExpressions(source, stmt)).toBe(roots); }
		expect(parse).toHaveBeenCalledTimes(values.length);
	});

	it.each([10, 1000])('shares one linear traversal across %i file outputs and ReDim bounds', (count) => {
		const { source, mod } = fixture(Array(count).fill('Print #1, a + 1\nReDim a(n Is Nothing)').join('\n'));
		const parse = vi.spyOn(parser, 'parseExpression');
		const scalar = vi.fn(), identity = vi.fn();
		walkProcedureExpressions(mod, undefined, [() => scalar, () => identity], source, undefined, [['BinaryExpr'], ['BinaryExpr']]);
		expect(parse).toHaveBeenCalledTimes(count * 2);
		expect(scalar).toHaveBeenCalledTimes(count * 2);
		expect(identity).toHaveBeenCalledTimes(count * 2);
		walkProcedureExpressions(mod, undefined, [() => vi.fn()], source, undefined, [['BinaryExpr']]);
		expect(parse).toHaveBeenCalledTimes(count * 2);
	});

	it.each([10, 1000])('dispatches only selected kinds across %i assignments, preserving generic visitors', (count) => {
		const { source, mod } = fixture(Array(count).fill('v = n + 1').join('\n'));
		const binary = vi.fn(), typeOf = vi.fn(), generic = vi.fn();
		walkProcedureExpressions(mod, undefined, [() => binary, () => typeOf, () => generic], source, undefined,
			[['BinaryExpr'], ['TypeOfIsExpr'], undefined]);
		expect(binary).toHaveBeenCalledTimes(count);
		expect(typeOf).not.toHaveBeenCalled();
		expect(generic).toHaveBeenCalledTimes(count * 4);
	});

	it('dispatches each header once to every opted-in rule, including the opening If', () => {
		const { source, mod, proc } = fixture('If n Then\nElseIf m Then\nx = 1\nEnd If\nFor i = 1 To 2\nx = i\nNext\nDo\nx = 2\nLoop While n');
		const takes = DIAGNOSTIC_RULE_REGISTRY.filter((rule) => rule.blockHeaders);
		const counts = takes.map(() => new Map<number, number>());
		walkProcedureStatements(mod, undefined, takes.map((_, index) => () => (stmt) => {
			const here = counts[index]; here.set(stmt.span.start, (here.get(stmt.span.start) ?? 0) + 1);
		}), undefined, { source, takes: takes.map(() => true) });
		const expected: number[] = [];
		forEachStatementWithHeaders(source, proc.body, (stmt) => expected.push(stmt.span.start));
		expect(new Set(expected).size).toBe(expected.length);
		for (let i = 0; i < takes.length; i++) {
			expect([...counts[i]], takes[i].name).toEqual(expected.map((start) => [start, 1]));
		}
	});

	it('shares synthetic header leaves across walks, invalidating a changed source snapshot', () => {
		const { source, proc } = fixture('If n Then\nEnd If\nDo While n\nLoop Until m');
		const block = proc.body[0];
		if (block.kind !== 'IfBlock') { throw new Error('missing If'); }
		const before = ifBranchStatement(source, block.branches[0]);
		expect(ifBranchStatement(source, block.branches[0])).toBe(before);
		const loop = proc.body[1] as BodyNode;
		const headers = blockHeaderStatements(source, loop);
		expect(blockHeaderStatements(source, loop)).toBe(headers);
		const fresh = Buffer.from(source).toString('utf8');
		expect(ifBranchStatement(fresh, block.branches[0])).toBe(before);
		expect(blockHeaderStatements(fresh, loop)).toBe(headers);
		const changed = source.replaceAll(' n', ' q');
		expect(ifBranchStatement(changed, block.branches[0])).not.toBe(before);
		expect(ifBranchStatement(changed, block.branches[0]).raw).toBe('If q Then');
		expect(blockHeaderStatements(changed, loop)).not.toBe(headers);
	});
});

describe('relative expression tokens with a span offset', () => {
	it.each([
		'-((a + 1) * 2)', 'TypeOf obj.Member(1) Is Worksheet', 'rs!Field', 'x$',
		'F(first:=a, , last:=g(2))', 'AddressOf Callback', 'New Collection', '.Value',
		'CStr(1 Is 1)', 'a ^ -2', '#1/1/2000#', 'a +', 'f(', 'TypeOf Is Object',
	])('matches absolute-token parsing, including diagnostics: %s', (expression) => {
		const tokens = rawExpressionTokens(expression);
		const offset = 73;
		const absolute = tokens.map((tok) => ({ ...tok, start: tok.start + offset, end: tok.end + offset }));
		const expected = parser.parseExpression(absolute);
		for (const token of tokens) { Object.freeze(token); }
		Object.freeze(tokens);
		expect(parser.parseExpression(tokens, 0, tokens.length, offset)).toEqual(expected);
	});
	it('offsets positional, named and omitted arguments without changing token boundaries', () => {
		const tokens = rawExpressionTokens('first:=a, , last:=obj!Field');
		const absolute = tokens.map((tok) => ({ ...tok, start: tok.start + 73, end: tok.end + 73 }));
		expect(parser.parseParenlessArguments(tokens, 0, tokens.length, 73))
			.toEqual(parser.parseParenlessArguments(absolute, 0, absolute.length));
	});
});
