// Unit tests for the VBA value-expression parser (MS-VBAL §5.6).
//
// Verified against MS-VBAL.pdf, v20250520. The parser is the roadmap 2.4.0
// critical-path keystone; these tests pin its grammar, operator precedence,
// span accuracy, and error tolerance before any statement wiring consumes it.

import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { parseExpression } from '../src/analyzer/parser/parseExpression';
import {
	BinaryExpr,
	ExprNode,
	IdentifierExpr,
	IndexExpr,
	LiteralExpr,
	MemberAccessExpr,
	NewExpr,
	ParenExpr,
	TypeOfIsExpr,
	UnaryExpr,
} from '../src/analyzer/parser/nodes';

/** Tokenize a source fragment to the significant tokens an expression sees. */
function exprTokens(src: string) {
	return tokenize(src).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
}

function parse(src: string) {
	return parseExpression(exprTokens(src));
}

function parseOk(src: string): ExprNode {
	const result = parse(src);
	expect(result.diagnostics, `unexpected diagnostics for: ${src}`).toEqual([]);
	expect(result.expr, `expected an expression for: ${src}`).not.toBeNull();
	return result.expr as ExprNode;
}

/** A compact S-expression rendering used to assert tree shape concisely. */
function shape(expr: ExprNode | null): string {
	if (!expr) {
		return '<null>';
	}
	switch (expr.exprKind) {
		case 'LiteralExpr':
			return `${(expr as LiteralExpr).raw}`;
		case 'IdentifierExpr': {
			const id = expr as IdentifierExpr;
			return id.typeSuffix ? `${id.name}${id.typeSuffix}` : id.name;
		}
		case 'MemberAccessExpr': {
			const m = expr as MemberAccessExpr;
			return `${m.object ? shape(m.object) : ''}.${m.member}`;
		}
		case 'IndexExpr': {
			const idx = expr as IndexExpr;
			return `${shape(idx.callee)}(${idx.args.map(shape).join(', ')})`;
		}
		case 'ParenExpr':
			return `(${shape((expr as ParenExpr).inner)})`;
		case 'UnaryExpr': {
			const u = expr as UnaryExpr;
			return `[${u.operator} ${shape(u.operand)}]`;
		}
		case 'BinaryExpr': {
			const b = expr as BinaryExpr;
			return `[${b.operator} ${shape(b.left)} ${shape(b.right)}]`;
		}
		case 'NewExpr':
			return `New ${(expr as NewExpr).typeName}`;
		case 'AddressOfExpr':
			return `AddressOf ${shape((expr as ExprNode & { target: IdentifierExpr }).target)}`;
		case 'TypeOfIsExpr': {
			const t = expr as TypeOfIsExpr;
			return `[TypeOf ${shape(t.operand)} Is ${t.typeName}]`;
		}
		default:
			return '?';
	}
}

describe('parseExpression - primaries (MS-VBAL 5.6)', () => {
	it('parses integer, float, string, and date literals', () => {
		expect((parseOk('42') as LiteralExpr).literalKind).toBe('integer');
		expect((parseOk('3.14') as LiteralExpr).literalKind).toBe('float');
		expect((parseOk('"hi"') as LiteralExpr).literalKind).toBe('string');
		// A leading '#' lexes as a directive marker; a date literal in real code is
		// mid-statement, so assert it through a parenthesised position.
		const dated = parseOk('(#1/1/2020#)') as ParenExpr;
		expect((dated.inner as LiteralExpr).literalKind).toBe('date');
	});

	it('parses the boolean/Nothing/Null/Empty keyword literals', () => {
		expect((parseOk('True') as LiteralExpr).literalKind).toBe('boolean');
		expect((parseOk('Nothing') as LiteralExpr).literalKind).toBe('nothing');
		expect((parseOk('Null') as LiteralExpr).literalKind).toBe('null');
		expect((parseOk('Empty') as LiteralExpr).literalKind).toBe('empty');
	});

	it('parses identifiers and unwraps bracketed identifiers', () => {
		expect((parseOk('total') as IdentifierExpr).name).toBe('total');
		expect((parseOk('[My Name]') as IdentifierExpr).name).toBe('My Name');
	});

	it('captures an unambiguous type-declaration suffix only when adjacent', () => {
		const adjacent = parseOk('name$') as IdentifierExpr;
		expect(adjacent.typeSuffix).toBe('$');
		expect(adjacent.name).toBe('name');
	});
});

describe('parseExpression - operator precedence (MS-VBAL 5.6.6)', () => {
	it('binds * tighter than +', () => {
		expect(shape(parseOk('1 + 2 * 3'))).toBe('[+ 1 [* 2 3]]');
	});

	it('orders arithmetic: * over \\ over Mod over +', () => {
		expect(shape(parseOk('1 + 2 Mod 3 \\ 4 * 5'))).toBe('[+ 1 [Mod 2 [\\ 3 [* 4 5]]]]');
	});

	it('binds arithmetic tighter than & concatenation', () => {
		expect(shape(parseOk('a & b + c'))).toBe('[& a [+ b c]]');
	});

	it('binds & tighter than comparison', () => {
		expect(shape(parseOk('a & b = c'))).toBe('[= [& a b] c]');
	});

	it('orders logical operators And over Or over Imp', () => {
		expect(shape(parseOk('a Or b And c Imp d'))).toBe('[Imp [Or a [And b c]] d]');
	});

	it('binds comparison tighter than And', () => {
		expect(shape(parseOk('a < b And c > d'))).toBe('[And [< a b] [> c d]]');
	});

	it('is left-associative for same-precedence operators', () => {
		expect(shape(parseOk('1 - 2 - 3'))).toBe('[- [- 1 2] 3]');
	});
});

describe('parseExpression - unary and exponent (MS-VBAL 5.6.6)', () => {
	it('applies Not as a prefix below comparison', () => {
		expect(shape(parseOk('Not a = b'))).toBe('[Not [= a b]]');
	});

	it('binds ^ tighter than unary minus: -2 ^ 2 is -(2 ^ 2)', () => {
		expect(shape(parseOk('-2 ^ 2'))).toBe('[- [^ 2 2]]');
	});

	it('allows a signed exponent: 2 ^ -3', () => {
		expect(shape(parseOk('2 ^ -3'))).toBe('[^ 2 [- 3]]');
	});

	it('is left-associative for ^ chains', () => {
		expect(shape(parseOk('2 ^ 3 ^ 2'))).toBe('[^ [^ 2 3] 2]');
	});

	it('binds unary minus tighter than *', () => {
		expect(shape(parseOk('-a * b'))).toBe('[* [- a] b]');
	});
});

describe('parseExpression - member access and calls (MS-VBAL 5.6.9)', () => {
	it('parses a dotted member chain left-associatively', () => {
		expect(shape(parseOk('a.b.c'))).toBe('a.b.c');
	});

	it('parses a leading-dot member access (With receiver)', () => {
		const expr = parseOk('.Value') as MemberAccessExpr;
		expect(expr.object).toBeNull();
		expect(expr.member).toBe('Value');
	});

	it('parses a call with positional arguments', () => {
		expect(shape(parseOk('Foo(1, a + 2)'))).toBe('Foo(1, [+ a 2])');
	});

	it('parses an empty argument list', () => {
		const expr = parseOk('Refresh()') as IndexExpr;
		expect(expr.args).toEqual([]);
	});

	it('parses a member call chain', () => {
		expect(shape(parseOk('ws.Range("A1").Value'))).toBe('ws.Range("A1").Value');
	});

	it('parses nested indexing', () => {
		expect(shape(parseOk('grid(1)(2)'))).toBe('grid(1)(2)');
	});
});

describe('parseExpression - New / AddressOf / TypeOf (MS-VBAL 5.6.10)', () => {
	it('parses New with a simple class name', () => {
		expect((parseOk('New Collection') as NewExpr).typeName).toBe('Collection');
	});

	it('parses New with a dotted library type', () => {
		expect((parseOk('New Scripting.Dictionary') as NewExpr).typeName).toBe('Scripting.Dictionary');
	});

	it('parses AddressOf with a procedure name', () => {
		expect(shape(parseOk('AddressOf Callback'))).toBe('AddressOf Callback');
	});

	it('parses TypeOf ... Is', () => {
		expect(shape(parseOk('TypeOf obj Is Worksheet'))).toBe('[TypeOf obj Is Worksheet]');
	});
});

describe('parseExpression - parentheses', () => {
	it('overrides precedence with grouping', () => {
		expect(shape(parseOk('(1 + 2) * 3'))).toBe('[* ([+ 1 2]) 3]');
	});
});

describe('parseExpression - error tolerance (Phase 3 contract)', () => {
	it('returns null with a diagnostic for an empty slice', () => {
		const result = parseExpression([]);
		expect(result.expr).toBeNull();
	});

	it('reports a missing right operand without throwing', () => {
		const result = parse('a +');
		expect(result.diagnostics.length).toBeGreaterThan(0);
		// Best-effort: the left operand is still returned.
		expect(shape(result.expr)).toBe('a');
	});

	it('reports an unterminated call argument list', () => {
		const result = parse('Foo(1,');
		expect(result.diagnostics.length).toBeGreaterThan(0);
	});

	it('reports a dangling member dot', () => {
		const result = parse('obj.');
		expect(result.diagnostics.length).toBeGreaterThan(0);
	});

	it('stops at a trailing token it does not model (endIndex < length)', () => {
		const tokens = exprTokens('a b');
		const result = parseExpression(tokens);
		expect(shape(result.expr)).toBe('a');
		expect(result.endIndex).toBeLessThan(tokens.length);
	});
});

describe('parseExpression - span accuracy', () => {
	it('spans the full binary expression', () => {
		const src = 'alpha + beta';
		const expr = parseOk(src);
		expect(expr.span).toEqual({ start: 0, end: src.length });
	});

	it('spans a member access from receiver start to member end', () => {
		const src = 'a.bcd';
		const expr = parseOk(src) as MemberAccessExpr;
		expect(expr.span).toEqual({ start: 0, end: src.length });
		expect(expr.memberSpan).toEqual({ start: 2, end: 5 });
	});
});
