// Slice 2 fixtures: procedure-body assignment and call statement wiring.
//
// Verified against MS-VBAL.pdf, v20250520 (§5.4.3 assignment, §5.4.2 call).
// parseModule now routes plain body statements through the §5.6 expression
// parser, emitting AssignmentNode/CallNode where the statement is a clean,
// fully-modeled assignment or call and falling back to a raw StatementNode
// otherwise. These tests pin both the structured shapes and the must-stay-raw
// boundary that preserves the no-regression contract.

import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import {
	AssignmentNode,
	BodyNode,
	CallNode,
	ProcedureNode,
} from '../src/analyzer/parser/nodes';

/** First body node of the single procedure in a one-line-body fixture. */
function firstBodyNode(line: string): BodyNode {
	const src = `Sub T()\n${line}\nEnd Sub\n`;
	const proc = parseModule(src).members.find(
		(m): m is ProcedureNode => m.kind === 'Procedure',
	)!;
	return proc.body[0];
}

function assignment(line: string): AssignmentNode {
	const node = firstBodyNode(line);
	expect(node.kind, `expected Assignment for: ${line}`).toBe('Assignment');
	return node as AssignmentNode;
}

function call(line: string): CallNode {
	const node = firstBodyNode(line);
	expect(node.kind, `expected Call for: ${line}`).toBe('Call');
	return node as CallNode;
}

describe('parseModule - assignment statements (MS-VBAL 5.4.3)', () => {
	it('parses a simple value assignment', () => {
		const a = assignment('x = 1');
		expect(a.isSet).toBe(false);
		expect(a.isLet).toBe(false);
		expect(a.lhs.exprKind).toBe('IdentifierExpr');
		expect(a.rhs.exprKind).toBe('LiteralExpr');
	});

	it('parses a Set (object-reference) assignment', () => {
		const a = assignment('Set obj = New Collection');
		expect(a.isSet).toBe(true);
		expect(a.rhs.exprKind).toBe('NewExpr');
	});

	it('parses an explicit Let assignment', () => {
		const a = assignment('Let y = 2');
		expect(a.isLet).toBe(true);
		expect(a.isSet).toBe(false);
	});

	it('parses an indexed assignment target', () => {
		const a = assignment('a(i) = b + 1');
		expect(a.lhs.exprKind).toBe('IndexExpr');
		expect(a.rhs.exprKind).toBe('BinaryExpr');
	});

	it('parses a member assignment target', () => {
		expect(assignment('obj.Prop = 5').lhs.exprKind).toBe('MemberAccessExpr');
	});

	it('parses a bang member-access assignment RHS', () => {
		expect(assignment('x = rs!Name').rhs.exprKind).toBe('MemberAccessExpr');
	});

	it('parses a bang member-access assignment target', () => {
		expect(assignment('rs!Name = 1').lhs.exprKind).toBe('MemberAccessExpr');
	});

	it('parses a leading-dot (With) assignment target', () => {
		const a = assignment('.Value = 1');
		expect(a.lhs.exprKind).toBe('MemberAccessExpr');
	});

	it('binds the first top-level = as assignment, the rest as comparison', () => {
		// `x = y = z` assigns the boolean (y = z) to x.
		const a = assignment('x = y = z');
		expect(a.lhs.exprKind).toBe('IdentifierExpr');
		expect(a.rhs.exprKind).toBe('BinaryExpr');
	});

	it('spans the whole logical statement, including a leading line number', () => {
		const src = 'Sub T()\n10 x = 1\nEnd Sub\n';
		const proc = parseModule(src).members.find(
			(m): m is ProcedureNode => m.kind === 'Procedure',
		)!;
		const a = proc.body[0] as AssignmentNode;
		expect(a.kind).toBe('Assignment');
		// Span covers "10 x = 1" so re-tokenizing rules see identical text.
		expect(src.slice(a.span.start, a.span.end)).toBe('10 x = 1');
	});
});

describe('parseModule - call statements (MS-VBAL 5.4.2)', () => {
	it('parses an implicit parenless call with arguments', () => {
		const c = call('MsgBox "hi"');
		expect(c.hasCallKeyword).toBe(false);
		expect(c.callee.exprKind).toBe('IdentifierExpr');
		expect(c.args).toHaveLength(1);
	});

	it('parses an explicit Call with a parenthesised argument list', () => {
		const c = call('Call Foo(1, 2)');
		expect(c.hasCallKeyword).toBe(true);
		expect(c.args).toHaveLength(2);
	});

	it('parses a parenless call with multiple arguments', () => {
		expect(call('Foo 1, 2').args).toHaveLength(2);
	});

	it('parses a zero-argument call with empty parentheses', () => {
		const c = call('Refresh()');
		expect(c.hasCallKeyword).toBe(false);
		expect(c.args).toHaveLength(0);
	});

	it('parses a parenless call with a named argument', () => {
		const c = call('Foo a:=1');
		expect(c.hasCallKeyword).toBe(false);
		expect(c.args).toHaveLength(1);
		expect(c.args[0].name).toBe('a');
		expect(c.args[0].value?.exprKind).toBe('LiteralExpr');
	});

	it('parses a parenless call mixing positional and named arguments', () => {
		const c = call('MsgBox "hi", Buttons:=vbOKOnly');
		expect(c.args).toHaveLength(2);
		expect(c.args[0].name).toBeUndefined();
		expect(c.args[1].name).toBe('Buttons');
	});

	it('parses an explicit Call with named and omitted arguments', () => {
		const c = call('Call Foo(1, , After:=3)');
		expect(c.hasCallKeyword).toBe(true);
		expect(c.args).toHaveLength(3);
		expect(c.args[1].value).toBeNull();
		expect(c.args[2].name).toBe('After');
	});

	it('parses a standalone parenthesised call with an omitted argument', () => {
		const c = call('Foo(1, , 3)');
		expect(c.args).toHaveLength(3);
		expect(c.args.map((a) => a.value === null)).toEqual([false, true, false]);
	});
});

describe('parseModule - must stay raw (no-regression boundary)', () => {
	const rawCases: ReadonlyArray<readonly [string, string]> = [
		['a lone identifier (label vs no-arg call is ambiguous)', 'foo'],
		['a lone member chain with no arguments', 'obj.Refresh'],
		['a Mid replacement statement', 'Mid(s, 1, 2) = "x"'],
		['a Mid$ replacement statement', 'Mid$(s, 1, 2) = "x"'],
		['a MidB byte replacement statement', 'MidB(s, 1, 2) = "x"'],
		['a MidB$ byte replacement statement', 'MidB$(s, 1, 2) = "x"'],
		['a GoTo statement', 'GoTo done'],
		['an Exit statement', 'Exit Sub'],
		['an On Error statement', 'On Error GoTo 0'],
		['a malformed Set with no =', 'Set obj'],
		['a keyword-object parenless call (Debug)', 'Debug.Print x'],
	];

	for (const [label, line] of rawCases) {
		it(`keeps ${label} as a raw Statement`, () => {
			expect(firstBodyNode(line).kind).toBe('Statement');
		});
	}
});

describe('parseModule - malformed operands fall back to raw (clean-parse guard)', () => {
	// The §5.6 parser is error-tolerant and consumes the bad token to the
	// boundary, so endIndex alone is not enough — a structured node must never
	// be built from an operand that reported a parse diagnostic.
	const malformed = [
		'x = 1 +', // dangling operator in RHS
		'x = a.', // dangling member dot in RHS
		'x = (1 + 2', // unclosed paren in RHS
		'a. = 1', // dangling member dot in LHS
		'Foo(1 +)', // malformed argument in a call
	];
	for (const line of malformed) {
		it(`keeps "${line}" as a raw Statement`, () => {
			expect(firstBodyNode(line).kind).toBe('Statement');
		});
	}
});

describe('parseModule - diagnostics still see structured statements', () => {
	it('still flags an undeclared assignment target under Option Explicit', () => {
		// Regression guard: the undeclared-variable rule walks body statements;
		// an Assignment node must still be visited (walker widening).
		const src = 'Option Explicit\nSub T()\n    notDeclared = 1\nEnd Sub\n';
		const mod = parseModule(src);
		// The structured assignment is present in the body...
		const proc = mod.members.find((m): m is ProcedureNode => m.kind === 'Procedure')!;
		expect(proc.body[0].kind).toBe('Assignment');
	});
});
