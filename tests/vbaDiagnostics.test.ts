import { describe, it, expect } from 'vitest';
import { analyzeModule, VbaDiagnostic, DIAGNOSTIC_RULES } from '../src/analyzer';

/** Returns the diagnostics whose code matches `code`. */
function byCode(diags: VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((d) => d.code === code);
}

/** Resolves the source substring a diagnostic span covers. */
function spanText(source: string, d: VbaDiagnostic): string {
	return source.slice(d.span.start, d.span.end);
}

describe('analyzeModule - unterminated string', () => {
	it('flags a string with no closing quote', () => {
		const src = 'Sub T()\n    MsgBox "hello\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unterminated-string');
		expect(hits.length).toBe(1);
		expect(spanText(src, hits[0])).toBe('"hello');
		expect(hits[0].severity).toBe('error');
	});

	it('accepts a properly closed string, including doubled-quote escapes', () => {
		const src = 'Sub T()\n    MsgBox "say ""hi"" now"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(0);
	});

	it('treats a trailing escaped pair without a real close as unterminated', () => {
		const src = 'Sub T()\n    x = "ab""\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(1);
	});
});

describe('analyzeModule - duplicate procedures', () => {
	it('flags two Subs with the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-procedure');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Foo');
	});

	it('flags a Function colliding with a Sub of the same name', () => {
		const src = 'Sub Foo()\nEnd Sub\nFunction Foo()\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});

	it('allows Property Get/Let/Set to share a name', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Let Item(v As Long)\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(0);
	});

	it('flags a duplicate Property Get', () => {
		const src =
			'Property Get Item() As Long\nEnd Property\n' +
			'Property Get Item() As Long\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'duplicate-procedure')).toHaveLength(1);
	});
});

describe('analyzeModule - duplicate declarations in scope', () => {
	it('flags the same local declared twice', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim x As String\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-declaration');
		expect(hits).toHaveLength(1);
	});

	it('flags a local colliding with a parameter', () => {
		const src = 'Sub T(ByVal n As Long)\n    Dim n As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('flags duplicate parameter names', () => {
		const src = 'Sub T(a As Long, a As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('treats procedure scope as flat across branches', () => {
		const src =
			'Sub T()\n' +
			'    If True Then\n        Dim x As Long\n    End If\n' +
			'    Dim x As Long\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(1);
	});

	it('does not flag distinct local names', () => {
		const src = 'Sub T()\n    Dim x As Long\n    Dim y As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-declaration')).toHaveLength(0);
	});
});

describe('analyzeModule - duplicate module members', () => {
	it('flags the same module variable declared twice', () => {
		const src = 'Private Total As Long\nPublic Total As String\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(1);
	});

	it('does not flag distinct module variables', () => {
		const src = 'Private A As Long\nPrivate B As Long\n';
		expect(byCode(analyzeModule(src), 'duplicate-module-variable')).toHaveLength(0);
	});
});

describe('analyzeModule - assignment to constant', () => {
	it('flags assigning to a module-level Const', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    MAX = 5\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
	});

	it('flags assigning to a local Const', () => {
		const src = 'Sub T()\n    Const PI As Double = 3.14\n    PI = 3\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(1);
	});

	it('does not flag comparing a Const in a condition', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    If MAX = 10 Then Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag assigning to a non-constant variable', () => {
		const src = 'Sub T()\n    Dim x As Long\n    x = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('does not flag a member or indexed left-hand side that shares a Const name', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n    obj.MAX = 5\n    arr(MAX) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});
});

describe('analyzeModule - Option Explicit', () => {
	it('warns when a code module omits Option Explicit', () => {
		const src = 'Sub T()\n    Dim x As Long\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'option-explicit-missing');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('warning');
	});

	it('is silent when Option Explicit is present', () => {
		const src = 'Option Explicit\n\nSub T()\n    Dim x As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'option-explicit-missing')).toHaveLength(0);
	});

	it('is silent for an attribute-only / empty module', () => {
		const src = 'Attribute VB_Name = "Sheet1"\n';
		expect(byCode(analyzeModule(src), 'option-explicit-missing')).toHaveLength(0);
	});

	it('respects a severity override', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severities: { optionExplicitMissing: 'error' } }),
			'option-explicit-missing',
		);
		expect(hits[0].severity).toBe('error');
	});

	it('can be switched off', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severities: { optionExplicitMissing: 'off' } }),
			'option-explicit-missing',
		);
		expect(hits).toHaveLength(0);
	});
});

describe('analyzeModule - general contract', () => {
	it('never throws on malformed input', () => {
		const samples = ['', 'Sub', 'End Sub', 'Dim', '"', 'If Then', ':::'];
		for (const s of samples) {
			expect(() => analyzeModule(s)).not.toThrow();
		}
	});

	it('every emitted code is in the rule catalogue', () => {
		const known = new Set(
			Object.values(DIAGNOSTIC_RULES).map((r) => r.code),
		);
		const src =
			'Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\n' +
			'Const C As Long = 1\nSub Bar()\n    C = 2\n    MsgBox "x\nEnd Sub\n';
		for (const d of analyzeModule(src)) {
			expect(known.has(d.code)).toBe(true);
		}
	});

	it('produces a clean module with no diagnostics', () => {
		const src =
			'Option Explicit\n\n' +
			'Const MAX As Long = 10\n\n' +
			'Sub Greet(ByVal name As String)\n' +
			'    Dim msg As String\n' +
			'    msg = "Hello " & name\n' +
			'    MsgBox msg\n' +
			'End Sub\n';
		expect(analyzeModule(src)).toHaveLength(0);
	});
});

describe('analyzeModule - unknown call statement', () => {
	const opts = { knownProcedures: new Set<string>() };

	it('flags a bare identifier that resolves to nothing', () => {
		const src = 'Sub Main()\n    asdfjalsdkfjas\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('asdfjalsdkfjas');
		expect(hits[0].severity).toBe('error');
	});

	it('does not flag a call to a procedure in another module', () => {
		const src = 'Sub Main()\n    DoWork\nEnd Sub\n';
		const known = { knownProcedures: new Set(['dowork']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a call to a Sub defined in the same module', () => {
		const src = 'Sub Main()\n    Helper\nEnd Sub\nSub Helper()\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag VBA runtime functions/statements used bare', () => {
		const src = 'Sub Main()\n    DoEvents\n    Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a host global / Application member used bare', () => {
		const src = 'Sub Main()\n    Calculate\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag an in-scope local variable used bare', () => {
		const src = 'Sub Main()\n    Dim total As Long\n    total\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not run at all when knownProcedures is omitted', () => {
		const src = 'Sub Main()\n    asdfjalsdkfjas\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unknown-call')).toHaveLength(0);
	});

	it('ignores assignments, member calls, and known arg-bearing calls', () => {
		const src =
			'Sub Main()\n' +
			'    x = 1\n' +
			'    Debug.Print x\n' +
			'    MsgBox "hi"\n' +
			'    Foo 1, 2\n' +
			'End Sub\n';
		const known = { knownProcedures: new Set(['foo']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('ignores a line label', () => {
		const src = 'Sub Main()\n    GoTo done\ndone:\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('flags an unknown parenless call with arguments', () => {
		const src = 'Sub Main()\n    msrbox ""\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('msrbox');
	});

	it('flags an unknown call with multiple arguments', () => {
		const src = 'Sub Main()\n    Frobnicate 1, 2, 3\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Frobnicate');
	});

	it('flags an unknown explicit Call statement', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
	});

	it('does not flag a known procedure called with arguments', () => {
		const src = 'Sub Main()\n    DoWork 1, 2\nEnd Sub\n';
		const known = { knownProcedures: new Set(['dowork']) };
		expect(byCode(analyzeModule(src, known), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a runtime function called with arguments', () => {
		const src = 'Sub Main()\n    Debug.Print Left("abc", 1)\n    MsgBox "hi", vbOKOnly\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});

	it('does not flag a bare host member used with parentheses', () => {
		const src =
			'Sub Main()\n' +
			'    Cells(1, 1).Select\n' +
			'    Range("A1").Value = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid procedure header', () => {
	it('flags a procedure name that contains a space', () => {
		const src = 'Sub My Sub\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Sub');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a junk token after a Function name', () => {
		const src = 'Function Calc Total() As Long\nEnd Function\n';
		const hits = byCode(analyzeModule(src), 'invalid-proc-header');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Total');
	});

	it('does not flag a valid parameterless Sub', () => {
		const src = 'Sub Run\n    Beep\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a valid Sub with parameters', () => {
		const src = 'Sub Greet(ByVal who As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Function with a return type', () => {
		const src = 'Function Total() As Long\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});

	it('does not flag a Property Get with a return type', () => {
		const src = 'Property Get Name() As String\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'invalid-proc-header')).toHaveLength(0);
	});
});

describe('analyzeModule - unbalanced parentheses', () => {
	it('flags a missing closing parenthesis', () => {
		const src = 'Sub T()\n    x = (1 + 2\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('(');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unexpected closing parenthesis', () => {
		const src = 'Sub T()\n    x = 1 + 2)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});

	it('reports at most one diagnostic per statement', () => {
		const src = 'Sub T()\n    x = ((1 + 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(1);
	});

	it('does not flag balanced and nested parentheses', () => {
		const src =
			'Sub T()\n' +
			'    x = (1 + (2 * 3))\n' +
			'    Debug.Print Left("ab", 1)\n' +
			'    Cells(1, 1).Value = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('does not count parentheses inside strings or comments', () => {
		const src =
			'Sub T()\n' +
			'    x = "a (b c"  \' a ) in a comment\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('balances across a line continuation', () => {
		const src = 'Sub T()\n    x = (1 + _\n        2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'unbalanced-parens')).toHaveLength(0);
	});

	it('treats a colon-separated statement independently', () => {
		const src = 'Sub T()\n    a = (1 + 2) : b = 3)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'unbalanced-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe(')');
	});
});

describe('analyzeModule - argument count', () => {
	it('flags too few arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Greet');
		expect(hits[0].message).toContain('expected 2 arguments');
		expect(hits[0].message).toContain('got 1');
	});

	it('flags too many arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob", "Cat"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('got 3');
	});

	it('flags a lone-identifier call that omits required arguments', () => {
		const src =
			'Sub Main()\n' +
			'    Greet\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('got 0');
	});

	it('validates an explicit Call statement', () => {
		const src =
			'Sub Main()\n' +
			'    Call Greet("Ann", "Bob", "Cat")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
	});

	it('accepts a correct argument count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours Optional parameters', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'    Greet "Ann", "Bob"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('honours ParamArray (no upper bound)', () => {
		const src =
			'Sub Main()\n' +
			'    Log "a", "b", "c", "d"\n' +
			'End Sub\n' +
			'Sub Log(ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('validates named-argument names but not the count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet who:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('who');
		expect(hits[0].message).toContain('Named argument not found');
	});

	it('accepts a valid named argument', () => {
		const src =
			'Sub Main()\n' +
			'    Greet a:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, Optional ByVal b As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check calls to unknown or cross-module procedures', () => {
		const src =
			'Sub Main()\n' +
			'    SomethingElse 1, 2, 3\n' +
			'    MsgBox "hi", vbOKOnly, "title", 0, 0\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not arity-check an ambiguous (duplicated) procedure name', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});
});

