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

describe('analyzeModule - non-callable call statements', () => {
	it('flags a bare local variable used as a statement', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('testStr');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('local variable');
	});

	it('flags an explicit Call to a parameter', () => {
		const src = 'Sub Main(ByVal testStr As String)\n    Call testStr\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('testStr');
		expect(hits[0].message).toContain('parameter');
	});

	it('flags a module variable used as a call target', () => {
		const src = 'Private total As Long\nSub Main()\n    total\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'non-callable-call');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('module variable');
	});

	it('does not flag callable procedures or runtime statements', () => {
		const src = 'Sub Main()\n    Helper\n    Beep\nEnd Sub\nSub Helper()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'non-callable-call')).toHaveLength(0);
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

	it('flags an omitted required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('TaxRate');
		expect(hits[0].message).toContain('Argument not optional');
	});

	it('flags an omitted leading required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(, 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Subtotal');
	});

	it('accepts an omitted Optional argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(total, )\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
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

describe('analyzeModule - argument type validation', () => {
	it('flags a nonnumeric string literal passed to a same-module numeric parameter', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].severity).toBe('error');
		expect(hits[0].message).toContain('Subtotal');
		expect(hits[0].message).toContain('Currency');
	});

	it('accepts numeric literals and numeric string literals for numeric parameters', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    a = InvoiceTotal(100, 0.08)\n' +
			'    b = InvoiceTotal("100", 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn on string variables whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim label As String\n' +
			'    total = InvoiceTotal(label)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('maps named arguments to the named parameter type', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(TaxRate:=0.08, Subtotal:="blah")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"blah"');
		expect(hits[0].message).toContain('Subtotal');
	});

	it('validates selected native VBA runtime parameter types', () => {
		const src = 'Sub T()\n    x = Left("abcdef", "bad")\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"bad"');
		expect(hits[0].message).toContain('Length');
	});

	it('does not infer runtime parameter types from display names', () => {
		const src = 'Sub T()\n    Randomize "bad"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn when a type is unknown or Variant-like', () => {
		const src = 'Sub T()\n    x = Format("blah", "0.00")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - declaration initializer', () => {
	it('flags a module-level Dim with a VB.NET-style initializer', () => {
		const src = 'Private x As Long = 1\n';
		const hits = byCode(analyzeModule(src), 'dim-initializer');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('=');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a local Dim with an initializer', () => {
		const src = 'Sub T()\n    Dim n As Long = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('flags an initializer without an As clause', () => {
		const src = 'Sub T()\n    Dim n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(1);
	});

	it('does not flag a Const declaration', () => {
		const src = 'Const Pi As Double = 3.14\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag a plain declaration', () => {
		const src = 'Sub T()\n    Dim n As Long\n    n = 5\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});

	it('does not flag an array bound that contains digits', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'dim-initializer')).toHaveLength(0);
	});
});

describe('analyzeModule - Call requires parentheses', () => {
	it('flags an unparenthesised Call argument list', () => {
		const src = 'Sub T()\n    Call MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"hello"');
		expect(hits[0].severity).toBe('error');
	});

	it('flags an unparenthesised member Call', () => {
		const src = 'Sub T()\n    Call obj.Method 1, 2\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(1);
	});

	it('accepts Call with parentheses', () => {
		const src = 'Sub T()\n    Call MsgBox("hello")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Call', () => {
		const src = 'Sub T()\n    Call DoWork\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('accepts a Call to a parenthesised member chain', () => {
		const src = 'Sub T()\n    Call obj.Method(1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});

	it('does not flag a parenless non-Call statement', () => {
		const src = 'Sub T()\n    MsgBox "hello"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - expression call requires parentheses', () => {
	it('flags a same-module Function called with parenless arguments in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a runtime Function called with parenless arguments in an assignment', () => {
		const src = 'Sub T()\n    answer = MsgBox "hello"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'expression-call-requires-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MsgBox');
	});

	it('accepts a parenthesized Function call in an assignment', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal(100, 0.08)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('accepts a parameterless Function reference in an expression', () => {
		const src =
			'Public Function CurrentTotal() As Currency\nEnd Function\n' +
			'Public Sub TestTotal()\n' +
			'    total = CurrentTotal + 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - parameter order', () => {
	it('flags a required parameter after an Optional one', () => {
		const src =
			'Sub T(Optional ByVal a As Long = 1, ByVal b As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'required-param-after-optional');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toContain('b');
	});

	it('accepts trailing Optional parameters', () => {
		const src =
			'Sub T(ByVal a As Long, Optional ByVal b As Long = 1)\nEnd Sub\n';
		expect(
			byCode(analyzeModule(src), 'required-param-after-optional'),
		).toHaveLength(0);
	});

	it('flags a ParamArray that is not last', () => {
		const src = 'Sub T(ParamArray items() As Variant, ByVal n As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'paramarray-not-last');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toContain('items');
	});

	it('accepts a trailing ParamArray', () => {
		const src = 'Sub T(ByVal n As Long, ParamArray items() As Variant)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'paramarray-not-last')).toHaveLength(0);
	});
});

describe('analyzeModule - Exit statement matches procedure', () => {
	it('flags Exit Function inside a Sub', () => {
		const src = 'Sub T()\n    Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
		expect(hits[0].severity).toBe('error');
	});

	it('flags Exit Sub inside a Function', () => {
		const src = 'Function F() As Long\n    Exit Sub\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('flags Exit Sub inside a Property', () => {
		const src = 'Property Get Name() As String\n    Exit Sub\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('accepts a matching Exit Sub', () => {
		const src = 'Sub T()\n    Exit Sub\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('accepts a matching Exit Property', () => {
		const src = 'Property Let Name(v As String)\n    Exit Property\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('ignores Exit For and Exit Do', () => {
		const src =
			'Sub T()\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});
});

describe('analyzeModule - statement context', () => {
	it('flags an If statement missing Then', () => {
		const src = 'Sub T()\n    If x > 0\n        x = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'if-missing-then');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('If');
	});

	it('accepts multiline and single-line If statements with Then', () => {
		const src =
			'Sub T()\n' +
			'    If x > 0 Then\n        x = 1\n    End If\n' +
			'    If x > 1 Then x = 2\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'if-missing-then')).toHaveLength(0);
	});

	it('flags Case outside Select Case', () => {
		const src = 'Sub T()\n    Case 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'case-outside-select');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Case');
	});

	it('accepts Case inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'    Select Case x\n' +
			'        Case 1, 2\n' +
			'            x = 3\n' +
			'        Case Else\n' +
			'            x = 4\n' +
			'    End Select\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
	});

	it('flags leading-dot member access outside With', () => {
		const src = 'Sub T()\n    .Value = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'member-access-outside-with');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With Range("A1")\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
	});

	it('flags Exit For and Exit Do outside matching loops', () => {
		const src = 'Sub T()\n    Exit For\n    Exit Do\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-outside-block');
		expect(hits).toHaveLength(2);
		expect(spanText(src, hits[0])).toBe('Exit For');
		expect(spanText(src, hits[1])).toBe('Exit Do');
	});

	it('accepts Exit For and Exit Do inside matching loops', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n        Exit For\n    Next i\n' +
			'    Do\n        Exit Do\n    Loop\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-outside-block')).toHaveLength(0);
	});
});

describe('analyzeModule - Option placement', () => {
	it('flags an Option after a declaration', () => {
		const src = 'Private m_Count As Long\nOption Explicit\n';
		const hits = byCode(analyzeModule(src), 'option-after-declaration');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
	});

	it('flags an Option after a procedure', () => {
		const src = 'Sub T()\nEnd Sub\nOption Base 1\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(1);
	});

	it('accepts Options at the top of the module', () => {
		const src = 'Option Explicit\nOption Base 1\nPrivate m_Count As Long\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});

	it('accepts Options after Attribute lines', () => {
		const src =
			'Attribute VB_Name = "Module1"\nOption Explicit\nSub T()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'option-after-declaration')).toHaveLength(0);
	});
});

