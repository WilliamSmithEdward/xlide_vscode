// Diagnostics tests: expressions rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

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

describe('analyzeModule - division by zero', () => {
	it('errors on literal zero divisors for division, integer division, and Mod', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'    a = 1 \\ 0\n' +
			'    a = 1 Mod 0\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'division-by-zero', [
			{ span: '0' },
			{ span: '0' },
			{ span: '0' },
		]);
	});

	it('detects parenthesized and signed zero divisors inside nested expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Variant\n' +
			'    a = IIf(True, 1, 1 / (-0))\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('-0');
	});

	it('errors on non-decimal zero literal divisors', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Long\n' +
			'    a = 1 / &H0\n' +
			'    a = 1 \\ &O0\n' +
			'    a = 1 Mod &H1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['&H0', '&O0']);
	});

	it('errors on hex and octal Const zero divisors', () => {
		const src =
			'Private Const ModuleHexZero As Long = &H0\n' +
			'Public Sub T()\n' +
			'    Const LocalOctalZero As Long = &O0\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleHexZero\n' +
			'    a = 1 \\ LocalOctalZero\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleHexZero', 'LocalOctalZero']);
	});

	it('folds hex Const expressions for zero divisor checks', () => {
		const src =
			'Public Sub T()\n' +
			'    Const HexZero As Long = &H1 - &H1\n' +
			'    Const HexOne As Long = &H1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HexZero\n' +
			'    a = 1 Mod HexOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HexZero');
	});

	it('errors on zero-valued Enum member divisors', () => {
		const src =
			'Private Enum Denominator\n' +
			'    ExplicitZero = 0\n' +
			'    ImplicitOne\n' +
			'    ExplicitTwo = 2\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ExplicitZero\n' +
			'    a = 1 / ImplicitOne\n' +
			'    a = 1 Mod ExplicitTwo\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('ExplicitZero');
	});

	it('folds implicit and expression Enum member values for zero divisor checks', () => {
		const src =
			'Public Enum Denominator\n' +
			'    ImplicitZero\n' +
			'    ImplicitOne\n' +
			'    ExpressionZero = ImplicitOne - 1\n' +
			'End Enum\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 \\ ImplicitZero\n' +
			'    a = 1 / ExpressionZero\n' +
			'    a = 1 Mod ImplicitOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ImplicitZero', 'ExpressionZero']);
	});

	it('errors on zero-valued module and local Const divisors', () => {
		const src =
			'Private Const ModuleZero As Long = 0\n' +
			'Private Const ModuleOne As Long = 0 + 1\n' +
			'Public Sub T()\n' +
			'    Const LocalZero As Long = 1 - 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / ModuleZero\n' +
			'    a = 1 / LocalZero\n' +
			'    a = 1 / ModuleOne\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['ModuleZero', 'LocalZero']);
	});

	it('folds visible cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 / SharedOne\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'    a = 1 \\ SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['SharedZero', 'SharedZeroDivisor']);
	});

	it('folds verified runtime and host constants for zero divisor checks', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / vbFalse\n' +
			'    a = 1 / VBA.vbFalse\n' +
			'    a = 1 / xlAbove\n' +
			'    a = 1 / xlNo\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');

		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['vbFalse', 'VBA.vbFalse', 'xlAbove']);
	});

	it('lets source declarations shadow verified runtime and host constants', () => {
		const src =
			'Private Const vbFalse As Long = 1\n' +
			'Public Sub T()\n' +
			'    Dim xlAbove As Long\n' +
			'    Dim a As Double\n' +
			'    xlAbove = 1\n' +
			'    a = 1 / vbFalse\n' +
			'    a = 1 / xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});

	it('folds module-qualified cross-module Const and Enum values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedDivisors.SharedZero\n' +
			'    a = 1 / SharedDivisors.SharedOne\n' +
			'    a = 1 \\ SharedDivisors.SharedZeroDivisor\n' +
			'    a = 1 \\ SharedDivisors.SharedOneDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Const SharedOne As Long = 1\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'    SharedOneDivisor = 1\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedDivisors.SharedZero',
			'SharedDivisors.SharedZeroDivisor',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZeroFromHidden\n' +
			'    a = 1 / SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedZeroFromHidden\n' +
			'    a = 1 / SharedDivisors.SharedOneFromHidden\n' +
			'    a = 1 / SharedDivisors.HiddenZero\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenZero As Long = 0\n' +
			'Private Const HiddenOne As Long = 1\n' +
			'Public Const SharedZeroFromHidden As Long = HiddenZero\n' +
			'Public Const SharedOneFromHidden As Long = HiddenOne\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedZeroFromHidden',
			'SharedDivisors.SharedZeroFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for zero divisor checks', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / HiddenZero\n' +
			'    a = 1 / AmbiguousZero\n' +
			'End Sub\n';
		const hidden =
			'Private Const HiddenZero As Long = 0\n';
		const first =
			'Public Const AmbiguousZero As Long = 0\n';
		const second =
			'Public Const AmbiguousZero As Long = 0\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'HiddenDivisors', source: hidden, moduleKind: 'standard' },
				{ moduleName: 'FirstDivisors', source: first, moduleKind: 'standard' },
				{ moduleName: 'SecondDivisors', source: second, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module zero divisors', () => {
		const caller =
			'Private Const SharedZero As Long = 1\n' +
			'Public Sub T()\n' +
			'    Const SharedZeroDivisor As Long = 1\n' +
			'    Dim a As Double\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('lets source variables shadow cross-module zero divisors', () => {
		const caller =
			'Private SharedZero As Double\n' +
			'Public Sub T()\n' +
			'    Dim SharedZeroDivisor As Double\n' +
			'    Dim a As Double\n' +
			'    SharedZero = 2\n' +
			'    SharedZeroDivisor = 2\n' +
			'    a = 1 / SharedZero\n' +
			'    a = 1 \\ SharedZeroDivisor\n' +
			'End Sub\n';
		const shared =
			'Public Const SharedZero As Long = 0\n' +
			'Public Enum SharedDivisor\n' +
			'    SharedZeroDivisor = 0\n' +
			'End Enum\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Caller', source: caller },
				{ moduleName: 'SharedDivisors', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'division-by-zero',
		);

		expect(hits).toHaveLength(0);
	});

	it('folds parenthesized Const expressions used as divisors', () => {
		const src =
			'Private Const Zero As Long = 0\n' +
			'Private Const One As Long = 1\n' +
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    a = 1 / (Zero + 0)\n' +
			'    a = 1 / (Zero + One)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'division-by-zero');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Zero + 0');
	});

	it('does not flag nonzero literals, variables, or nonzero parenthesized expressions', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim a As Double\n' +
			'    Dim denominator As Double\n' +
			'    a = 1 / 2\n' +
			'    a = 1 / denominator\n' +
			'    a = 1 / (0 + 1)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});

	it('ignores literal zero divisors in inactive conditional-compilation branches', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'#If Enabled Then\n' +
			'    Dim a As Double\n' +
			'    a = 1 / 0\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { Enabled: false } },
			}), 'division-by-zero'),
		).toHaveLength(0);
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

	it('flags a standalone zero-argument member call that uses empty parentheses without Call', () => {
		const src = 'Sub T()\n    ThisWorkbook.CanCheckIn()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('CanCheckIn()');
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument host method call with empty parentheses before arity checks', () => {
		const src = 'Sub T()\n    Application.Calculate()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a standalone zero-argument class method call with empty parentheses', () => {
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    p.Save()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags a same-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('leaves required-argument procedure empty calls to arity diagnostics', () => {
		const src =
			'Sub mySub()\n' +
			'    Greet()\n' +
			'End Sub\n' +
			'\n' +
			'Sub Greet(ByVal message As String)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Greet');
	});

	it('flags a standalone zero-argument runtime call with empty parentheses', () => {
		const src = 'Sub T()\n    DoEvents()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'call-statement-forbids-parens');
		expectDiagnostic(src, hits, 'call-statement-forbids-parens', { span: 'DoEvents()' });
		expect(hits[0].message).not.toContain('prefixed with Call');
	});

	it('flags DoEvents as an invalid explicit Call target', () => {
		const src =
			'Sub T()\n' +
			'    Call DoEvents\n' +
			'    Call DoEvents()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-explicit-call-target');
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DoEvents', 'DoEvents']);
		expect(byCode(analyzeModule(src), 'call-requires-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not apply runtime-only call syntax rules to local-shadowed runtime names', () => {
		const src =
			'Sub T()\n' +
			'    Dim Erl\n' +
			'    Erl()\n' +
			'    Call Erl()\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(diagnostics, 'invalid-explicit-call-target')).toHaveLength(0);
	});

	it('flags an unqualified same-class method call statement with empty parentheses', () => {
		const src =
			'Public Sub SaveAll()\n' +
			'    Save()\n' +
			'End Sub\n' +
			'\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'Person', moduleKind: 'class' }),
			'call-statement-forbids-parens',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Save()');
	});

	it('flags an exported cross-module zero-argument Function call statement with empty parentheses', () => {
		const src =
			'Sub mySub()\n' +
			'    myFunction()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(src, [
				{ moduleName: 'Helpers', source: 'Public Function myFunction() As String\nEnd Function\n' },
			], 'Caller'),
			'call-statement-forbids-parens',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('myFunction()');
	});

	it('validates non-empty standalone member call parentheses with known signatures', () => {
		const src = 'Sub T()\n    Application.Calculate(1)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(1);
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

	it('accepts parenthesized member calls with Call or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim ok As Boolean\n' +
			'    Call ThisWorkbook.CanCheckIn()\n' +
			'    ok = ThisWorkbook.CanCheckIn()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('accepts known Function calls with Call or in expression context', () => {
		const src =
			'Sub mySub()\n' +
			'    Dim value As String\n' +
			'    Call myFunction()\n' +
			'    value = myFunction()\n' +
			'End Sub\n' +
			'\n' +
			'Function myFunction() As String\n' +
			'    myFunction = "hello world!"\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('accepts zero-argument runtime calls as bare statements or in expression context', () => {
		const src =
			'Sub T()\n' +
			'    Dim value As Integer\n' +
			'    DoEvents\n' +
			'    value = DoEvents()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'invalid-explicit-call-target')).toHaveLength(0);
	});

	it('does not flag unknown bare empty-parentheses statements as project procedure calls', () => {
		const src = 'Sub T()\n    MaybeExternal()\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
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

	it('does not require expression-call parentheses for local-shadowed runtime names', () => {
		const src =
			'Sub T()\n' +
			'    Dim Format\n' +
			'    answer = Format 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('flags a unique exported project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('flags a module-qualified project Function called with parenless arguments in an assignment', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim total As Double\n' +
			'    total = Invoices.InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'expression-call-requires-parens',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('InvoiceTotal');
	});

	it('does not guess when a bare exported project Function name is ambiguous', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal 100, 0.08\n' +
			'End Sub\n';
		const first =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const second =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Invoices', source: first },
			{ moduleName: 'AlternateInvoices', source: second },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
	});

	it('does not treat exported project Subs as expression Functions', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = PrintTotal 100\n' +
			'End Sub\n';
		const helpers = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';

		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller'), 'expression-call-requires-parens')).toHaveLength(0);
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

	it('accepts parameterless property reads followed by keyword operators', () => {
		const src =
			'Public Property Get Style() As Long\nEnd Property\n' +
			'Public Sub T()\n' +
			'    Dim resizable As Boolean\n' +
			'    resizable = Style And &H40000\n' +
			'    resizable = Style Or &H40000\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'expression-call-requires-parens')).toHaveLength(0);
	});
});

describe('analyzeModule - invalid expression syntax', () => {
	it('flags an impossible operator sequence in a call argument expression', () => {
		const src =
			'Sub T()\n' +
			'    MsgBox myFunctionTest***\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('***');
	});

	it('flags unsupported C-style ternary syntax', () => {
		const src = 'Sub T()\n    value = flag ? 1 : 2\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'invalid-expression-syntax', { span: '?' });
	});

	it('flags a statement that ends with a binary operator', () => {
		const src = 'Sub T()\n    total = subtotal *\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('*');
	});

	it('flags trailing member-access dots on object receivers', () => {
		const src = 'Sub T()\n    ThisWorkbook.\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'invalid-expression-syntax', { span: '.' });
	});

	it('flags a bare leading dot inside With as incomplete final source', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('accepts complete leading-dot member access inside With', () => {
		const src =
			'Sub T()\n' +
			'    With ActiveSheet\n' +
			'        .Range("A1").Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('lets scalar member access own known scalar trailing dots', () => {
		const src = 'Sub T()\n    Dim value As String\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(1);
	});

	it('lets scalar member access own visible exported scalar trailing dots', () => {
		const caller = 'Sub T()\n    SharedText.\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'invalid-expression-syntax')).toHaveLength(0);
		const hits = byCode(diagnostics, 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedText.');
	});

	it('uses generic incomplete member syntax for local shadows of exported scalars', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedText\n' +
			'    SharedText.\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
		const hits = byCode(diagnostics, 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('.');
	});

	it('treats fixed-length String declarations as scalar String receivers', () => {
		const src = 'Sub T()\n    Dim value As String * 20\n    value.\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag valid arithmetic or string expressions', () => {
		const src =
			'Sub T()\n' +
			'    total = subtotal * taxRate\n' +
			'    message = prefix & suffix\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('accepts IIf as the supported inline conditional function', () => {
		const src = 'Sub T()\n    value = IIf(flag, 1, 2)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});
});
