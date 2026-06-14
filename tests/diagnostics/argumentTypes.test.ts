// Diagnostics tests: argumentTypes rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule, projectClassMembers, projectProcedures } from './helpers';

describe('analyzeModule - argument type validation', () => {
	it('flags a nonnumeric string literal passed to a same-module numeric parameter', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', {
			severity: 'error',
			span: '"blah"',
		});
	});

	it('flags intrinsic CVErr Error Variants passed to scalar parameters', () => {
		const src =
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub NeedsVariant(ByVal value As Variant)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong CVErr(2015)\n' +
			'    NeedsLong VBA.CVErr(2015)\n' +
			'    NeedsVariant CVErr(2015)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');

		expectDiagnostics(src, hits, 'argument-type-mismatch', [
			{ span: 'CVErr(2015)' },
			{ span: 'VBA.CVErr(2015)' },
		]);
	});

	it('flags Null passed to scalar parameters', () => {
		const src =
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub NeedsString(ByVal value As String)\n' +
			'End Sub\n' +
			'Public Sub NeedsVariant(ByVal value As Variant)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong Null\n' +
			'    NeedsString Null\n' +
			'    NeedsVariant Null\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-type-mismatch');

		expectDiagnostics(src, hits, 'argument-type-mismatch', [{ span: 'Null' }, { span: 'Null' }]);
	});

	it('lets a source function named CVErr shadow the intrinsic in argument expressions', () => {
		const src =
			'Private Function CVErr(ByVal errorNumber As Long) As Long\n' +
			'    CVErr = errorNumber\n' +
			'End Function\n' +
			'Public Sub NeedsLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsLong CVErr(2015)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('uses a unique exported project signature for cross-module argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Invoices', source: invoices },
			], 'Caller'),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"blah"' });
	});

	it('validates same-module Declare argument types', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Public Sub T()\n' +
			'    Sleep "bad"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses module-qualified project signatures for argument types', () => {
		const caller =
			'Public Sub TestInvoiceTotal()\n' +
			'    total = Invoices.InvoiceTotal("blah", 0.08)\n' +
			'End Sub\n';
		const invoices =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Invoices', source: invoices },
				]),
			}),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"blah"' });
	});

	it('uses a unique exported project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses visible exported scalar globals as bare argument value types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'argument-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses module-qualified exported scalar globals as argument value types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Globals.SharedText\n' +
			'    NeedsObject Globals.SharedText & ""\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'argument-object-type-mismatch',
		);

		expectDiagnostics(caller, hits, 'argument-object-type-mismatch', [
			{ span: 'SharedText', message: 'Globals.SharedText As String' },
			{ span: 'Globals.SharedText & ""', message: 'string concatenation expression' },
		]);
	});

	it('keeps unknown module-qualified value types quiet for argument inference', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Missing.SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(caller, [], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('keeps local shadows and ambiguous exported globals quiet for bare argument value types', () => {
		const shadowCaller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim SharedText\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses a module-qualified project function return type in nested calls', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Labels.MakeLabel()\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses source-backed member return types in argument expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'Name' });
	});

	it('does not infer bare member functions that require arguments as values', () => {
		const person =
			'Public Function Name(ByVal index As Long) As String\n' +
			'End Function\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses member return types inside string concatenation argument expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    NeedsObject p.Name & "!"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'p.Name & "!"' });
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

	it('errors on decimal literals outside Byte and Integer parameter bounds', () => {
		const src =
			'Public Sub TakesByte(ByVal value As Byte)\n' +
			'End Sub\n' +
			'Public Sub TakesInteger(ByVal value As Integer)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    TakesByte 0\n' +
			'    TakesByte 255\n' +
			'    TakesByte 256\n' +
			'    TakesByte -1\n' +
			'    TakesInteger -32768\n' +
			'    TakesInteger 32767\n' +
			'    TakesInteger 32768\n' +
			'    TakesInteger -32769\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-type-mismatch', [
			{ span: '256', message: ['Byte', "Run-time error '6'"] },
			{ span: '-1', message: ['Byte', "Run-time error '6'"] },
			{ span: '32768', message: ['Integer', "Run-time error '6'"] },
			{ span: '-32769', message: ['Integer', "Run-time error '6'"] },
		]);
	});

	it('errors on a decimal literal outside Long parameter bounds', () => {
		// VBE oracle: long_argument_overflow_literal_runtime compiles then raises
		// Run-time error '6': Overflow; the 2147483647 control runs clean. Hex and
		// LongLong arguments carry no provable overflow, so they stay silent.
		const src =
			'Public Sub TakesLong(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub TakesLongLong(ByVal value As LongLong)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    TakesLong 2147483647\n' +
			'    TakesLong 2147483648\n' +
			'    TakesLong &HFFFFFFFF\n' +
			'    TakesLongLong 5000000000\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-type-mismatch', [
			{ span: '2147483648', message: ['Long', "Run-time error '6'"] },
		]);
	});

	it('errors on a whole-number literal outside Currency parameter bounds', () => {
		// VBE oracle: currency_argument_overflow_literal_runtime compiles then
		// raises Run-time error '6': Overflow; the 922337203685477 control runs clean.
		const src =
			'Public Sub TakesCurrency(ByVal value As Currency)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    TakesCurrency 922337203685477\n' +
			'    TakesCurrency 922337203685478\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-type-mismatch', [
			{ span: '922337203685478', message: ['Currency', "Run-time error '6'"] },
		]);
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

	it('flags known scalar variable type mismatches for ByRef parameters', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'byref-argument-type-mismatch', { span: 'amount' });
	});

	it('treats omitted parameter passing markers as default ByRef', () => {
		const src =
			'Public Sub Mutate(value As Long)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(1);
	});

	it('flags known Object variable mismatches for ByRef scalar parameters', () => {
		const src =
			'Public Sub NeedsString(ByRef value As String)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    NeedsString obj\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');

		expectDiagnostic(src, hits, 'byref-argument-type-mismatch', { span: 'obj' });
	});

	it('does not apply ByRef exactness to ByVal parameters, literals, or expression temporaries', () => {
		const src =
			'Public Sub Mutate(ByRef value As Long)\n' +
			'End Sub\n' +
			'Public Sub ReadValue(ByVal value As Long)\n' +
			'End Sub\n' +
			'Public Sub ReadString(ByVal value As String)\n' +
			'End Sub\n' +
			'Public Sub NeedsString(ByRef value As String)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Dim longAmount As Long\n' +
			'    Dim flexible As Variant\n' +
			'    Dim obj As Object\n' +
			'    ReadValue amount\n' +
			'    ReadString obj\n' +
			'    NeedsString flexible\n' +
			'    Mutate 1\n' +
			'    Mutate (longAmount)\n' +
			'    Mutate longAmount + 1\n' +
			'    Mutate Range("A1").Value\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('uses module #Const activity when checking ByRef exactness in conditional branches', () => {
		const src =
			'#Const NegActiveBranch = True\n' +
			'#Const NegInactiveBranch = False\n' +
			'Public Sub NegSupport01_NeedsLong(ByRef value As Long)\n' +
			'End Sub\n' +
			'#If NegActiveBranch Then\n' +
			'Public Sub NegCond01_ActiveInvalidBranch()\n' +
			'    Dim amount As Integer\n' +
			'    NegSupport01_NeedsLong amount\n' +
			'End Sub\n' +
			'#Else\n' +
			'Public Sub NegCond02_InactiveAlternateBranch()\n' +
			'End Sub\n' +
			'#End If\n' +
			'#If NegInactiveBranch Then\n' +
			'Public Sub NegCond03_InactiveInvalidBranch()\n' +
			'    Dim amount As Integer\n' +
			'    NegSupport01_NeedsLong amount\n' +
			'End Sub\n' +
			'#Else\n' +
			'Public Sub NegCond04_ActiveValidBranch()\n' +
			'End Sub\n' +
			'#End If\n';
		const hits = byCode(analyzeModule(src), 'byref-argument-type-mismatch');

		expectDiagnostic(src, hits, 'byref-argument-type-mismatch', { span: 'amount' });
	});

	it('uses unique exported project signatures for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim amount As Integer\n' +
			'    Mutate amount\n' +
			'End Sub\n';
		const helpers = 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helpers },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('amount');
	});

	it('uses visible exported project variables for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
				{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'byref-argument-type-mismatch', { span: 'SharedAmount' });
	});

	it('uses module-qualified exported project variables for ByRef exactness', () => {
		const caller =
			'Public Sub T()\n' +
			'    Mutate Globals.SharedAmount\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
				{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
			], 'Caller'),
			'byref-argument-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'byref-argument-type-mismatch', {
			span: 'Globals.SharedAmount',
		});
	});

	it('keeps local shadows and ambiguous exported variables quiet for ByRef exactness', () => {
		const shadowCaller =
			'Public Sub T()\n' +
			'    Dim SharedAmount As Variant\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
			{ moduleName: 'Globals', source: 'Public SharedAmount As Integer\n' },
		], 'Caller'), 'byref-argument-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub T()\n' +
			'    Mutate SharedAmount\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'Helpers', source: 'Public Sub Mutate(ByRef value As Long)\nEnd Sub\n' },
			{ moduleName: 'GlobalsA', source: 'Public SharedAmount As Integer\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedAmount As String\n' },
		], 'Caller'), 'byref-argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn on Variant arguments whose runtime value is unknown', () => {
		const src =
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Public Sub TestInvoiceTotal()\n' +
			'    Dim value As Variant\n' +
			'    total = InvoiceTotal(value)\n' +
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
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"blah"' });
	});

	it('validates selected native VBA runtime parameter types', () => {
		const src = 'Sub T()\n    x = Left("abcdef", "bad")\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('does not validate runtime parameter types when a local shadows the runtime function', () => {
		const src =
			'Sub T()\n' +
			'    Dim Left As Variant\n' +
			'    value = Left("abcdef", "bad")\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'runtime-argument-value')).toHaveLength(0);
	});

	it('flags negative literal values for selected native VBA runtime argument bounds', () => {
		const src =
			'Sub T()\n' +
			'    a = Left$("abcdef", -1)\n' +
			'    b = Left("abcdef", -2)\n' +
			'    c = String$(-3, "x")\n' +
			'    d = String(-4, "x")\n' +
			'    e = VBA.Left$("abcdef", -5)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: ['Left$', 'Length', "Run-time error '5'"] },
			{ span: '-2' },
			{ span: '-3', message: ['String$', 'Number'] },
			{ span: '-4' },
			{ span: '-5' },
		]);
	});

	it('flags oracle-backed runtime argument bounds for Right Space and Mid', () => {
		const src =
			'Sub T()\n' +
			'    a = Right$("abcdef", -1)\n' +
			'    b = Right("abcdef", -2)\n' +
			'    c = Space$(-3)\n' +
			'    d = Space(-4)\n' +
			'    e = Mid$("abcdef", 0, 1)\n' +
			'    f = Mid("abcdef", -1, 1)\n' +
			'    g = Mid$("abcdef", 1, -5)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: ['Right$', 'Length'] },
			{ span: '-2' },
			{ span: '-3', message: ['Space$', 'Number'] },
			{ span: '-4' },
			{ span: '0', message: ['Mid$', 'Start'] },
			{ span: '-1' },
			{ span: '-5', message: 'Length' },
		]);
	});

	it('flags oracle-backed runtime argument bounds for Replace', () => {
		const src =
			'Sub T()\n' +
			'    a = Replace("abcdef", "a", "z", 0)\n' +
			'    b = Replace("abcdef", "a", "z", -1)\n' +
			'    c = Replace("aaaa", "a", "z", 1, -2)\n' +
			'    d = Replace("aaaa", "a", "z", Count:=-2, Start:=1)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '0', message: ['Replace', 'Start'] },
			{ span: '-1' },
			{ span: '-2', message: 'Count' },
			{ span: '-2' },
		]);
	});

	it('folds deterministic Const Enum and integer expressions for runtime argument bounds', () => {
		const src =
			'Private Const BadLength As Long = -1\n' +
			'Private Const GoodLength As Long = 0\n' +
			'Private Enum RuntimeArgStart\n' +
			'    EnumBadStart = 0\n' +
			'End Enum\n' +
			'Sub T()\n' +
			'    Const BadStart As Long = 1 - 1\n' +
			'    Const BadCount As Long = -1 - 1\n' +
			'    Const GoodCount As Long = 0 - 1\n' +
			'    a = Left$("abcdef", BadLength)\n' +
			'    b = Left$("abcdef", GoodLength)\n' +
			'    c = Left$("abcdef", 1 - 2)\n' +
			'    d = Right$("abcdef", 1 - 1)\n' +
			'    e = Replace("abcdef", "a", "z", BadStart)\n' +
			'    f = Replace("aaaa", "a", "z", 1, BadCount)\n' +
			'    g = Replace("aaaa", "a", "z", 1, GoodCount)\n' +
			'    h = Mid$("abcdef", EnumBadStart, 1)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: 'BadLength', message: 'is -1' },
			{ span: '1 - 2', message: 'is -1' },
			{ span: 'BadStart', message: 'is 0' },
			{ span: 'BadCount', message: 'is -2' },
			{ span: 'EnumBadStart', message: 'is 0' },
		]);
	});

	it('folds verified runtime and host constants for runtime argument bounds', () => {
		const src =
			'Sub T()\n' +
			'    a = Left$("abcdef", vbFalse - 1)\n' +
			'    b = Left$("abcdef", VBA.vbFalse - 1)\n' +
			'    c = Left$("abcdef", xlAbove - 1)\n' +
			'    d = Left$("abcdef", xlNo - 1)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-argument-value');

		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'vbFalse - 1',
			'VBA.vbFalse - 1',
			'xlAbove - 1',
		]);
	});

	it('lets source declarations shadow verified runtime and host constants for runtime argument bounds', () => {
		const src =
			'Private Const vbFalse As Long = 1\n' +
			'Sub T()\n' +
			'    Dim xlAbove As Long\n' +
			'    xlAbove = 1\n' +
			'    a = Left$("abcdef", vbFalse - 1)\n' +
			'    b = Left$("abcdef", xlAbove - 1)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'runtime-argument-value')).toHaveLength(0);
	});

	it('flags oracle-backed InStr Start values below one without flagging two-argument calls', () => {
		const src =
			'Sub T()\n' +
			'    Const BadStart As Long = 0\n' +
			'    a = InStr(0, "abcdef", "a")\n' +
			'    b = InStr(-1, "abcdef", "a")\n' +
			'    c = InStr(1, "abcdef", "a")\n' +
			'    d = InStr("abcdef", "a")\n' +
			'    e = InStr(BadStart, "abcdef", "a")\n' +
			'    f = InStr(Start:=0, String1:="abcdef", String2:="a")\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '0', message: "'Start' of 'InStr' is 0" },
			{ span: '-1', message: "'Start' of 'InStr' is -1" },
			{ span: 'BadStart', message: "'Start' of 'InStr' is 0" },
		]);
	});

	it('flags oracle-backed Chr and ChrW CharCode values outside proven bounds', () => {
		const src =
			'Sub T()\n' +
			'    Const BadChrLow As Long = -1\n' +
			'    Const BadChrHigh As Long = 256\n' +
			'    Const BadChrWHigh As Long = 65536\n' +
			'    a = Chr(-1)\n' +
			'    b = Chr(0)\n' +
			'    c = Chr(255)\n' +
			'    d = Chr(256)\n' +
			'    e = Chr(BadChrLow)\n' +
			'    f = Chr(BadChrHigh)\n' +
			'    g = ChrW(-1)\n' +
			'    h = ChrW(65535)\n' +
			'    i = ChrW(65536)\n' +
			'    j = ChrW(BadChrWHigh)\n' +
			'    k = VBA.Chr(256)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'runtime-argument-value', [
			{ span: '-1', message: "'CharCode' of 'Chr' is -1" },
			{ span: '256', message: "'CharCode' of 'Chr' is 256" },
			{ span: 'BadChrLow', message: "'CharCode' of 'Chr' is -1" },
			{ span: 'BadChrHigh', message: "'CharCode' of 'Chr' is 256" },
			{ span: '65536', message: "'CharCode' of 'ChrW' is 65536" },
			{ span: 'BadChrWHigh', message: "'CharCode' of 'ChrW' is 65536" },
			{ span: '256', message: "'CharCode' of 'Chr' is 256" },
		]);
	});

	it('folds visible cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedBadStart, 1)\n' +
			'    d = Replace("aaaa", "a", "z", 1, SharedBadCount)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Const SharedBadCount As Long = -2\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(3);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLength',
			'SharedBadStart',
			'SharedBadCount',
		]);
	});

	it('folds module-qualified cross-module Const and Enum values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedRuntimeArgs.SharedBadLength)\n' +
			'    b = Left$("abcdef", SharedRuntimeArgs.SharedGoodLength)\n' +
			'    c = Mid$("abcdef", SharedRuntimeArgs.SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Const SharedGoodLength As Long = 0\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedRuntimeArgs.SharedBadLength',
			'SharedRuntimeArgs.SharedBadStart',
		]);
	});

	it('folds exported cross-module Const values through private same-module helpers for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", SharedBadLengthFromHidden)\n' +
			'    b = Left$("abcdef", SharedGoodLengthFromHidden)\n' +
			'    c = Left$("abcdef", SharedRuntimeArgs.SharedBadLengthFromHidden)\n' +
			'    d = Left$("abcdef", SharedRuntimeArgs.SharedGoodLengthFromHidden)\n' +
			'    e = Left$("abcdef", SharedRuntimeArgs.HiddenBadLength)\n' +
			'End Sub\n';
		const shared =
			'Private Const HiddenBadLength As Long = -1\n' +
			'Private Const HiddenGoodLength As Long = 0\n' +
			'Public Const SharedBadLengthFromHidden As Long = HiddenBadLength\n' +
			'Public Const SharedGoodLengthFromHidden As Long = HiddenGoodLength\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'SharedRuntimeArgs', source: shared, moduleKind: 'standard' },
			], 'Caller'),
			'runtime-argument-value',
		);

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual([
			'SharedBadLengthFromHidden',
			'SharedRuntimeArgs.SharedBadLengthFromHidden',
		]);
	});

	it('defers hidden and ambiguous cross-module Const values for runtime argument bounds', () => {
		const caller =
			'Sub T()\n' +
			'    a = Left$("abcdef", HiddenBadLength)\n' +
			'    b = Left$("abcdef", AmbiguousLength)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'PrivateGlobals',
					source:
						'Private Const HiddenBadLength As Long = -1\n' +
						'Public Const AmbiguousLength As Long = -1\n',
				},
				{
					moduleName: 'MoreGlobals',
					source: 'Public Const AmbiguousLength As Long = 0\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('lets current-module and local Const values shadow cross-module runtime argument bounds', () => {
		const caller =
			'Private Const SharedBadLength As Long = 0\n' +
			'Sub T()\n' +
			'    Const SharedBadStart As Long = 1\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Mid$("abcdef", SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);
		expect(hits).toHaveLength(0);
	});

	it('lets source variables shadow cross-module runtime argument constants', () => {
		const caller =
			'Private SharedBadLength As Long\n' +
			'Sub T()\n' +
			'    Dim SharedBadStart As Long\n' +
			'    SharedBadLength = 1\n' +
			'    SharedBadStart = 1\n' +
			'    a = Left$("abcdef", SharedBadLength)\n' +
			'    b = Mid$("abcdef", SharedBadStart, 1)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{
					moduleName: 'SharedRuntimeArgs',
					source:
						'Public Const SharedBadLength As Long = -1\n' +
						'Public Enum SharedRuntimeStart\n' +
						'    SharedBadStart = 0\n' +
						'End Enum\n',
				},
			], 'Caller'),
			'runtime-argument-value',
		);

		expect(hits).toHaveLength(0);
	});

	it('accepts zero and unknown runtime argument values for selected native bounds', () => {
		const src =
			'Sub T()\n' +
			'    Dim count As Long\n' +
			'    a = Left$("abcdef", 0)\n' +
			'    b = Left("abcdef", count)\n' +
			'    c = String$(0, "x")\n' +
			'    d = String(count, "x")\n' +
			'    e = Right$("abcdef", 0)\n' +
			'    f = Right("abcdef", count)\n' +
			'    g = Space$(0)\n' +
			'    h = Space(count)\n' +
			'    i = Mid$("abcdef", 1, 0)\n' +
			'    j = Mid("abcdef", count, count)\n' +
			'    k = Replace("abcdef", "a", "z", 1)\n' +
			'    l = Replace("aaaa", "a", "z", 1, -1)\n' +
			'    m = Replace("aaaa", "a", "z", 1, 0)\n' +
			'    n = Replace("aaaa", "a", "z", count, count)\n' +
			'    o = Replace$("aaaa", "a", "z", 0)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'runtime-argument-value')).toHaveLength(0);
	});

	it('does not treat shadowed Left calls as native runtime argument-value checks', () => {
		const src =
			'Public Function Left(ByVal text As String, ByVal count As Long) As String\n' +
			'End Function\n' +
			'Sub T()\n' +
			'    Dim localValue As Long\n' +
			'    Dim Left As Long\n' +
			'    a = Left("abcdef", -1)\n' +
			'    b = VBA.Left$("abcdef", -1)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: '-1' });
	});

	it('flags plainly invalid literal CDate conversions', () => {
		const src =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-conversion-value');

		expectDiagnostics(src, hits, 'runtime-conversion-value', [
			{ span: '"not a date"', message: ['CDate', "Run-time error '13'"] },
			{ span: '""', message: 'VBA.CDate' },
		]);
	});

	it('keeps date-looking, localized, variable, and non-CDate conversions quiet', () => {
		const src =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Dim text As String\n' +
			'    Value = CDate("1/2/2020")\n' +
			'    Value = CDate("March")\n' +
			'    Value = CDate("Mar 1")\n' +
			'    Value = CDate("März")\n' +
			'    Value = CDate(text)\n' +
			'    Value = DateValue("not a date")\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'runtime-conversion-value')).toHaveLength(0);
	});

	it('does not treat shadowed CDate calls as native conversion checks', () => {
		const src =
			'Sub T()\n' +
			'    Dim CDate As Variant\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("not a date")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'runtime-conversion-value');

		expectDiagnostic(src, hits, 'runtime-conversion-value', { span: '"not a date"' });
	});

	it('does not treat project-visible source names as native conversion checks', () => {
		const globals = 'Public Const CDate As Long = 1\n';
		const caller =
			'Sub T()\n' +
			'    Dim Value As Date\n' +
			'    Value = CDate("not a date")\n' +
			'    Value = VBA.CDate("not a date")\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(
			caller,
			[
				{ moduleName: 'Globals', moduleKind: 'standard', source: globals },
				{ moduleName: 'Caller', moduleKind: 'standard', source: caller },
			],
			'Caller',
		);
		const hits = byCode(diagnostics, 'runtime-conversion-value');

		expectDiagnostic(caller, hits, 'runtime-conversion-value', { span: '"not a date"' });
	});

	it('keeps Close of a literal file number quiet', () => {
		const src = 'Option Explicit\nSub T()\n    Close #99\nEnd Sub\n';
		expect(analyzeModule(src)).toEqual([]);
	});

	it('keeps string-suffixed Left and Right intrinsics quiet', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Debug.Print Left$("abcdef", 2)\n' +
			'    Debug.Print Left("abcdef", 2)\n' +
			'    Debug.Print Right$("abcdef", 2)\n' +
			'    Debug.Print Right("abcdef", 2)\n' +
			'End Sub\n';
		expect(analyzeModule(src)).toEqual([]);
	});

	it('does not infer runtime parameter types from display names', () => {
		const src = 'Sub T()\n    Randomize "bad"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('does not warn when a type is unknown or Variant-like', () => {
		const src = 'Sub T()\n    x = Format("blah", "0.00")\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(0);
	});

	it('uses same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			severity: 'error',
			span: 'MakeLabel',
		});
	});

	it('uses string-suffixed runtime aliases as argument types', () => {
		const src =
			'Option Explicit\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Left$("abcdef", 2)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', { span: 'Left$' });
	});

	it('uses verified runtime and host constants as argument value types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject vbFalse\n' +
			'    NeedsObject VBA.vbFalse\n' +
			'    NeedsObject xlAbove\n' +
			'    NeedsObject vbNullString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');

		expectDiagnostics(src, hits, 'argument-object-type-mismatch', [
			{ span: 'vbFalse', message: 'vbFalse As VbTriState' },
			{ span: 'vbFalse', message: 'VBA.vbFalse As VbTriState' },
			{ span: 'xlAbove', message: 'xlAbove As Constants' },
			{ span: 'vbNullString', message: 'vbNullString As String' },
		]);
	});

	it('lets source declarations shadow verified external constants for argument value types', () => {
		const src =
			'Public Function vbFalse() As Object\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim xlAbove As Object\n' +
			'    NeedsObject vbFalse\n' +
			'    NeedsObject xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses parameterless Function and Property Get references as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Property Get CurrentLabel() As String\n' +
			'End Property\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel\n' +
			'    NeedsObject (CurrentLabel)\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-object-type-mismatch', [
			{ span: 'MakeLabel', message: 'MakeLabel As String' },
			{ span: 'CurrentLabel', message: 'CurrentLabel As String' },
		]);
	});

	it('keeps the current Function return variable available for bare argument inference', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Function CurrentLabel(ByVal index As Long) As String\n' +
			'    NeedsObject CurrentLabel\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');

		expectDiagnostic(src, hits, 'argument-object-type-mismatch', { span: 'CurrentLabel' });
	});

	it('uses module-qualified parameterless project Function references as argument types', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject Labels.MakeLabel\n' +
			'End Sub\n';
		const labels = 'Public Function MakeLabel() As String\nEnd Function\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Labels', source: labels },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('does not infer ambiguous parameterless project Function references', () => {
		const caller =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject MakeLabel\n' +
			'End Sub\n';
		const first = 'Public Function MakeLabel() As String\nEnd Function\n';
		const second = 'Public Function MakeLabel() As String\nEnd Function\n';
		expect(byCode(analyzeProjectModule(caller, [
			{ moduleName: 'FirstLabels', source: first },
			{ moduleName: 'SecondLabels', source: second },
		], 'Caller'), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses generated host member signatures for scalar argument types', () => {
		const src =
			'Sub T()\n' +
			'    Call Application.DeleteCustomList("bad")\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.NeedsObject("bad")\n' +
			'End Sub\n';
		const person =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-object-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-object-type-mismatch', { span: '"bad"' });
	});

	it('uses parenless source-backed class member signatures for argument types', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save "bad"\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Count As Long)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-type-mismatch',
		);
		expectDiagnostic(caller, hits, 'argument-type-mismatch', { span: '"bad"' });
	});

	it('uses nested same-module function return types as argument types', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Function EchoLabel(ByVal value As String) As String\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject EchoLabel(MakeLabel())\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'EchoLabel',
		});
	});

	it('uses curated runtime conversion function return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    NeedsObject CStr(123)\n' +
			'    NeedsObject CDbl("1")\n' +
			'    NeedsObject CCur(1)\n' +
			'    NeedsObject CLng(1)\n' +
			'    NeedsObject CBool("True")\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-object-type-mismatch');
		expect(hits).toHaveLength(5);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'CStr',
			'CDbl',
			'CCur',
			'CLng',
			'CBool',
		]);
	});

	it('does not infer runtime return types when a local shadows the runtime function', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim Format\n' +
			'    NeedsObject Format(123)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses obvious numeric arithmetic expression return types as argument types', () => {
		const src =
			'Public Function TaxRate() As Double\n' +
			'End Function\n' +
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim subtotal As Currency\n' +
			'    Dim fee As Double\n' +
			'    NeedsObject subtotal + fee * TaxRate()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'subtotal + fee * TaxRate()',
		});
	});

	it('does not infer arithmetic expressions with unknown, Variant, or string operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim amount As Double\n' +
			'    Dim flexible As Variant\n' +
			'    Dim label As String\n' +
			'    NeedsObject amount + flexible\n' +
			'    NeedsObject amount + UnknownValue\n' +
			'    NeedsObject amount + label\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});

	it('uses string concatenation expression return types as argument types', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim amount As Double\n' +
			'    NeedsObject prefix & amount & CStr(123)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-object-type-mismatch', {
			span: 'prefix & amount & CStr(123)',
		});
	});

	it('does not infer string concatenation expressions with unknown or Variant operands', () => {
		const src =
			'Public Sub NeedsObject(ByVal item As Object)\n' +
			'End Sub\n' +
			'Public Sub T()\n' +
			'    Dim prefix As String\n' +
			'    Dim flexible As Variant\n' +
			'    NeedsObject prefix & flexible\n' +
			'    NeedsObject prefix & UnknownValue\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-object-type-mismatch')).toHaveLength(0);
	});
});
