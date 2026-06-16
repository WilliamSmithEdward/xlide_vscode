// Diagnostics tests: callArity rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule, projectClassMembers, projectProcedures } from './helpers';

describe('analyzeModule - argument count', () => {
	it('flags too few arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'Greet' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'b',
			edit: {
				span: {
					start: src.indexOf('"Ann"') + '"Ann"'.length,
					end: src.indexOf('"Ann"') + '"Ann"'.length,
				},
				newText: ', TODO_b',
			},
		});
	});

	it('checks argument counts for calls after numeric line labels', () => {
		const src =
			'Sub Main()\n' +
			'10 Greet "Ann"\n' +
			'20 Call Greet("Ann")\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'argument-count', [
			{ span: 'Greet', message: 'expected 2 arguments' },
			{ span: 'Greet', message: 'expected 2 arguments' },
		]);
	});

	it('flags too many arguments to a same-module Sub', () => {
		const src =
			'Sub Main()\n' +
			'    Greet "Ann", "Bob", "Cat"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count');
	});

	it('flags a lone-identifier call that omits required arguments', () => {
		const src =
			'Sub Main()\n' +
			'    Greet\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count');
	});

	it('flags a runtime function call that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox()\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'MsgBox' });
		const argStart = src.indexOf('MsgBox(') + 'MsgBox('.length;
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: argStart,
					end: argStart,
				},
				newText: 'TODO_Prompt',
			},
		});
		expect(byCode(analyzeModule(src), 'call-statement-forbids-parens')).toHaveLength(0);
	});

	it('does not arity-check runtime signatures when source names shadow them', () => {
		const src =
			'Private Const Format = 1\n' +
			'Sub Main()\n' +
			'    Dim MsgBox\n' +
			'    a = Format()\n' +
			'    b = MsgBox()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('flags a parenless runtime function statement that omits required arguments', () => {
		const src = 'Sub Main()\n    MsgBox\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: 'MsgBox' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Prompt',
			edit: {
				span: {
					start: src.indexOf('MsgBox') + 'MsgBox'.length,
					end: src.indexOf('MsgBox') + 'MsgBox'.length,
				},
				newText: ' TODO_Prompt',
			},
		});
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
		expectDiagnostic(src, hits, 'argument-count', { severity: 'error', span: ',' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'TaxRate',
			edit: {
				span: {
					start: src.indexOf(', )'),
					end: src.indexOf(', )') + 2,
				},
				newText: ', TODO_TaxRate',
			},
		});
	});

	it('flags an omitted leading required argument in an expression call', () => {
		const src =
			'Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency\n' +
			'End Function\n' +
			'Sub Main()\n' +
			'    total2 = InvoiceTotal(, 0.08)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expectDiagnostic(src, hits, 'argument-count', { span: ',' });
		expect(hits[0].data?.missingRequiredArgumentPlaceholder).toEqual({
			parameterName: 'Subtotal',
			edit: {
				span: {
					start: src.indexOf('(,') + 1,
					end: src.indexOf('(,') + 3,
				},
				newText: 'TODO_Subtotal, ',
			},
		});
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

	it('treats CallByName trailing args as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    CallByName target, "Run", VbMethod, a0, a1, a2, a3\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('treats VBA Array ArgList as a runtime ParamArray', () => {
		const src =
			'Sub Main()\n' +
			'    Dim monthNames As Variant\n' +
			'    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _\n' +
			'                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", _\n' +
			'                       "Jan+1", "Feb+1", "Mar+1", "Apr+1", "May+1", _\n' +
			'                       "Jun+1", "Jul+1", "Aug+1")\n' +
			'    monthNames = Array()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('validates named-argument names but not the count', () => {
		const src =
			'Sub Main()\n' +
			'    Greet who:="Ann"\n' +
			'End Sub\n' +
			'Sub Greet(ByVal a As String, ByVal b As String)\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'who' });
	});

	it('flags duplicate named arguments', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs alpha:=1, alpha:=2\n' +
			'    NamedArgs alpha:=1, ALPHA:=2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');

		expectDiagnostics(src, hits, 'argument-count', [{ span: 'alpha' }, { span: 'ALPHA' }]);
	});

	it('flags a positional argument after a named argument (PCEC_008)', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs alpha:=1, 2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('2');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a positional after a named argument in a parenthesized function call', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Long\n' +
			'    r = FNamed(alpha:=1, 2)\n' +
			'End Sub\n' +
			'Function FNamed(ByVal alpha As Long, ByVal beta As Long) As Long\nEnd Function\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('2');
	});

	it('reports positional-after-named only once per call', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs alpha:=1, 2, 3\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'argument-count');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('2');
	});

	it('accepts positional arguments followed by named arguments (legal ordering)', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs 1, beta:=2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('accepts all-named arguments (no positional after a named one)', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs alpha:=1, beta:=2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not flag a plain positional call with no named arguments', () => {
		const src =
			'Sub Main()\n' +
			'    NamedArgs 1, 2\n' +
			'End Sub\n' +
			'Sub NamedArgs(ByVal alpha As Long, ByVal beta As Long)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
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

	it('does not arity-check runtime statements whose curated signature has no parameter list', () => {
		const src = 'Sub Main()\n    Randomize 1\nEnd Sub\n';
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

	it('uses a unique exported project signature for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const helper =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helper },
			], 'Caller'),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'PrintTotal' });
	});

	it('validates same-module Declare argument counts', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Sleep' });
	});

	it('uses the active conditional Declare signature for same-module calls', () => {
		const src =
			'#If VBA7 Then\n' +
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'#Else\n' +
			'Private Declare Sub Sleep Lib "kernel32" ()\n' +
			'#End If\n' +
			'Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const vba7Hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: true } },
			}),
			'argument-count',
		);
		expect(vba7Hits).toHaveLength(1);
		expect(spanText(src, vba7Hits[0])).toBe('Sleep');

		const legacyHits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'argument-count',
		);
		expect(legacyHits).toHaveLength(0);
	});

	it('uses exported project Declare signatures for cross-module argument count', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Sleep\n' +
			'End Sub\n';
		const nativeApi =
			'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'NativeApi', source: nativeApi },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Sleep' });
	});

	it('does not arity-check ambiguous exported project signatures', () => {
		const caller =
			'Public Sub Main()\n' +
			'    PrintTotal 100\n' +
			'End Sub\n';
		const first = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const second =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'First', source: first },
						{ moduleName: 'Second', source: second },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('uses a module-qualified project signature even when the bare name is ambiguous', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				projectProcedures: projectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Helpers', source: helpers },
					{ moduleName: 'Alternate', source: alternate },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'PrintTotal' });
	});

	it('keeps module-qualified project diagnostics stable under module order changes', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.PrintTotal 100\n' +
			'End Sub\n';
		const helpers =
			'Public Sub PrintTotal(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		const alternate = 'Public Sub PrintTotal(ByVal amount As Currency)\nEnd Sub\n';
		const ordered = projectProcedures([
			{ moduleName: 'Caller', source: caller },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Alternate', source: alternate },
		]);
		const reversed = projectProcedures([
			{ moduleName: 'Alternate', source: alternate },
			{ moduleName: 'Helpers', source: helpers },
			{ moduleName: 'Caller', source: caller },
		]);
		const messages = (projectProcedures: ReturnType<ProjectIndex['procedureSignatures']>) =>
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures,
				}),
				'argument-count',
			).map((hit) => `${spanText(caller, hit)}:${hit.message}`);
		expect(messages(ordered)).toEqual(messages(reversed));
	});

	it('does not validate a module-qualified private project procedure', () => {
		const caller =
			'Public Sub Main()\n' +
			'    Helpers.Hidden 100\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden(ByVal amount As Currency, ByVal caption As String)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(caller, {
					moduleName: 'Caller',
					projectProcedures: projectProcedures([
						{ moduleName: 'Caller', source: caller },
						{ moduleName: 'Helpers', source: helpers },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('flags missing required arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Set wb = Workbooks.Open()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Open' });
	});

	it('flags extra arguments on generated host member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Call Application.Calculate(1)\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Calculate' });
	});

	it('flags missing required arguments on host members with fallback signatures', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = ActiveSheet.Range()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('does not treat collection indexing as member-call arity', () => {
		const src =
			'Sub Main()\n' +
			'    Dim r As Range\n' +
			'    Set r = Workbooks(1).Sheets(1).Range()\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('accepts correct host member argument counts', () => {
		const src =
			'Sub Main()\n' +
			'    Dim wb As Workbook\n' +
			'    Application.Calculate\n' +
			'    Call Application.Calculate()\n' +
			'    Set wb = Workbooks.Open("Book1.xlsx")\n' +
			'    ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('does not treat zero-argument property result indexing as member-call arity', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Dim child As Object\n' +
			'    Set child = p.Children(1)\n' +
			'End Sub\n';
		const person =
			'Public Property Get Children() As Collection\n' +
			'End Property\n';
		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Person', moduleKind: 'class', source: person },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('accepts Excel collection default Item calls returned from properties', () => {
		const src =
			'Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Dim loTests As ListObject\n' +
			'    Dim lcPassed As ListColumn\n' +
			'    Set ws = ActiveSheet\n' +
			'    Set loTests = ws.ListObjects("Tests")\n' +
			'    Set lcPassed = loTests.ListColumns("Passed")\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'argument-count')).toHaveLength(0);
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
		expect(diagnostics.filter((diag) => diag.severity === 'error').map((diag) => ({
			code: diag.code,
			text: spanText(src, diag),
			message: diag.message,
		}))).toEqual([]);
	});

	it('flags missing required arguments on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    Call p.Save()\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Save' });
	});

	it('honours ParamArray on source-backed class member calls', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim cb As Callback\n' +
			'    cb.Run first, second, third\n' +
			'End Sub\n';
		const callback =
			'Public Function Run(ParamArray params() As Variant) As Variant\n' +
			'End Function\n';

		expect(
			byCode(
				analyzeModule(caller, {
					projectClassMembers: projectClassMembers([
						{ moduleName: 'Callback', moduleKind: 'class', source: callback },
					]),
				}),
				'argument-count',
			),
		).toHaveLength(0);
	});

	it('validates parenless source-backed class member call statements', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim p As Person\n' +
			'    p.Save\n' +
			'    p.Save "ok"\n' +
			'End Sub\n';
		const person =
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(caller, hits, 'argument-count', { span: 'Save' });
	});

	it('validates parenless generated host member call statements', () => {
		const src =
			'Sub Main()\n' +
			'    ActiveSheet.Range\n' +
			'    ActiveSheet.Range "A1"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Range' });
	});

	it('validates required arguments on runtime object member calls', () => {
		const src =
			'Sub Main()\n' +
			'    Err.Raise\n' +
			'    Err.Raise vbObjectError + 1\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Raise' });
	});

	it('accepts Debug.Print output lists and validates Debug.Assert arity', () => {
		const src =
			'Sub Main()\n' +
			'    Debug.Print "value", 1, True\n' +
			'    Debug.Assert\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'argument-count', { span: 'Assert' });
	});

	it('flags missing required arguments on current class Me member calls', () => {
		const src =
			'Public Sub Main()\n' +
			'    Call Me.Save()\n' +
			'End Sub\n' +
			'Public Sub Save(ByVal Caption As String)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'argument-count',
		);
		expectDiagnostic(src, hits, 'argument-count', { span: 'Save' });
	});
});
