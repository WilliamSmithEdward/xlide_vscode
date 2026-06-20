// Diagnostics tests: undeclared rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import {
	analyzeProjectModule,
	projectOptions,
	visibleProjectIdentifiers,
	visibleProjectProcedures,
} from './helpers';

describe('analyzeModule - Option Explicit', () => {
	it('warns when a code module omits Option Explicit', () => {
		const src = 'Sub T()\n    Dim x As Long\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'option-explicit-missing');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('warning');
	});

	it('anchors as a zero-width marker at the module top, not a full first line', () => {
		// A full first-line range collides with a line-1 error (e.g. the
		// missing-block-closer on an unclosed first procedure), letting the
		// warning obscure the red squiggle. Keep it a zero-width top marker.
		const src = 'Sub test()\n    Dim x As Long\n'; // unclosed Sub, no Option Explicit
		const hit = byCode(analyzeModule(src), 'option-explicit-missing')[0];
		expect(hit.span).toEqual({ start: 0, end: 0 });
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
		const src = 'Option Explicit\nSub T()\n    MissingProc\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				knownProcedures: new Set<string>(),
				severityOverrides: { 'unknown-call': 'warning' },
			}),
			'unknown-call',
		);
		expect(hits[0].severity).toBe('warning');
	});

	it('ignores severity overrides that violate rule guardrails', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severityOverrides: { 'option-explicit-missing': 'error' } }),
			'option-explicit-missing',
		);
		expect(hits[0].severity).toBe('warning');
	});

	it('flags an undeclared bare assignment target when Option Explicit is present', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    notDeclared = ThisWorkbook.CanCheckIn()\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('notDeclared');
		expect(hits[0].severity).toBe('error');
	});

	it('allows implicit Variant assignment when Option Explicit is absent', () => {
		const src = 'Sub T()\n    notDeclared = ThisWorkbook.CanCheckIn()\nEnd Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts declared local, module, parameter, and Function-return assignment targets', () => {
		const src =
			'Option Explicit\n' +
			'Private moduleValue As Long\n' +
			'Function T(ByVal arg As Long) As Long\n' +
			'    Dim localValue As Long\n' +
			'    localValue = 1\n' +
			'    moduleValue = localValue\n' +
			'    arg = moduleValue\n' +
			'    T = arg\n' +
			'End Function\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('recognizes hidden built-ins (pointer fns, byte-string fns, vbLongLong) under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim x As Long, s As String, o As Object\n' +
			'    Debug.Print VarPtr(x), StrPtr(s), ObjPtr(o)\n' +
			'    Debug.Print LeftB(s, 2), RightB(s, 2), MidB(s, 1, 2), LenB(s), AscB(s), ChrB(65)\n' +
			'    Debug.Print vbLongLong\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('flags an undeclared Set assignment target under Option Explicit', () => {
		const src = 'Option Explicit\nSub T()\n    Set notDeclared = ActiveSheet\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('notDeclared');
	});

	it('accepts known member and indexed assignment receivers under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Range("A1").Value = 1\n' +
			'    declaredArr(1) = 2\n' +
			'End Sub\n';
		const knownIdentifiers = new Set<string>(['declaredarr']);
		expect(
			byCode(analyzeModule(src, { knownIdentifiers }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('flags undeclared identifiers read from assignment right-hand sides and call arguments', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim total As Long\n' +
			'    total = missingValue + 1\n' +
			'    MsgBox missingMessage\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['missingValue', 'missingMessage']);
	});

	it('flags undeclared identifiers in control-flow expressions and loop targets', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    If missingCondition Then Beep\n' +
			'    For i = 1 To maxCount\n' +
			'    Next i\n' +
			'    With missingObject\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'missingCondition',
			'i',
			'maxCount',
			'missingObject',
		]);
	});

	it('flags undeclared member receivers and indexed bases under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    obj.Value = 1\n' +
			'    arr(ix) = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['obj', 'arr', 'ix']);
	});

	it('accepts declared reads, exported globals, and exported enum members', () => {
		const src =
			'Option Explicit\n' +
			'Sub T(ByVal arg As Long)\n' +
			'    Dim localValue As Long\n' +
			'    localValue = arg + sharedValue + SharedOnly\n' +
			'End Sub\n';
		const knownIdentifiers = visibleProjectIdentifiers(
			[
				{ moduleName: 'Caller', source: src },
				{
					moduleName: 'Globals',
					source:
						'Public sharedValue As Long\n' +
						'Public Enum SharedMode\n    SharedOnly\nEnd Enum\n',
				},
			],
			'Caller',
		);
		expect(byCode(analyzeModule(src, { knownIdentifiers }), 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts built-in VBA and Excel constants under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    MsgBox "hi", vbOKOnly\n' +
			'    Application.DisplayAlerts = vbFalse\n' +
			'    ActiveSheet.Range("A1").End(xlUp).Select\n' +
			'    Dim dashStyle As Long\n' +
			'    dashStyle = msoLineDash\n' +
			'    Err.Raise vbObjectError + 1, "Person.Age", "Age must be between 0 and 120"\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts the VBA namespace and compare aliases under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim seen As Scripting.Dictionary\n' +
			'    Set seen = New Scripting.Dictionary\n' +
			'    seen.CompareMode = TextCompare\n' +
			'    MsgBox VBA.CStr(1)\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts Erl in line-numbered error handlers under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'10 Dim message As String\n' +
			'20 On Error GoTo Handler\n' +
			'30 Err.Raise 5\n' +
			'40 Exit Sub\n' +
			'Handler:\n' +
			'50 message = "Error on line " & Erl\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>() }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('handles line-numbered assignment targets under Option Explicit', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    Dim declared As Long\n' +
			'    Dim obj As Object\n' +
			'10 declared = 1\n' +
			'20 Set obj = ActiveSheet\n' +
			'30 missing = 1\n' +
			'40 Set missingObj = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { knownIdentifiers: new Set<string>() }),
			'undeclared-variable',
		);
		expectDiagnostics(src, hits, 'undeclared-variable', [
			{ span: 'missing', message: 'assigning to it' },
			{ span: 'missingObj', message: 'assigning to it' },
		]);
	});

	it('does not flag type names, labels, named-argument labels, or unknown external-style calls as reads', () => {
		const src =
			'Option Explicit\n' +
			'Sub T()\n' +
			'done:\n' +
			'    Set p = New Person\n' +
			'    If TypeOf p Is Person Then GoTo done\n' +
			'    MsgBox Prompt:=p\n' +
			'    MaybeExternal missingArg\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { knownIdentifiers: new Set<string>(['p']) }), 'undeclared-variable'),
		).toHaveLength(0);
	});

	it('accepts exported standard-module globals visible through the project index', () => {
		const src = 'Option Explicit\nSub T()\n    sharedValue = 1\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(
			src,
			[
				{ moduleName: 'Globals', source: 'Public sharedValue As Long\n' },
			],
			'Caller',
		);
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('uses project-visible source bindings as Option Explicit declarations', () => {
		const caller =
			'Option Explicit\n' +
			'Sub T()\n' +
			'    SharedValue = 1\n' +
			'    Debug.Print SharedOnly\n' +
			'End Sub\n';
		const globals =
			'Public SharedValue As Long\n' +
			'Public Enum SharedMode\n' +
			'    SharedOnly\n' +
			'End Enum\n';
		const options = projectOptions([
			{ moduleName: 'Caller', source: caller },
			{ moduleName: 'Globals', source: globals },
		], 'Caller');
		const diagnostics = analyzeModule(caller, {
			moduleName: 'Caller',
			knownIdentifiers: new Set<string>(),
			projectVisibleSymbols: options.projectVisibleSymbols,
		});

		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts module-qualified Function calls on a Set assignment right-hand side', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    Set item = Factories.MakeItem()\n' +
			'    Set item = Factories.MakeOther\n' +
			'End Sub\n';
		const factories =
			'Public Function MakeItem() As Object\n' +
			'End Function\n' +
			'Public Function MakeOther() As Object\n' +
			'End Function\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Factories', source: factories },
		], 'Caller');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('accepts module-qualified standard-module values on expression right-hand sides', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = Globals.SharedConst + Settings.SomePublicValue + Globals.SharedOnly\n' +
			'End Sub\n';
		const globals =
			'Public Const SharedConst As Long = 1\n' +
			'Public Enum SharedMode\n' +
			'    SharedOnly\n' +
			'End Enum\n';
		const settings = 'Public SomePublicValue As Long\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: globals },
			{ moduleName: 'Settings', source: settings },
		], 'Caller');
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
	});

	it('keeps unknown module-qualified value qualifiers visible under Option Explicit', () => {
		const caller =
			'Option Explicit\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = MissingModule.SharedConst\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [], 'Caller');
		const hits = byCode(diagnostics, 'undeclared-variable');
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('MissingModule');
	});

	it('can be switched off', () => {
		const src = 'Sub T()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(src, { severityOverrides: { 'option-explicit-missing': 'off' } }),
			'option-explicit-missing',
		);
		expect(hits).toHaveLength(0);
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

	it('uses project visibility for calls to other standard modules', () => {
		const caller = 'Sub Main()\n    DoWork\n    Secret\nEnd Sub\n';
		const helpers =
			'Public Sub DoWork()\nEnd Sub\n' +
			'Private Sub Secret()\nEnd Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Helpers', source: helpers },
			], 'Caller'),
			'unknown-call',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Secret');
	});

	it('does not treat class-module methods as bare cross-module procedures', () => {
		const caller = 'Sub Main()\n    Save\nEnd Sub\n';
		const customer = 'Public Sub Save()\nEnd Sub\n';
		const hits = byCode(
			analyzeModule(caller, {
				moduleName: 'Caller',
				knownProcedures: visibleProjectProcedures([
					{ moduleName: 'Caller', source: caller },
					{ moduleName: 'Customer', moduleKind: 'class', source: customer },
				], 'Caller'),
			}),
			'unknown-call',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('Save');
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

	it('does not report unknown-call for a dangling member-access dot', () => {
		const src = 'Sub Main()\n    aadf.\nEnd Sub\n';
		expect(byCode(analyzeModule(src, opts), 'unknown-call')).toHaveLength(0);
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
		expect(hits[0].data?.createProcedureStub).toMatchObject({
			procedureName: 'Frobnicate',
			edit: {
				span: { start: src.length, end: src.length },
				newText:
					'\nPrivate Sub Frobnicate(ByVal arg1 As Variant, ByVal arg2 As Variant, ByVal arg3 As Variant)\n' +
					'End Sub\n',
			},
		});
	});

	it('flags an unknown explicit Call statement', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
		expect(hits[0].data?.createProcedureStub?.edit.newText).toContain(
			'Private Sub DoesNotExist(ByVal arg1 As Variant)\n',
		);
	});

	it('flags unknown calls after numeric line labels', () => {
		const src =
			'Sub Main()\n' +
			'10 DoesNotExist 1\n' +
			'20 Call AlsoMissing(2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['DoesNotExist', 'AlsoMissing']);
	});

	it('does not attach procedure-stub metadata for omitted argument slots', () => {
		const src = 'Sub Main()\n    Call DoesNotExist(1, , 3)\nEnd Sub\n';
		const hits = byCode(analyzeModule(src, opts), 'unknown-call');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('DoesNotExist');
		expect(hits[0].data?.createProcedureStub).toBeUndefined();
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
	it('flags a bare local variable statement rejected by VBE Compile', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'non-callable-call', {
			severity: 'error',
			span: 'testStr',
		});
	});

	it('flags a local variable used with call arguments', () => {
		const src = 'Sub Main()\n    Dim testStr As String\n    testStr "hello"\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'non-callable-call', {
			severity: 'error',
			span: 'testStr',
		});
	});

	it('flags an explicit Call to a parameter', () => {
		const src = 'Sub Main(ByVal testStr As String)\n    Call testStr\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'non-callable-call', { span: 'testStr' });
	});

	it('does not flag indexed object variables feeding member access', () => {
		const src =
			'Sub Main()\n' +
			'    Dim buckets As Object\n' +
			'    Call buckets("ready").Add(1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'non-callable-call')).toHaveLength(0);
	});

	it('flags project-visible exported non-callables used as call statements', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedValue\n' +
			'    MaxValue 1\n' +
			'    Call SharedMode\n' +
			'    Active\n' +
			'End Sub\n';
		const helpers =
			'Public SharedValue As Long\n' +
			'Public Const MaxValue As Long = 10\n' +
			'Public Enum SharedMode\n' +
			'    Active\n' +
			'End Enum\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller');
		const hits = byCode(diagnostics, 'non-callable-call');

		expectDiagnostics(caller, hits, 'non-callable-call', [
			{ span: 'SharedValue', message: 'module variable' },
			{ span: 'MaxValue', message: 'constant' },
			{ span: 'SharedMode', message: 'enum type' },
			{ span: 'Active', message: 'enum member' },
		]);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('keeps project non-callables silent when a visible procedure shares the name', () => {
		const caller = 'Sub Main()\n    SharedName\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedName As Long\n' },
			{ moduleName: 'Helpers', source: 'Public Sub SharedName()\nEnd Sub\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'non-callable-call')).toHaveLength(0);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('uses local and module precedence before exported callables for call statements', () => {
		const caller =
			'Private SharedName As Long\n' +
			'Sub Main()\n' +
			'    Dim DoWork As Long\n' +
			'    DoWork\n' +
			'    SharedName\n' +
			'End Sub\n';
		const helpers =
			'Public Sub DoWork()\nEnd Sub\n' +
			'Public Sub SharedName()\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', source: helpers },
		], 'Caller');
		const hits = byCode(diagnostics, 'non-callable-call');

		expectDiagnostics(caller, hits, 'non-callable-call', [
			{ span: 'DoWork', message: 'local variable' },
			{ span: 'SharedName', message: 'module variable' },
		]);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('keeps duplicate project non-callables silent instead of unknown', () => {
		const caller = 'Sub Main()\n    SharedName\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedName As Long\n' },
			{ moduleName: 'OtherGlobals', source: 'Public Const SharedName As Long = 1\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'non-callable-call')).toHaveLength(0);
		expect(byCode(diagnostics, 'unknown-call')).toHaveLength(0);
	});

	it('does not flag callable procedures or runtime statements', () => {
		const src = 'Sub Main()\n    Helper "ok"\n    Beep\nEnd Sub\nSub Helper(ByVal s As String)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'non-callable-call')).toHaveLength(0);
	});
});
