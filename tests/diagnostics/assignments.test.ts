// Diagnostics tests: assignments rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule, type HostObjectModel } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule, projectClassMembers, projectMemberSurfaces } from './helpers';

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

	it('uses local shadows before module-level Const declarations', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n' +
			'    Dim MAX As Long\n' +
			'    MAX = 5\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'const-assignment')).toHaveLength(0);
	});

	it('flags assigning to a visible exported Const', () => {
		const caller = 'Sub T()\n    SharedMax = 5\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source: 'Public Const SharedMax As Long = 10\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'const-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedMax');
	});

	it('keeps ambiguous exported Const assignments quiet', () => {
		const caller = 'Sub T()\n    SharedMax = 5\nEnd Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'GlobalsA',
				source: 'Public Const SharedMax As Long = 10\n',
			},
			{
				moduleName: 'GlobalsB',
				source: 'Public Const SharedMax As Long = 20\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'const-assignment')).toHaveLength(0);
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

	it('flags assigning to a Const after a numeric line label', () => {
		const src =
			'Const MAX As Long = 10\n' +
			'Sub T()\n' +
			'10 MAX = 5\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'const-assignment');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MAX');
	});
});

describe('analyzeModule - assignment type validation', () => {
	it('errors on a nonnumeric string literal assigned to a numeric variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    total = "blah"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', {
			severity: 'error',
			span: '"blah"',
		});
	});

	it('errors on intrinsic CVErr Error Variants assigned to scalar variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim first As Long\n' +
			'    Dim second As Long\n' +
			'    Dim flexible As Variant\n' +
			'    first = CVErr(2015)\n' +
			'    second = VBA.CVErr(2015)\n' +
			'    flexible = CVErr(2015)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-type-mismatch', [
			{ span: 'CVErr(2015)' },
			{ span: 'VBA.CVErr(2015)' },
		]);
	});

	it('errors on Null assigned to scalar variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim number As Long\n' +
			'    Dim message As String\n' +
			'    Dim flexible As Variant\n' +
			'    number = Null\n' +
			'    message = Null\n' +
			'    flexible = Null\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-type-mismatch', [{ span: 'Null' }, { span: 'Null' }]);
	});

	it('lets a source function named CVErr shadow the intrinsic in assignment expressions', () => {
		const src =
			'Private Function CVErr(ByVal errorNumber As Long) As Long\n' +
			'    CVErr = errorNumber\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim value As Long\n' +
			'    value = CVErr(2015)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('errors on a non-Boolean string literal assigned to a Boolean variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim enabled As Boolean\n' +
			'    enabled = "maybe"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', { span: '"maybe"' });
	});

	it('errors on decimal literals outside Byte and Integer assignment bounds', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim small As Byte\n' +
			'    Dim count As Integer\n' +
			'    small = 0\n' +
			'    small = 255\n' +
			'    small = 256\n' +
			'    small = -1\n' +
			'    count = -32768\n' +
			'    count = 32767\n' +
			'    count = 32768\n' +
			'    count = -32769\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'assignment-type-mismatch', [
			{ span: '256', message: ['Byte', "Run-time error '6'"] },
			{ span: '-1', message: ['Byte', "Run-time error '6'"] },
			{ span: '32768', message: ['Integer', "Run-time error '6'"] },
			{ span: '-32769', message: ['Integer', "Run-time error '6'"] },
		]);
	});

	it('errors on decimal literals outside Long assignment bounds', () => {
		// VBE oracle: long_assignment_overflow_literal_runtime /
		// long_assignment_below_min_literal_runtime compile then raise Run-time
		// error '6': Overflow; the 2147483647 / -2147483648 controls run clean.
		const src =
			'Public Sub T()\n' +
			'    Dim n As Long\n' +
			'    n = 2147483647\n' +
			'    n = 2147483648\n' +
			'    n = -2147483648\n' +
			'    n = -2147483649\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'assignment-type-mismatch', [
			{ span: '2147483648', message: ['Long', "Run-time error '6'"] },
			{ span: '-2147483649', message: ['Long', "Run-time error '6'"] },
		]);
	});

	it('does not over-flag Long bounds (hex, LongLong, and unrepresentable literals stay silent)', () => {
		// No-FP guards: hex literals carry no numericValue (bit-pattern semantics),
		// LongLong is intentionally outside the bounds table, and literals beyond
		// JS safe-integer range cannot be proven so XLIDE stays quiet.
		const src =
			'Public Sub T()\n' +
			'    Dim n As Long\n' +
			'    Dim big As LongLong\n' +
			'    n = 2000000000\n' +
			'    n = &H7FFFFFFF\n' +
			'    n = &HFFFFFFFF\n' +
			'    big = 5000000000\n' +
			'    n = 99999999999999999999\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('errors on whole-number decimal literals outside Currency assignment bounds', () => {
		// VBE oracle: currency_assignment_overflow_literal_runtime (...478) and
		// currency_assignment_below_min_literal_runtime compile then raise Run-time
		// error '6': Overflow; the ±922337203685477 controls run clean. The integer
		// boundary is symmetric (the .5808/.5807 fraction rounds inward to 477).
		const src =
			'Public Sub T()\n' +
			'    Dim m As Currency\n' +
			'    m = 922337203685477\n' +
			'    m = 922337203685478\n' +
			'    m = -922337203685477\n' +
			'    m = -922337203685478\n' +
			'    m = 1000000000000000\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'assignment-type-mismatch', [
			{ span: '922337203685478', message: ['Currency', "Run-time error '6'"] },
			{ span: '-922337203685478', message: ['Currency', "Run-time error '6'"] },
			{ span: '1000000000000000', message: ['Currency', "Run-time error '6'"] },
		]);
	});

	it('does not over-flag Currency bounds (in-range, fractional, and @-suffixed literals stay silent)', () => {
		// No-FP guards: an in-range whole literal is fine; fractional and
		// @-suffixed Currency literals are floatLiteral tokens with no numericValue
		// and intentionally bypass the integer-only range check.
		const src =
			'Public Sub T()\n' +
			'    Dim m As Currency\n' +
			'    m = 100\n' +
			'    m = 922337203685477\n' +
			'    m = 12.34\n' +
			'    m = 12.34@\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('accepts VBA scalar coercions and unknown assignment values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim total As Double\n' +
			'    Dim label As String\n' +
			'    Dim flexible As Variant\n' +
			'    total = "100"\n' +
			'    label = 123\n' +
			'    total = flexible\n' +
			'    total = UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('uses visible exported scalar globals for assignment target types', () => {
		const caller =
			'Public Sub T()\n' +
			'    SharedCount = "bad"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedCount As Long\n' },
			], 'Caller'),
			'assignment-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-type-mismatch', { span: '"bad"' });
	});

	it('does not leak a broader typed declaration through an untyped local assignment target', () => {
		const src =
			'Private Value As Long\n' +
			'Public Sub T()\n' +
			'    Dim Value\n' +
			'    Value = "not numeric"\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('flags array variables assigned to scalar variables', () => {
		const src =
			'Private ModuleValues(1 To 3) As Long\n' +
			'Public Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    Dim DynamicValues() As String\n' +
			'    Dim Value As Long\n' +
			'    Dim ScalarText As String\n' +
			'    Value = Values\n' +
			'    ScalarText = DynamicValues\n' +
			'    Value = ModuleValues\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'array-assignment-to-scalar');

		expectDiagnostics(src, hits, 'array-assignment-to-scalar', [
			{ severity: 'error', span: 'Values' },
			{ span: 'DynamicValues' },
			{ span: 'ModuleValues' },
		]);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('uses visible exported array globals for array assignment to scalar', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
			], 'Caller'),
			'array-assignment-to-scalar',
		);

		expect(hits).toHaveLength(1);
		expect(spanText(caller, hits[0])).toBe('SharedValues');
	});

	it('does not leak exported array shapes through local untyped assignment source shadows', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim SharedValues\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported array assignment sources quiet', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Value = SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValues() As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('does not treat visible exported array assignment targets as scalar targets', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    SharedValues = Values\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('accepts array assignments to Variant, array targets, indexed elements, and unknown values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    Dim Target() As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Dim Value As Long\n' +
			'    Flexible = Values\n' +
			'    Target = Values\n' +
			'    Value = Values(1)\n' +
			'    Value = UnknownValues\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('accepts array assignments to array-returning Function and Property Get return variables', () => {
		const src =
			'Public Function Names() As String()\n' +
			'    Dim values() As String\n' +
			'    Names = values\n' +
			'End Function\n' +
			'\n' +
			'Public Property Get MeasuresKeys() As String()\n' +
			'    Dim sOut() As String\n' +
			'    MeasuresKeys = sOut\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'array-assignment-to-scalar')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking array assignment to scalar', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'#If VBA7 Then\n' +
			'    Dim Values() As Long\n' +
			'#Else\n' +
			'    Dim Values As Long\n' +
			'#End If\n' +
			'    Value = Values\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'array-assignment-to-scalar',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'array-assignment-to-scalar',
			),
		).toHaveLength(0);
	});

	it('flags scalar variables passed to array bound functions', () => {
		const src =
			'Private ModuleValue As String\n' +
			'Public Sub T()\n' +
			'    Dim Value As Long\n' +
			'    Debug.Print LBound(Value)\n' +
			'    Debug.Print UBound(ModuleValue, 1)\n' +
			'    Debug.Print VBA.LBound(Value)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'array-bound-requires-array');

		expectDiagnostics(src, hits, 'array-bound-requires-array', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
			{ span: 'Value' },
		]);
	});

	it('uses visible exported scalar globals for array bound argument shapes', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'array-bound-requires-array',
		);

		expectDiagnostic(caller, hits, 'array-bound-requires-array', { span: 'SharedValue' });
	});

	it('does not leak exported scalar shapes through local untyped array-bound shadows', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim SharedValue\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-bound-requires-array')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported scalar array-bound arguments quiet', () => {
		const caller =
			'Public Sub T()\n' +
			'    Debug.Print LBound(SharedValue)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'array-bound-requires-array')).toHaveLength(0);
	});

	it('accepts array, Variant, unknown, and member-expression array bound arguments', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Dim FixedValues(1 To 3) As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Debug.Print LBound(Values)\n' +
			'    Debug.Print UBound(FixedValues, 1)\n' +
			'    Debug.Print LBound(Flexible)\n' +
			'    Debug.Print UBound(UnknownValues)\n' +
			'    Debug.Print LBound(Settings.Values)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'array-bound-requires-array')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking array bound functions', () => {
		const src =
			'Public Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    Debug.Print LBound(Values)\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'array-bound-requires-array',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'array-bound-requires-array',
			),
		).toHaveLength(0);
	});

	it('requires Set when assigning to a known object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws = ActiveSheet\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'set-required', { span: 'ws' });
	});

	it('reports missing Set for Object variables without treating it as scalar coercion', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    item = "blah"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-required')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});

	it('checks scalar Function return assignment types', () => {
		const src =
			'Public Function Total() As Double\n' +
			'    Total = "blah"\n' +
			'End Function\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('requires Set for object Function return assignments', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    MakePerson = New Person\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MakePerson');
	});

	it('checks object Function return assignment compatibility', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Class1\n' +
			'End Function\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('accepts compatible object Function return assignment', () => {
		const src =
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: '' },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('checks Property Get return assignments like Function returns', () => {
		const src =
			'Public Property Get Child() As Person\n' +
			'    Child = New Person\n' +
			'End Property\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set assignment between incompatible project class types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('accepts Set assignment to a project interface implemented by another class', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Class1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: 'Public Sub Save()\nEnd Sub\n' },
				{ moduleName: 'Class1', moduleKind: 'class', source: 'Implements Person\n' },
			]),
		});
		expect(byCode(diagnostics, 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('flags scalar Set assignment to a project object variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = "bad"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: '"bad"' });
	});

	it('flags scalar Set assignment from a parameterless Function reference', () => {
		const src =
			'Public Function MakeLabel() As String\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = MakeLabel\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'MakeLabel' });
	});

	it('uses source-backed member return types in Set assignment expressions', () => {
		const person = 'Public Property Get Name() As String\nEnd Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Dim child As Person\n' +
			'    Set child = p.Name\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Name' });
	});

	it('uses host member-call return types in Set assignment compatibility', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    Dim rng As Range\n' +
			'    Set rng = ActiveSheet.Range("A1")\n' +
			'    Set wb = ActiveSheet.Range("A1")\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'assignment-object-type-mismatch', { span: 'Range' });
	});

	it('errors on a nonnumeric string literal assigned to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('accepts numeric string assignment to a typed class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
	});

	it('checks assignment types for public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "blah"\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'assignment-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-type-mismatch', { span: '"blah"' });
	});

	it('accepts compatible assignments to public class fields', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = "2"\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'assignment-type-mismatch')).toHaveLength(0);
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(0);
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('uses visible UDT fields for missing-member and assignment type diagnostics', () => {
		const types = 'Public Type TPoint\n    X As Long\nEnd Type\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As TPoint\n' +
			'    p.X = "bad"\n' +
			'    p.Missing = 1\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			moduleName: 'Caller',
			projectClassMembers: projectMemberSurfaces(
				[
					{ moduleName: 'Caller', source: src },
					{ moduleName: 'Types', source: types },
				],
				'Caller',
			),
		});
		expectDiagnostic(src, diagnostics, 'assignment-type-mismatch', { span: '"bad"' });
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'Missing' });
	});

	it('requires Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expectDiagnostic(src, hits, 'set-required', { span: 'Child' });
	});

	it('accepts Set for object-valued public class fields', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Person\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'set-required')).toHaveLength(0);
		expect(byCode(diagnostics, 'set-requires-object')).toHaveLength(0);
	});

	it('requires Set for Property Set object members', () => {
		const person =
			'Private mChild As Person\n' +
			'Public Property Set Child(ByVal value As Person)\n' +
			'    Set mChild = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Child = New Person\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-required',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Child');
	});

	it('flags Set used against scalar source-backed members', () => {
		const person = 'Public Age As Integer\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'set-requires-object',
		);
		expectDiagnostic(src, hits, 'set-requires-object', { span: 'Age' });
	});

	it('flags Set assignment between incompatible source-backed object member types', () => {
		const person = 'Public Child As Person\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    Set p.Child = New Class1\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
					{ moduleName: 'Class1', moduleKind: 'class', source: '' },
				]),
			}),
			'assignment-object-type-mismatch',
		);
		expectDiagnostic(src, hits, 'assignment-object-type-mismatch', { span: 'Class1' });
	});

	it('errors on assignment to a read-only class property', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expectDiagnostic(src, hits, 'readonly-member-assignment', { span: 'Age' });
	});

	it('errors when a known class receiver uses an unknown member in assignment', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Height = 2\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Height' });
	});

	it('errors when a known class receiver calls an unknown method', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Delete');
	});

	it('errors when current class Me uses an unknown member', () => {
		const src =
			'Public Sub Save()\n' +
			'    Me.Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				moduleName: 'Person',
				moduleKind: 'class',
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: src },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
	});

	it('accepts known class members and ambiguous project receiver types', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n' +
			'Public Property Let Age(ByVal value As Integer)\n' +
			'    mAge = value\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim p As Person\n' +
			'    Set p = New Person\n' +
			'    p.Age = 2\n' +
			'    p.Unknown = 2\n' +
			'End Sub\n';
		const knownDiagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(knownDiagnostics, 'member-not-found')).toHaveLength(1);
		expect(spanText(src, byCode(knownDiagnostics, 'member-not-found')[0])).toBe(
			'Unknown',
		);
		const ambiguous = projectClassMembers([
			{ moduleName: 'Person', moduleKind: 'class', source: person },
			{ moduleName: 'Other', moduleKind: 'class', source: person.replace(/Age/g, 'Size') },
		]).map((type) => ({ ...type, name: 'Person' }));
		expect(
			byCode(analyzeModule(src, { projectClassMembers: ambiguous }), 'member-not-found'),
		).toHaveLength(0);
	});

	it('does not use source-only document module members to prove absence', () => {
		const project =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim wb As ThisWorkbook\n' +
			'    wb.DoesntExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: project },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('errors when a known standard-module qualifier uses an unknown exported member', () => {
		const caller =
			'Public Sub T()\n' +
			'    Globals.SharedValue = 1\n' +
			'    Globals.MissingValue = 2\n' +
			'    Globals.MissingProcedure\n' +
			'End Sub\n';
		const globals =
			'Public SharedValue As Long\n' +
			'Public Sub Save()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', moduleKind: 'standard', source: globals },
		], 'Caller');
		expectDiagnostics(caller, diagnostics, 'member-not-found', [
			{ span: 'MissingValue', message: 'Globals.MissingValue' },
			{ span: 'MissingProcedure', message: 'Globals.MissingProcedure' },
		]);
	});

	it('does not expose private standard-module members through qualified member diagnostics', () => {
		const caller =
			'Public Sub T()\n' +
			'    Helpers.Hidden\n' +
			'End Sub\n';
		const helpers =
			'Private Sub Hidden()\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Helpers', moduleKind: 'standard', source: helpers },
		], 'Caller');
		expectDiagnostic(caller, diagnostics, 'member-not-found', { span: 'Hidden' });
	});

	it('does not treat unknown external-style qualifiers as project member surfaces', () => {
		const caller =
			'Public Sub T()\n' +
			'    Scripting.Dictionary.CompareMode = TextCompare\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [], 'Caller');
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('errors when ThisWorkbook uses a member absent from source and the exhaustive Workbook host surface', () => {
		const project =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.doesnotexist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: project },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('does not treat Workbook events as callable object members', () => {
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.AfterSave True\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'AfterSave' });
	});

	it('accepts ThisWorkbook members from source and the exhaustive Workbook host surface', () => {
		const project =
			'Public Sub Hello()\n' +
			'End Sub\n';
		const src =
			'Public Sub T()\n' +
			'    ThisWorkbook.Hello\n' +
			'    ThisWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: project },
			]),
		});
		expect(byCode(diagnostics, 'member-not-found')).toHaveLength(0);
	});

	it('uses the exhaustive Workbook host surface for ActiveWorkbook', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveWorkbook.doesnotexist\n' +
			'    ActiveWorkbook.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('uses the exhaustive Workbook host surface for declared Workbook variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.doesnotexist\n' +
			'    wb.AcceptAllChanges\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'ThisWorkbook', moduleKind: 'document', source: '' },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'doesnotexist' });
	});

	it('uses the exhaustive Worksheet host surface for ActiveSheet', () => {
		const src =
			'Public Sub T()\n' +
			'    ActiveSheet.asdf\n' +
			'    ActiveSheet.Range("A1")\n' +
			'    ActiveSheet.Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('lets local object declarations shadow host globals for member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet As Object\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('lets untyped local declarations shadow host globals for member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set refinement for untyped host-global shadows in hard member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet\n' +
			'    Set ActiveSheet = ThisWorkbook\n' +
			'    ActiveSheet.asdf\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('uses local project declarations before host globals for member diagnostics', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Public Sub T()\n' +
			'    Dim ActiveSheet As Person\n' +
			'    ActiveSheet.Range\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);

		expectDiagnostic(src, hits, 'member-not-found', { span: 'Range' });
	});

	it('uses the exhaustive Worksheet host surface through workbook worksheet chains', () => {
		const src =
			'Public Sub T()\n' +
			'    Workbooks(1).Worksheets(1).asdf\n' +
			'    Workbooks(1).Worksheets(1).Range("A1")\n' +
			'    Workbooks(1).Worksheets(1).Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('uses the exhaustive Range host surface for declared, global, and chained receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim rng As Range\n' +
			'    ActiveCell.NoSuchMember\n' +
			'    ActiveCell.Value2\n' +
			'    rng.DoesNotExist\n' +
			'    rng.ClearContents\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").MissingRangeMember\n' +
			'    Workbooks(1).Worksheets(1).Range("A1").Offset(1, 0).Value = 1\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'NoSuchMember', message: 'Excel.Range.' },
			{ span: 'DoesNotExist', message: 'Excel.Range.' },
			{ span: 'MissingRangeMember', message: 'Excel.Range.' },
		]);
	});

	it('uses the exhaustive Application host surface only after generated promotion', () => {
		const src =
			'Public Sub T()\n' +
			'    Application.DoesNotExist\n' +
			'    Application.CentimetersToPoints 1\n' +
			'    Application.SheetCalculate\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'DoesNotExist', message: 'Excel.Application.DoesNotExist' },
			{ span: 'SheetCalculate', message: 'Excel.Application.SheetCalculate' },
		]);
	});

	it('uses exhaustive generated collection surfaces, including mixed sheet items once all candidates are promoted', () => {
		const src =
			'Public Sub T()\n' +
			'    Workbooks.MissingCollectionMember\n' +
			'    Workbooks.Open "Book.xlsx"\n' +
			'    Worksheets.MissingCollectionMember\n' +
			'    Worksheets.Add\n' +
			'    Sheets.MissingCollectionMember\n' +
			'    Sheets.Add\n' +
			'    Sheets(1).UnknownSheetOrChartMember\n' +
			'End Sub\n';
		expectDiagnostics(src, analyzeModule(src), 'member-not-found', [
			{ span: 'MissingCollectionMember', message: 'Excel.Workbooks.MissingCollectionMember' },
			{ span: 'MissingCollectionMember', message: 'Excel.Worksheets.MissingCollectionMember' },
			{ span: 'MissingCollectionMember', message: 'Excel.Sheets.MissingCollectionMember' },
			{ span: 'UnknownSheetOrChartMember' },
		]);
	});

	it('uses promoted table, chart, formatting, name, and window host surfaces', () => {
		const src =
			'Public Sub T(ws As Worksheet, rng As Range)\n' +
			'    ws.ListObjects.MissingListObjectsMember\n' +
			'    ws.ListObjects(1).MissingListObjectMember\n' +
			'    ws.ListObjects(1).ListColumns.MissingListColumnsMember\n' +
			'    ws.ListObjects(1).ListRows.MissingListRowsMember\n' +
			'    ws.ChartObjects.MissingChartObjectsMember\n' +
			'    ws.ChartObjects(1).MissingChartObjectMember\n' +
			'    ws.Shapes.MissingShapesMember\n' +
			'    ws.Shapes(1).MissingShapeMember\n' +
			'    rng.Font.MissingFontMember\n' +
			'    rng.Interior.MissingInteriorMember\n' +
			'    rng.Borders.MissingBordersMember\n' +
			'    rng.Borders(1).MissingBorderMember\n' +
			'    rng.FormatConditions.MissingFormatConditionsMember\n' +
			'    rng.FormatConditions.Add(xlCellValue, xlEqual, 1).MissingFormatConditionMember\n' +
			'    rng.Validation.MissingValidationMember\n' +
			'    rng.Hyperlinks.MissingHyperlinksMember\n' +
			'    rng.Hyperlinks(1).MissingHyperlinkMember\n' +
			'    rng.Areas.MissingAreasMember\n' +
			'    rng.Style.MissingStyleMember\n' +
			'    ws.PageSetup.MissingPageSetupMember\n' +
			'    ThisWorkbook.Names.MissingNamesMember\n' +
			'    ThisWorkbook.Names(1).MissingNameMember\n' +
			'    Application.Windows.MissingWindowsMember\n' +
			'    Application.Windows(1).MissingWindowMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MissingListObjectsMember',
			'MissingListObjectMember',
			'MissingListColumnsMember',
			'MissingListRowsMember',
			'MissingChartObjectsMember',
			'MissingChartObjectMember',
			'MissingShapesMember',
			'MissingShapeMember',
			'MissingFontMember',
			'MissingInteriorMember',
			'MissingBordersMember',
			'MissingBorderMember',
			'MissingFormatConditionsMember',
			'MissingFormatConditionMember',
			'MissingValidationMember',
			'MissingHyperlinksMember',
			'MissingHyperlinkMember',
			'MissingAreasMember',
			'MissingStyleMember',
			'MissingPageSetupMember',
			'MissingNamesMember',
			'MissingNameMember',
			'MissingWindowsMember',
			'MissingWindowMember',
		]);
	});

	it('keeps oracle-accepted WorksheetFunction and Pivot unknown members out of hard diagnostics', () => {
		const src =
			'Public Sub T(ws As Worksheet)\n' +
			'    Application.WorksheetFunction.MissingWorksheetFunctionMember\n' +
			'    Application.WorksheetFunction.Sum 1, 2\n' +
			'    ws.PivotTables.MissingPivotTablesMember\n' +
			'    ws.PivotTables(1).MissingPivotTableMember\n' +
			'    ws.PivotTables(1).PivotFields.MissingPivotFieldsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").MissingPivotFieldMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").PivotItems.MissingPivotItemsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").PivotItems(1).MissingPivotItemMember\n' +
			'    ws.PivotTables(1).PivotCache.MissingPivotCacheMember\n' +
			'    ws.PivotTables(1).PivotFilters.MissingPivotFiltersMember\n' +
			'    ws.PivotTables(1).PivotFilters(1).MissingPivotFilterMember\n' +
			'    ws.PivotTables(1).CalculatedFields.MissingCalculatedFieldsMember\n' +
			'    ws.PivotTables(1).CalculatedFields.Add("Calc", "=1").MissingCalculatedPivotFieldMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").CalculatedItems.MissingCalculatedItemsMember\n' +
			'    ws.PivotTables(1).PivotFields("Year").CalculatedItems.Add("Q1", "=1").MissingCalculatedPivotItemMember\n' +
			'    ws.PivotTables(1).CubeFields.MissingCubeFieldsMember\n' +
			'    ws.PivotTables(1).CubeFields(1).MissingCubeFieldMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(0);
	});

	it('keeps newly promoted metadata-only host families out of hard diagnostics', () => {
		const src =
			'Public Sub T(ws As Worksheet, ch As Chart)\n' +
			'    ws.QueryTables(1).MissingQueryTableMember\n' +
			'    ws.QueryTables(1).Parameters(1).MissingParameterMember\n' +
			'    ThisWorkbook.Connections(1).MissingWorkbookConnectionMember\n' +
			'    ThisWorkbook.Connections(1).OLEDBConnection.MissingOleDbConnectionMember\n' +
			'    ThisWorkbook.Connections(1).ModelTables(1).MissingModelTableMember\n' +
			'    ch.SeriesCollection(1).MissingSeriesMember\n' +
			'    ch.SeriesCollection(1).Points(1).MissingPointMember\n' +
			'    ch.SeriesCollection(1).Trendlines(1).MissingTrendlineMember\n' +
			'    ch.Axes(1).MissingAxisMember\n' +
			'    ch.Axes(1).AxisTitle.MissingAxisTitleMember\n' +
			'    ws.Shapes.Range.MissingShapeRangeMember\n' +
			'    ws.Shapes.Range.GroupItems.MissingGroupShapesMember\n' +
			'    ws.Comments(1).MissingCommentMember\n' +
			'    ws.CommentsThreaded(1).MissingThreadedCommentMember\n' +
			'    ws.AutoFilter.Filters(1).MissingFilterMember\n' +
			'    ws.Sort.SortFields.Add(ws.Range("A1")).MissingSortFieldMember\n' +
			'    ws.OLEObjects(1).MissingOleObjectMember\n' +
			'    ws.Buttons(1).MissingButtonMember\n' +
			'    ws.CheckBoxes(1).MissingCheckBoxMember\n' +
			'    ws.DropDowns(1).MissingDropDownMember\n' +
			'    ws.OptionButtons(1).MissingOptionButtonMember\n' +
			'    ThisWorkbook.SlicerCaches.MissingSlicerCachesMember\n' +
			'    ThisWorkbook.SlicerCaches(1).MissingSlicerCacheMember\n' +
			'    ThisWorkbook.SlicerCaches(1).Slicers.MissingSlicersMember\n' +
			'    ThisWorkbook.SlicerCaches(1).SlicerItems(1).MissingSlicerItemMember\n' +
			'    ThisWorkbook.SlicerCaches(1).TimelineState.MissingTimelineStateMember\n' +
			'    ws.Shapes.Item(1).Fill.MissingFillFormatMember\n' +
			'    ws.Shapes.Item(1).Line.MissingLineFormatMember\n' +
			'    ws.Shapes.Item(1).TextFrame.MissingTextFrameMember\n' +
			'    ws.Shapes.Item(1).TextFrame2.MissingTextFrame2Member\n' +
			'    ws.Shapes.Item(1).PictureFormat.MissingPictureFormatMember\n' +
			'    ws.Shapes.Item(1).Nodes.MissingShapeNodesMember\n' +
			'    ch.ChartTitle.MissingChartTitleMember\n' +
			'    ch.Legend.LegendEntries.MissingLegendEntriesMember\n' +
			'    ch.ChartGroups.MissingChartGroupsMember\n' +
			'    ch.DataTable.MissingDataTableMember\n' +
			'    ch.Walls.MissingWallsMember\n' +
			'    ws.Range("A1").FormatConditions.AddDatabar.MissingDatabarMember\n' +
			'    ws.Range("A1").FormatConditions.AddColorScale(3).ColorScaleCriteria.MissingColorScaleCriteriaMember\n' +
			'    ws.Range("A1").DisplayFormat.MissingDisplayFormatMember\n' +
			'    ws.Drawings.MissingDrawingsMember\n' +
			'    ws.Pictures.MissingPicturesMember\n' +
			'    ws.Lines.MissingLinesMember\n' +
			'    ws.Range("A1").SparklineGroups.MissingSparklineGroupsMember\n' +
			'    ThisWorkbook.XmlMaps.MissingXmlMapsMember\n' +
			'    ThisWorkbook.XmlMaps(1).Schemas.MissingXmlSchemasMember\n' +
			'    ThisWorkbook.PublishObjects.MissingPublishObjectsMember\n' +
			'    ThisWorkbook.WebOptions.MissingWebOptionsMember\n' +
			'    Application.DefaultWebOptions.MissingDefaultWebOptionsMember\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'member-not-found');
		expect(hits).toHaveLength(0);
	});

	it('uses the exhaustive runtime object surface for Err', () => {
		const src =
			'Public Sub T()\n' +
			'    Err.Raise vbObjectError + 1, "M", "boom"\n' +
			'    Err.DoesNotExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'DoesNotExist' });
	});

	it('uses the exhaustive runtime object surface for Debug', () => {
		const src =
			'Public Sub T()\n' +
			'    Debug.Print "value", 1\n' +
			'    Debug.Assert True\n' +
			'    Debug.DoesNotExist\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, { knownIdentifiers: new Set<string>() });
		expect(byCode(diagnostics, 'undeclared-variable')).toHaveLength(0);
		expect(byCode(diagnostics, 'argument-count')).toHaveLength(0);
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'DoesNotExist' });
	});

	it('uses the current workbook Me host surface for ThisWorkbook modules', () => {
		const src =
			'Public Sub T()\n' +
			'    Me.asdf\n' +
			'    Me.AcceptAllChanges\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, { moduleName: 'ThisWorkbook', moduleKind: 'document' }),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'asdf' });
	});

	it('uses the exhaustive Worksheet host surface for declared Worksheet variables', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = ActiveSheet\n' +
			'    ws.asdf\n' +
			'    ws.Range("A1")\n' +
			'    ws.Buttons\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'member-not-found', { span: 'asdf' });
	});

	it('does not prove missing members from late-bound Object or Variant receivers', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.asdf\n' +
			'    flexible.asdf\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('does not use Set-assignment refinement for hard late-bound member-call diagnostics', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Set obj = ActiveSheet\n' +
			'    Set flexible = Workbooks(1).Worksheets(1)\n' +
			'    obj.Range()\n' +
			'    flexible.Range()\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'argument-count')).toHaveLength(0);
	});

	it('uses an exhaustive host object model to prove a missing member', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					exhaustive: true,
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src, { hostModel: model }), 'member-not-found', {
			span: 'Missing',
		});
	});

	it('does not use a curated non-exhaustive host object model to prove absence', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src =
			'Public Sub T()\n' +
			'    Thing.Missing\n' +
			'End Sub\n';
		expect(
			byCode(analyzeModule(src, { hostModel: model }), 'member-not-found'),
		).toHaveLength(0);
	});
});

describe('analyzeModule - missing Function return assignment', () => {
	it('warns when a Function never assigns its return variable', () => {
		const src =
			'Public Function myFunction()\n' +
			'\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expectDiagnostic(src, hits, 'missing-return-assignment', {
			severity: 'warning',
			span: 'myFunction',
		});
	});

	it('warns when a TYPED Function or Property Get falls through (issue #46)', () => {
		// This used to be silent: the rule skipped any declared return type, on the
		// reasoning that a typed default might be intentional. It is not, in the
		// case that matters - a Function As Double that never names itself returns
		// 0 to every caller, compiles, and gets found by a wrong number rather than
		// by an error. Measured over 67 modules of third-party code, widening it
		// costs 2 findings once the carve-outs below are in place, down from 40.
		const src =
			'Public Function Label() As String\n' +
			'    Dim x As Long\n' +
			'    x = 1\n' +
			'End Function\n' +
			'\n' +
			'Public Property Get Name() As String\n' +
			'    Dim y As Long\n' +
			'    y = 1\n' +
			'End Property\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(2);
	});

	it('accepts a return whose FIELDS are assigned (a UDT return)', () => {
		// `utc_DateToSystemTime.utc_wYear = ...` IS the return assignment. Reading
		// only a bare `Name =` counted 16 such functions in the corpus as silent.
		const src =
			'Private Type TPoint\n' +
			'    X As Long\n' +
			'End Type\n' +
			'Public Function MakePoint() As TPoint\n' +
			'    MakePoint.X = 1\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	// github.com/WilliamSmithEdward/xlide_vscode/issues/60. The carve-out used
	// to ask the PARSED module kind, which only says `class` when the source
	// carries `Attribute VB_Exposed` and friends. Module text read out of a
	// project carries no attribute lines, so it never fired for a real class.
	// It now asks whether some module in the project implements this one, which
	// is both the authoritative signal and a narrower one.
	describe('an interface declares its members empty on purpose', () => {
		const IFACE =
			'Option Explicit\n' +
			'Public Sub Save()\n' +
			'End Sub\n' +
			'Public Function Total(ByVal n As Long) As Currency\n' +
			'End Function\n' +
			'Public Property Get Name() As String\n' +
			'End Property\n';
		const IMPL =
			'Option Explicit\n' +
			'Implements IStore\n' +
			'Private Sub IStore_Save()\n' +
			'End Sub\n' +
			'Private Function IStore_Total(ByVal n As Long) As Currency\n' +
			'    IStore_Total = 0\n' +
			'End Function\n' +
			'Private Property Get IStore_Name() As String\n' +
			'    IStore_Name = ""\n' +
			'End Property\n';
		const PROJECT = [
			{ moduleName: 'IStore', source: IFACE, moduleKind: 'class' as const },
			{ moduleName: 'Store', source: IMPL, moduleKind: 'class' as const },
		];

		it('stays quiet on a module another module implements', () => {
			const hits = analyzeProjectModule(IFACE, PROJECT, 'IStore', { moduleKind: 'class' });
			expect(byCode(hits, 'missing-return-assignment')).toHaveLength(0);
		});

		it('stays quiet on the implementer, which supplies the bodies', () => {
			const hits = analyzeProjectModule(IMPL, PROJECT, 'Store', { moduleKind: 'class' });
			expect(byCode(hits, 'missing-return-assignment')).toHaveLength(0);
		});

		it('still reports an ordinary class nobody implements', () => {
			const alone = [{ moduleName: 'IStore', source: IFACE, moduleKind: 'class' as const }];
			const hits = analyzeProjectModule(IFACE, alone, 'IStore', { moduleKind: 'class' });
			expect(byCode(hits, 'missing-return-assignment')).toHaveLength(2);
		});

		it('still reports an interface member that has a body but forgets the return', () => {
			const forgot = IFACE.replace(
				'Public Function Total(ByVal n As Long) As Currency\nEnd Function',
				'Public Function Total(ByVal n As Long) As Currency\n    Debug.Print n\nEnd Function',
			);
			const project = [{ moduleName: 'IStore', source: forgot, moduleKind: 'class' as const }, PROJECT[1]];
			const hits = analyzeProjectModule(forgot, project, 'IStore', { moduleKind: 'class' });
			expect(byCode(hits, 'missing-return-assignment')).toHaveLength(1);
		});
	});

	it('accepts a typed Function whose work is to raise', () => {
		const src =
			'Public Function NotImplemented() As Long\n' +
			'    Err.Raise 5, "NotImplemented"\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts scalar and object Function return assignments', () => {
		const src =
			'Public Function Label() As String\n' +
			'    Label = "ready"\n' +
			'End Function\n' +
			'\n' +
			'Public Function MakePerson() As Person\n' +
			'    Set MakePerson = New Person\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to known ByRef helper parameters', () => {
		const src =
			'Public Function Run(ParamArray params() As Variant)\n' +
			'    Call CopyVariant(Run, RunEx(params))\n' +
			'End Function\n' +
			'\n' +
			'Private Function RunEx(ByVal values As Variant) As Variant\n' +
			'    RunEx = values\n' +
			'End Function\n' +
			'\n' +
			'Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to project-visible ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Call CopyVariant(dest:=Run, value:=1)\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts Function return variables passed to module-qualified ByRef helper parameters', () => {
		const runSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    Helpers.CopyVariant Run, 1\n' +
			'End Function\n';
		const helpersSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'    dest = value\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			runSrc,
			[
				{ moduleName: 'Runner', source: runSrc },
				{ moduleName: 'Helpers', source: helpersSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count ByVal or ambiguous project helper arguments as return assignments', () => {
		const byValRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byValHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByVal dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';
		const ambiguousRunSrc =
			'Option Explicit\n' +
			'Public Function Run()\n' +
			'    CopyVariant Run, 1\n' +
			'End Function\n';
		const byRefHelperSrc =
			'Option Explicit\n' +
			'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
			'End Sub\n';

		expect(byCode(analyzeProjectModule(
			byValRunSrc,
			[
				{ moduleName: 'Runner', source: byValRunSrc },
				{ moduleName: 'Helpers', source: byValHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
		expect(byCode(analyzeProjectModule(
			ambiguousRunSrc,
			[
				{ moduleName: 'Runner', source: ambiguousRunSrc },
				{ moduleName: 'HelpersA', source: byRefHelperSrc },
				{ moduleName: 'HelpersB', source: byRefHelperSrc },
			],
			'Runner',
		), 'missing-return-assignment')).toHaveLength(1);
	});

	it('accepts return assignments in the active default VBA7 branch', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    HandleValue = 7\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('accepts active VBA7 return assignments with blank lines before #End If', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			' HandleValue = 0\n' +
			'#Else\n' +
			'HandleValue = 1\n' +
			'\n' +
			'\n' +
			'#End If\n' +
			'    \n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('does not count return assignments from inactive default VBA7 branches', () => {
		const src =
			'Public Function HandleValue()\n' +
			'#If VBA7 Then\n' +
			'    Debug.Print "active branch"\n' +
			'#Else\n' +
			'    HandleValue = 6\n' +
			'#End If\n' +
			'End Function\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('HandleValue');
	});

	it('does not warn on conditionally split VBA7 Function headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function HandleValue()\n' +
			'#Else\n' +
			'Public Function HandleValue()\n' +
			'#End If\n' +
			'    HandleValue = 1\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'missing-return-assignment')).toHaveLength(0);
	});

	it('keeps parsing after conditionally split procedure headers with a shared body', () => {
		const src =
			'#If VBA7 Then\n' +
			'Public Function Create(ByVal pointer As LongPtr) As Object\n' +
			'#Else\n' +
			'Public Function Create(ByVal pointer As Long) As Object\n' +
			'#End If\n' +
			'    Set Create = Nothing\n' +
			'End Function\n' +
			'Friend Sub Init(ByVal mode As Long)\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Create(0)\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('does not flag unknown-constant alternate procedure headers as declarations in a procedure', () => {
		const src =
			'#If FULL_INTELLISENSE Then\n' +
			'Public Function AsAcc() As stdAcc\n' +
			'#Else\n' +
			'Public Function AsAcc() As Object\n' +
			'#End If\n' +
			'    Set AsAcc = Nothing\n' +
			'End Function\n';

		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('ignores exported member Attribute lines while parsing procedure bodies', () => {
		const src =
			'Attribute VB_Name = "Module1"\n' +
			'Friend Sub Init(ByVal mode As Long)\n' +
			'Attribute Init.VB_Description = "Initialises this object."\n' +
			'    Select Case mode\n' +
			'        Case 1\n' +
			'            With Nothing\n' +
			'                .Name = "ready"\n' +
			'            End With\n' +
			'    End Select\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'case-outside-select')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'member-access-outside-with')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'module-declaration-in-procedure')).toHaveLength(0);
	});

	it('checks Property Get procedures and ignores Subs', () => {
		const src =
			'Public Property Get Name()\n' +
			'End Property\n' +
			'\n' +
			'Public Sub Refresh()\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'missing-return-assignment');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Name');
	});
});

describe('analyzeModule - string arithmetic coercion', () => {
	it('errors on a nonnumeric string literal inside a numeric assignment expression', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As Integer\n' +
			'    shouldErrorTest1 = 1 + "string"\n' +
			'End Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'string-arithmetic-coercion', {
			severity: 'error',
			span: '"string"',
		});
	});

	it('does not warn on numeric strings or unknown arithmetic operands', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim n As Integer\n' +
			'    n = 1 + "2"\n' +
			'    n = 1 + UnknownValue\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
	});

	it('accepts plus between string literals assigned to a String variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest1 As String\n' +
			'    shouldErrorTest1 = "string" + "string"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'string-arithmetic-coercion')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'assignment-type-mismatch')).toHaveLength(0);
	});
});

describe('analyzeModule - Set assignment validation', () => {
	it('flags Set assignment to a known scalar variable', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim shouldErrorTest5 As Integer\n' +
			'    Set shouldErrorTest5 = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('shouldErrorTest5');
	});

	it('flags Set assignment to a String variable even when assigning a New object expression', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim Text As String\n' +
			'    Set text = New Collection\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'set-requires-object');

		expectDiagnostic(src, hits, 'set-requires-object', { span: 'text' });
	});

	it('flags Set assignment to visible exported scalar globals', () => {
		const caller =
			'Public Sub T()\n' +
			'    Set SharedText = New Collection\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'set-requires-object',
		);

		expectDiagnostic(caller, hits, 'set-requires-object', { span: 'SharedText' });
	});

	it('requires Set for visible exported object globals assigned with plain assignment', () => {
		const caller =
			'Public Sub T()\n' +
			'    SharedObject = New Collection\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedObject As Object\n' },
			], 'Caller'),
			'set-required',
		);

		expectDiagnostic(caller, hits, 'set-required', { span: 'SharedObject' });
	});

	it('uses visible exported scalar globals as bare Set RHS value types', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'assignment-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses module-qualified exported scalar globals as Set RHS value types', () => {
		const caller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = Globals.SharedText\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'assignment-object-type-mismatch',
		);

		expectDiagnostic(caller, hits, 'assignment-object-type-mismatch', { span: 'SharedText' });
	});

	it('uses verified runtime and host constants as Set RHS value types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = vbFalse\n' +
			'    Set target = VBA.vbFalse\n' +
			'    Set target = xlAbove\n' +
			'    Set target = vbNullString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-object-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-object-type-mismatch', [
			{ span: 'vbFalse', message: 'vbFalse As VbTriState' },
			{ span: 'vbFalse', message: 'VBA.vbFalse As VbTriState' },
			{ span: 'xlAbove', message: 'xlAbove As Constants' },
			{ span: 'vbNullString', message: 'vbNullString As String' },
		]);
	});

	it('lets source declarations shadow verified external constants for Set RHS value types', () => {
		const src =
			'Public Function vbFalse() As Object\n' +
			'End Function\n' +
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Dim xlAbove As Object\n' +
			'    Set target = vbFalse\n' +
			'    Set target = xlAbove\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('uses host globals as Set RHS object value types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim sheet As Worksheet\n' +
			'    Dim book As Workbook\n' +
			'    Set sheet = Application\n' +
			'    Set book = ActiveCell\n' +
			'    Set sheet = ActiveSheet\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'assignment-object-type-mismatch');

		expectDiagnostics(src, hits, 'assignment-object-type-mismatch', [
			{ span: 'Application', message: 'Application As Excel.Application' },
			{ span: 'ActiveCell', message: 'ActiveCell As Excel.Range' },
		]);
	});

	it('allows compatible and generic host global Set assignments', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim app As Application\n' +
			'    Dim item As Object\n' +
			'    Set app = Application\n' +
			'    Set item = ActiveCell\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('lets source declarations shadow host globals for Set RHS object values', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim sheet As Worksheet\n' +
			'    Dim Application As Worksheet\n' +
			'    Dim ActiveCell As Worksheet\n' +
			'    Set sheet = Application\n' +
			'    Set sheet = ActiveCell\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('keeps local shadows and ambiguous exported Set RHS globals quiet', () => {
		const shadowCaller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Dim SharedText\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller'), 'assignment-object-type-mismatch')).toHaveLength(0);

		const ambiguousCaller =
			'Public Sub T()\n' +
			'    Dim target As Object\n' +
			'    Set target = SharedText\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller'), 'assignment-object-type-mismatch')).toHaveLength(0);
	});

	it('does not guess Set target types for ambiguous visible exported globals', () => {
		const caller =
			'Public Sub T()\n' +
			'    Set SharedText = New Collection\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'set-requires-object')).toHaveLength(0);
	});

	it('does not flag Set assignment to Object, Variant, or unknown object types', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim item As Object\n' +
			'    Dim flexible As Variant\n' +
			'    Dim sheet As Worksheet\n' +
			'    Set item = Nothing\n' +
			'    Set flexible = Nothing\n' +
			'    Set sheet = Nothing\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'set-requires-object')).toHaveLength(0);
	});
});

describe('analyzeModule - assignment no-diagnostic boundary controls (rule audit backfill)', () => {
	it('keeps ambiguous exported object globals and untyped local object shadows quiet', () => {
		// set-required stays quiet when the target name is an ambiguous exported
		// object global (two modules declaring it) or an untyped local shadow.
		const ambiguousCaller =
			'Public Sub T()\n' +
			'    SharedObject = New Collection\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedObject As Object\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedObject As Object\n' },
		], 'Caller'), 'set-required')).toHaveLength(0);

		const shadowSrc =
			'Private ModuleObject As Object\n' +
			'Public Sub T()\n' +
			'    Dim ModuleObject\n' +
			'    ModuleObject = "blah"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(shadowSrc), 'set-required')).toHaveLength(0);
	});

	it('stays quiet for ambiguous project receiver types and late-bound Object/Variant receivers', () => {
		// An ambiguous project receiver and late-bound Object/Variant receivers
		// cannot prove member writability, so all hit the writable===undefined
		// suppression path and readonly-member-assignment stays quiet.
		const readonlyPerson =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Public Sub T()\n' +
			'    Dim ambiguous As Person\n' +
			'    Dim lateObj As Object\n' +
			'    Dim lateVar As Variant\n' +
			'    Set ambiguous = New Person\n' +
			'    ambiguous.Age = 2\n' +
			'    lateObj.Age = 2\n' +
			'    lateVar.Age = 2\n' +
			'End Sub\n';
		const ambiguous = projectClassMembers([
			{ moduleName: 'Person', moduleKind: 'class', source: readonlyPerson },
			{ moduleName: 'Other', moduleKind: 'class', source: readonlyPerson.replace(/Age/g, 'Size') },
		]).map((type) => ({ ...type, name: 'Person' }));
		expect(
			byCode(analyzeModule(src, { projectClassMembers: ambiguous }), 'readonly-member-assignment'),
		).toHaveLength(0);
	});
});

describe('analyzeModule - Mid statement literal target', () => {
	it('flags a Mid$ statement whose target is a string literal', () => {
		const src = 'Sub T()\n\tMid$("abcdef", 2, 3) = "XYZ"\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'mid-statement-literal-target');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('"abcdef"');
		expect(hits[0].severity).toBe('error');
	});

	it('flags a bare Mid (no $) statement whose target is a string literal', () => {
		const src = 'Sub T()\n\tMid("abcdef", 2) = "zz"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(1);
	});

	it('flags a MidB$ statement whose target is a string literal', () => {
		const src = 'Sub T()\n\tMidB$("abcdef", 2) = "z"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(1);
	});

	it('accepts a Mid$ statement with a writable String variable target', () => {
		const src = 'Sub T()\n\tDim s As String\n\ts = "abcdef"\n\tMid$(s, 2, 3) = "XYZ"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('does not flag Mid used as a function on the right-hand side', () => {
		const src = 'Sub T()\n\tDim x As String\n\tx = Mid$("abcdef", 2, 3)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('does not flag a Mid comparison expression (not the Mid statement)', () => {
		const src = 'Sub T()\n\tDebug.Print Mid$("abcdef", 1, 1) = "a"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('does not flag when the target is not a single string literal', () => {
		const src = 'Sub T(o As Object)\n\tMid$(o.Name, 2) = "z"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('stays silent for the whole module when a user symbol shadows Mid', () => {
		const src = 'Sub T()\n\tDim Mid(3) As Variant\n\tMid("0") = "x"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('stays silent when Mid is implicitly declared by a ReDim (no prior Dim)', () => {
		const src = 'Sub T()\n\tReDim Mid(3) As String\n\tMid("0") = "x"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('stays silent when Mid is implicitly declared by a ReDim Preserve', () => {
		const src = 'Sub T()\n\tReDim Preserve Mid(3)\n\tMid("0") = "x"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});
});

describe('analyzeModule - Mid statement literal target (adversarial)', () => {
	it('does not flag a Mid literal target inside an inactive #If branch', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'\tMid$("abcdef", 2, 3) = "XYZ"\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('stays conservative when the target is a concatenated-literal expression', () => {
		const src = 'Sub T()\n\tMid$("ab" & "cd", 1) = "z"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(0);
	});

	it('flags a Mid literal target after a numeric line label', () => {
		const src = 'Sub T()\n100 Mid$("abcdef", 2) = "z"\nEnd Sub\n';
		expect(byCode(analyzeModule(src), 'mid-statement-literal-target')).toHaveLength(1);
	});
});
