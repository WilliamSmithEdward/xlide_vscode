// Diagnostics tests: arrays rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

describe('analyzeModule - array ReDim', () => {
	it('flags ReDim of a local fixed-size array', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'    ReDim Values(1 To 10) As Long\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-array-redim');

		expectDiagnostic(src, hits, 'fixed-array-redim', { severity: 'error', span: 'Values' });
	});

	it('flags ReDim Preserve and module-level fixed-size arrays', () => {
		const src =
			'Private Values(1 To 3) As Long\n' +
			'Sub T()\n' +
			'    ReDim Preserve Values(1 To 10)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'fixed-array-redim');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Values');
	});

	it('flags ReDim of scalar variables', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim Value As Long\n' +
			'    ReDim Value(1 To 10)\n' +
			'    ReDim Preserve ModuleValue(1 To 10)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-redim');

		expectDiagnostics(src, hits, 'scalar-redim', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('does not flag a qualified ReDim target that resizes a member array', () => {
		// `ReDim b.b(...)` resizes the dynamic-array member `b.b`, not the scalar
		// container `b`; VBE accepts this (oracle-verified) and the rule must not
		// mistake the qualifier for the array being resized.
		const src =
			'Private Type TBuf\n' +
			'    n As Long\n' +
			'    b() As Byte\n' +
			'End Type\n' +
			'Sub T()\n' +
			'    Dim b As TBuf\n' +
			'    ReDim b.b(0 To 3)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(0);
	});

	it('uses visible exported scalar and fixed-array globals for ReDim target shapes', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedScalar(1 To 10)\n' +
			'    ReDim SharedFixed(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedScalar As Long\n' +
					'Public SharedFixed(1 To 3) As Long\n',
			},
		], 'Caller');
		const scalarHits = byCode(diagnostics, 'scalar-redim');
		const fixedHits = byCode(diagnostics, 'fixed-array-redim');

		expect(scalarHits).toHaveLength(1);
		expect(spanText(caller, scalarHits[0])).toBe('SharedScalar');
		expect(fixedHits).toHaveLength(1);
		expect(spanText(caller, fixedHits[0])).toBe('SharedFixed');
	});

	it('does not flag visible exported dynamic arrays as invalid ReDim targets', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValues(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues() As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts ReDim Preserve on Variant and implicit Variant targets', () => {
		const src =
			'Public Function RunEx(ByVal vArr As Variant) As Variant\n' +
			'    ReDim Preserve vArr(0 To 29)\n' +
			'    Dim flexible As Variant\n' +
			'    ReDim flexible(0 To 2)\n' +
			'    Dim implicit\n' +
			'    ReDim implicit(0 To 2)\n' +
			'End Function\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts ReDim on visible exported Variant globals', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValues(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedValues As Variant\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('lets local dynamic arrays shadow exported invalid ReDim targets', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedScalar() As Long\n' +
			'    Dim SharedFixed() As Long\n' +
			'    ReDim SharedScalar(1 To 10)\n' +
			'    ReDim SharedFixed(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedScalar As Long\n' +
					'Public SharedFixed(1 To 3) As Long\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported ReDim targets quiet', () => {
		const caller =
			'Sub T()\n' +
			'    ReDim SharedValue(1 To 10)\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue(1 To 3) As Long\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-redim')).toHaveLength(0);
		expect(byCode(diagnostics, 'fixed-array-redim')).toHaveLength(0);
	});

	it('accepts dynamic arrays and undeclared ReDim targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim dynamicValues() As Long\n' +
			'    ReDim dynamicValues(1 To 10)\n' +
			'    ReDim implicitValues(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('lets a local dynamic array shadow a module fixed-size array', () => {
		const src =
			'Private Values(1 To 3) As Long\n' +
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('lets a local dynamic array shadow a module scalar for ReDim', () => {
		const src =
			'Private Values As Long\n' +
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(0);
	});

	it('ignores fixed-size arrays in inactive conditional branches', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values(1 To 3) As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'fixed-array-redim',
			),
		).toHaveLength(0);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'scalar-redim',
			),
		).toHaveLength(0);
	});

	it('uses the active conditional branch when checking ReDim of scalar variables', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    ReDim Values(1 To 10)\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'scalar-redim',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'scalar-redim',
			),
		).toHaveLength(0);
	});

	it('flags ReDim dimensions whose explicit lower bound is greater than the upper bound', () => {
		const src =
			'Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim grid() As Long\n' +
			'    ReDim values(10 To 1)\n' +
			'    ReDim Preserve values(-1 To -2)\n' +
			'    ReDim grid(1 To 3, 5 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-impossible-bounds');

		expectDiagnostics(src, hits, 'redim-impossible-bounds', [
			{ severity: 'error', span: '10 To 1', message: ["Run-time error '9'", 'dimension 1'] },
			{ span: '-1 To -2' },
			{ span: '5 To 2', message: 'dimension 2' },
		]);
	});

	it('accepts ReDim bounds that are equal, increasing, unknown, upper-only, or inactive', () => {
		const src =
			'Sub T(ByVal first As Long, ByVal last As Long)\n' +
			'    Dim values() As Long\n' +
			'    ReDim values(1 To 1)\n' +
			'    ReDim values(1 To 10)\n' +
			'    ReDim values(first To last)\n' +
			'    ReDim values(10)\n' +
			'#If Win64 Then\n' +
			'    ReDim values(10 To 1)\n' +
			'#End If\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { Win64: false } },
				}),
				'redim-impossible-bounds',
			),
		).toHaveLength(0);
	});

	it('does not add runtime ReDim bounds diagnostics for known invalid ReDim targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim fixed(1 To 3) As Long\n' +
			'    Dim scalar As Long\n' +
			'    ReDim fixed(10 To 1)\n' +
			'    ReDim scalar(10 To 1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'redim-impossible-bounds')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'fixed-array-redim')).toHaveLength(1);
		expect(byCode(analyzeModule(src), 'scalar-redim')).toHaveLength(1);
	});

	it('flags ReDim Preserve changing a non-final dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 3, 1 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change', {
			severity: 'error',
			span: 'grid',
		});
	});

	it('flags ReDim Preserve changing the known dimension count', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change');
	});

	it('accepts ReDim Preserve changing only the last dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2, 1 To 3)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'redim-preserve-dimension-change')).toHaveLength(0);
	});

	it('flags ReDim Preserve changing the lower bound of the final dimension', () => {
		const src =
			'Sub T()\n' +
			'    Dim grid() As Long\n' +
			'    ReDim grid(1 To 2, 1 To 2)\n' +
			'    ReDim Preserve grid(1 To 2, 0 To 3)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'redim-preserve-dimension-change');

		expectDiagnostic(src, hits, 'redim-preserve-dimension-change');
	});

	it('does not let ReDim shapes learned inside nested blocks leak outward', () => {
		const src =
			'Sub T(ByVal flag As Boolean)\n' +
			'    Dim grid() As Long\n' +
			'    If flag Then\n' +
			'        ReDim grid(1 To 2, 1 To 2)\n' +
			'    End If\n' +
			'    ReDim Preserve grid(1 To 3, 1 To 2)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'redim-preserve-dimension-change')).toHaveLength(0);
	});
});

describe('analyzeModule - array declaration impossible bounds', () => {
	const CODE = 'array-declaration-impossible-bounds';

	it('flags a Dim array declaration whose lower bound exceeds the upper bound', () => {
		const src = 'Public Sub T()\n    Dim a(10 To 1) As Long\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('10 To 1');
	});

	it('flags module-level Private/Static arrays and negative reversed bounds', () => {
		expect(byCode(analyzeModule('Private p(5 To 2) As Long\n'), CODE)).toHaveLength(1);
		expect(byCode(analyzeModule('Public g(0 To -1) As Long\n'), CODE)).toHaveLength(1);
		expect(
			byCode(analyzeModule('Sub T()\n    Static s(-5 To -10) As Long\nEnd Sub\n'), CODE),
		).toHaveLength(1);
	});

	it('flags only the offending dimension in a multi-dimensional declaration', () => {
		const src = 'Dim grid(1 To 3, 10 To 1, 0 To 5) As Long\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('10 To 1');
		expect(hits[0].message).toContain('dimension 2');
	});

	it('stays quiet for valid, equal, upper-only, variable, and dynamic bounds', () => {
		expect(byCode(analyzeModule('Dim a(1 To 10) As Long\n'), CODE)).toHaveLength(0); // increasing
		expect(byCode(analyzeModule('Dim a(5 To 5) As Long\n'), CODE)).toHaveLength(0); // equal
		expect(byCode(analyzeModule('Dim a(10) As Long\n'), CODE)).toHaveLength(0); // upper-only
		expect(byCode(analyzeModule('Dim a() As Long\n'), CODE)).toHaveLength(0); // dynamic
		expect(
			byCode(analyzeModule('Sub T(ByVal n As Long, ByVal m As Long)\n    Dim a(n To m) As Long\nEnd Sub\n'), CODE),
		).toHaveLength(0); // variable
	});

	it('stays quiet for non-literal constant-reference bounds (cannot prove reversed)', () => {
		// HI/LO could be any constants; without folding we cannot prove lower > upper.
		const src = 'Const HI As Long = 10\nConst LO As Long = 1\nDim a(HI To LO) As Long\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('flags only the offending name in a multi-declaration statement', () => {
		const src = 'Dim a(1 To 3) As Long, b(10 To 1) As Long\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('10 To 1');
	});

	it('does not double-report on ReDim (owned by redim-impossible-bounds)', () => {
		const src = 'Sub T()\n    Dim a() As Long\n    ReDim a(10 To 1)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'redim-impossible-bounds')).toHaveLength(1);
	});

	it('stays quiet for a reversed bound in an inactive #If branch', () => {
		const src = '#If 0 Then\nDim a(10 To 1) As Long\n#End If\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - unallocated dynamic array access', () => {
	it('flags straight-line local dynamic array indexed before ReDim', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Debug.Print values(0)\n' +
			'    values(1) = 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expectDiagnostics(src, hits, 'unallocated-dynamic-array-access', [
			{ severity: 'error', span: 'values' },
			{ span: 'values' },
		]);
	});

	it('accepts allocated and fixed arrays but flags again after Erase', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim fixedValues(0 To 1) As Long\n' +
			'    ReDim values(0 To 2)\n' +
			'    Debug.Print values(0)\n' +
			'    Debug.Print fixedValues(0)\n' +
			'    Erase values\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('values');
	});

	it('flags LBound and UBound on unallocated local dynamic arrays', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Debug.Print LBound(values)\n' +
			'    Debug.Print UBound(values, 1)\n' +
			'    Debug.Print VBA.LBound(values)\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expectDiagnostics(src, hits, 'unallocated-dynamic-array-access', [
			{ span: 'values', message: 'LBound' },
			{ span: 'values', message: 'UBound' },
			{ span: 'values' },
		]);
	});

	it('accepts bound intrinsics after ReDim and non-intrinsic member calls', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    ReDim values(0 To 2)\n' +
			'    Debug.Print LBound(values)\n' +
			'    Debug.Print Helpers.LBound(values)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('keeps module-level arrays parameters Static locals and helper-touched arrays quiet', () => {
		const src =
			'Private moduleValues() As Long\n' +
			'Private Sub Fill(ByRef target() As Long)\n' +
			'End Sub\n' +
			'Public Sub T(ByRef paramValues() As Long)\n' +
			'    Static cached() As Long\n' +
			'    Dim helperValues() As Long\n' +
			'    Fill helperValues\n' +
			'    Debug.Print moduleValues(0)\n' +
			'    Debug.Print paramValues(0)\n' +
			'    Debug.Print cached(0)\n' +
			'    Debug.Print helperValues(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('treats an array passed to a function in an expression as possibly allocated (#70)', () => {
		const src =
			'Option Explicit\n' +
			'Private Type Thing\n' +
			'    Value As Long\n' +
			'End Type\n' +
			'Public Function FillUdt(ByRef items() As Thing) As Long\n' +
			'    ReDim items(0 To 2)\n' +
			'    FillUdt = 3\n' +
			'End Function\n' +
			'Public Function FillLong(ByRef items() As Long) As Long\n' +
			'    ReDim items(0 To 2)\n' +
			'    FillLong = 3\n' +
			'End Function\n' +
			'Public Sub CaseA_UdtByRef()\n' +
			'    Dim items() As Thing\n' +
			'    Dim n As Long\n' +
			'    n = FillUdt(items)\n' +
			'    Debug.Print "A " & n & " " & items(0).Value\n' +
			'End Sub\n' +
			'Public Sub CaseB_LongByRef()\n' +
			'    Dim items() As Long\n' +
			'    Dim n As Long\n' +
			'    n = FillLong(items)\n' +
			'    Debug.Print "B " & n & " " & items(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('treats an array passed in any call position as possibly allocated', () => {
		const src =
			'Public Sub T(ByVal loader As Object)\n' +
			'    Dim a() As Long\n' +
			'    Dim b() As Long\n' +
			'    Dim c() As Long\n' +
			'    Dim d() As Long\n' +
			'    If Load(a) Then Debug.Print a(0)\n' +
			'    loader.Fill b\n' +
			'    Debug.Print b(0)\n' +
			'    Debug.Print Load(c)\n' +
			'    Debug.Print c(0)\n' +
			'    Call loader.Fill(1, d)\n' +
			'    Debug.Print d(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('treats a With-relative call and a named argument as passing the array', () => {
		const src =
			'Public Sub T(ByVal loader As Object)\n' +
			'    Dim a() As Long\n' +
			'    Dim b() As Long\n' +
			'    With loader\n' +
			'        .Fill a\n' +
			'    End With\n' +
			'    Debug.Print a(0)\n' +
			'    Fill target:=b\n' +
			'    Debug.Print b(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('still flags an indexed access inside a call argument and after a bound call', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim n As Long\n' +
			'    n = Load(values(0))\n' +
			'    Debug.Print UBound(values)\n' +
			'    Debug.Print values(1)\n' +
			'End Sub\n';

		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');
		expect(hits).toHaveLength(3);
	});

	it('stays quiet after nested block allocation makes straight-line state unknown', () => {
		const src =
			'Public Sub T(ByVal ready As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    If ready Then\n' +
			'        ReDim values(0 To 1)\n' +
			'    End If\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('flags unallocated access inside a balanced If arm (branch-merge coverage)', () => {
		const src =
			'Public Sub T(ByVal flag As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    If flag Then\n' +
			'        Debug.Print values(0)\n' +
			'    Else\n' +
			'        ReDim values(0 To 1)\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'unallocated-dynamic-array-access');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('values');
	});

	it('does not flag a branch-arm access when the array was ReDim before the If', () => {
		const src =
			'Public Sub T(ByVal flag As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    ReDim values(0 To 2)\n' +
			'    If flag Then\n' +
			'        Debug.Print values(0)\n' +
			'    Else\n' +
			'        Debug.Print values(1)\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('stays quiet after a balanced If allocates the array on every arm', () => {
		const src =
			'Public Sub T(ByVal flag As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    If flag Then\n' +
			'        ReDim values(0 To 1)\n' +
			'    Else\n' +
			'        ReDim values(0 To 2)\n' +
			'    End If\n' +
			'    Debug.Print values(0)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('falls back to conservative flow when On Error is present', () => {
		const src =
			'Public Sub T(ByVal flag As Boolean)\n' +
			'    Dim values() As Long\n' +
			'    On Error Resume Next\n' +
			'    If flag Then\n' +
			'        Debug.Print values(0)\n' +
			'    Else\n' +
			'        ReDim values(0 To 1)\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('does not enter loop bodies (For stays conservative)', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'    Dim i As Long\n' +
			'    For i = 1 To 3\n' +
			'        Debug.Print values(i)\n' +
			'    Next i\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});

	it('ignores inactive conditional-compilation indexed access', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'    Dim values() As Long\n' +
			'#If Enabled Then\n' +
			'    Debug.Print values(0)\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});
});

describe('analyzeModule - array Erase', () => {
	it('flags an Erase arithmetic expression', () => {
		const src =
			'Sub T()\n' +
			'    Erase 1 + 2\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-erase-target');

		expectDiagnostic(src, hits, 'invalid-erase-target', { severity: 'error', span: '1 + 2' });
	});

	it('flags expression targets in comma-separated Erase lists', () => {
		const src =
			'Sub T()\n' +
			'    Erase Values, other + 1, "bad"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'invalid-erase-target');

		expect(hits.map((hit) => spanText(src, hit))).toEqual(['other + 1', '"bad"']);
	});

	it('accepts variable-like Erase targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Erase Values\n' +
			'    Erase Values, OtherValues\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'invalid-erase-target')).toHaveLength(0);
	});

	it('flags Erase of declared non-Variant scalar variables', () => {
		const src =
			'Private ModuleValue As Long\n' +
			'Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim Value As Long\n' +
			'    Erase obj, Value, ModuleValue\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'erase-requires-array');

		expectDiagnostics(src, hits, 'erase-requires-array', [
			{ severity: 'error', span: 'obj' },
			{ span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('uses visible exported scalar globals for Erase target shapes', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValue, SharedObject\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValue As Long\n' +
					'Public SharedObject As Object\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'erase-requires-array');

		expect(hits).toHaveLength(2);
		expect(hits.map((hit) => spanText(caller, hit))).toEqual(['SharedValue', 'SharedObject']);
	});

	it('accepts Erase of visible exported arrays and Variant globals', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValues, SharedFlexible\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValues() As Long\n' +
					'Public SharedFlexible As Variant\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('lets local array and Variant declarations shadow exported Erase scalars', () => {
		const caller =
			'Sub T()\n' +
			'    Dim SharedValue() As Long\n' +
			'    Dim SharedObject As Variant\n' +
			'    Erase SharedValue, SharedObject\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValue As Long\n' +
					'Public SharedObject As Object\n',
			},
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported Erase targets quiet', () => {
		const caller =
			'Sub T()\n' +
			'    Erase SharedValue\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As Object\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'erase-requires-array')).toHaveLength(0);
	});

	it('accepts Erase of arrays, Variants, unresolved names, and non-simple targets', () => {
		const src =
			'Sub T()\n' +
			'    Dim Values() As Long\n' +
			'    Dim FixedValues(1 To 3) As Long\n' +
			'    Dim Flexible As Variant\n' +
			'    Dim ImplicitVariant\n' +
			'    Erase Values, FixedValues, Flexible, ImplicitVariant, UnknownValues\n' +
			'    Erase Settings.Values, Values(1)\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'erase-requires-array')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking Erase scalar targets', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    Erase Values\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'erase-requires-array',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'erase-requires-array',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - array-subscript-out-of-bounds (RUNTIME_006)', () => {
	const CODE = 'array-subscript-out-of-bounds';

	it('flags a literal subscript above the declared upper bound', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\n    a(11) = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].severity).toBe('error');
		expect(spanText(src, hits[0])).toBe('11');
		expect(hits[0].message).toContain("Run-time error '9'");
	});

	it('flags a subscript below an explicit literal lower bound', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\n    a(0) = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('0');
	});

	it('flags a subscript above a single-bound upper (Option-Base-independent)', () => {
		const src = 'Sub T()\n    Dim a(5) As Long\n    a(6) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(1);
	});

	it('flags a negative subscript', () => {
		const src = 'Sub T()\n    Dim a(5) As Long\n    a(-1) = 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('-1');
	});

	it('flags an out-of-bounds access inside a loop or If (all active statements)', () => {
		const src =
			'Sub T()\n' +
			'    Dim a(1 To 10) As Long\n' +
			'    Dim i As Long\n' +
			'    For i = 1 To 3\n' +
			'        a(11) = i\n' +
			'    Next i\n' +
			'    If i > 0 Then a(20) = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(2);
	});

	it('flags a folded constant-expression subscript (5 + 6 on 1 To 10)', () => {
		const src = 'Sub T()\n    Dim a(1 To 10) As Long\n    a(5 + 6) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(1);
	});

	it('stays quiet at the boundary and within bounds', () => {
		const src =
			'Sub T()\n' +
			'    Dim a(1 To 10) As Long\n' +
			'    a(1) = 1\n' +
			'    a(10) = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a low subscript on a single-bound array (Option Base unmodeled)', () => {
		const src = 'Sub T()\n    Dim a(5) As Long\n    a(0) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for variable and Const subscripts (not provable)', () => {
		const src =
			'Sub T()\n' +
			'    Const MAX As Long = 99\n' +
			'    Dim a(1 To 10) As Long\n' +
			'    Dim i As Long\n' +
			'    i = 11\n' +
			'    a(i) = 1\n' +
			'    a(MAX) = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a dynamic array sized by ReDim', () => {
		const src =
			'Sub T()\n' +
			'    Dim a() As Long\n' +
			'    ReDim a(1 To 3)\n' +
			'    a(5) = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a multi-dimension array', () => {
		const src = 'Sub T()\n    Dim a(1 To 3, 1 To 4) As Long\n    a(99) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a non-constant (Const-sized) upper bound', () => {
		const src =
			'Sub T()\n' +
			'    Const N As Long = 10\n' +
			'    Dim a(1 To N) As Long\n' +
			'    a(99) = 1\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for a member access that resembles an index (obj.a(11))', () => {
		const src =
			'Sub T()\n' +
			'    Dim a(1 To 10) As Long\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Nothing\n' +
			'    obj.a(11) = 1\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});

	it('stays quiet for an out-of-bounds access in an inactive #If branch', () => {
		const src =
			'Sub T()\n' +
			'    Dim a(1 To 10) As Long\n' +
			'#If 0 Then\n' +
			'    a(11) = 1\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('analyzeModule - array index rules ignore the bang (!) member operator', () => {
	it('does not treat d!b(N) as a local-array subscript (fixed or dynamic)', () => {
		// d!b is d.Item("b") (bang/dictionary access); the identifier after ! is a
		// string key, not the local array b, so neither array rule may fire.
		const fixedSrc =
			'Sub T()\n' +
			'    Dim b(5) As Long\n' +
			'    Dim d As Object\n' +
			'    Set d = Nothing\n' +
			'    Debug.Print d!b(11)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(fixedSrc), 'array-subscript-out-of-bounds')).toHaveLength(0);
		const dynSrc =
			'Sub T()\n' +
			'    Dim b() As Long\n' +
			'    Dim d As Object\n' +
			'    Set d = Nothing\n' +
			'    Debug.Print d!b(11)\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(dynSrc), 'unallocated-dynamic-array-access')).toHaveLength(0);
	});
});
