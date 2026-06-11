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
