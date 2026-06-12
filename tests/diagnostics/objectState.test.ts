// Diagnostics tests: objectState rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

describe('analyzeModule - scalar member access', () => {
	it('flags a trailing dot on a local String variable', () => {
		const src = 'Sub Main()\n    Dim value As String\n    value.\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'scalar-member-access', {
			severity: 'error',
			span: 'value.',
		});
	});

	it('flags named member access on a local Integer variable', () => {
		const src = 'Sub Main()\n    Dim value As Integer\n    value.Length\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'scalar-member-access', { span: 'value.' });
	});

	it('flags scalar parameters and module variables', () => {
		const src =
			'Private moduleName As String\n' +
			'Sub Main(ByVal count As Long)\n' +
			'    moduleName.Length\n' +
			'    count.Value\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['moduleName.', 'count.']);
	});

	it('uses visible exported scalar globals for member receiver types', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
			], 'Caller'),
			'scalar-member-access',
		);

		expectDiagnostic(caller, hits, 'scalar-member-access', { span: 'SharedText.' });
	});

	it('does not leak exported scalar receiver types through local untyped shadows', () => {
		const caller =
			'Sub Main()\n' +
			'    Dim SharedText\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'Globals', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
	});

	it('keeps ambiguous visible exported scalar receivers quiet', () => {
		const caller =
			'Sub Main()\n' +
			'    SharedText.Length\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedText As String\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedText As String\n' },
		], 'Caller');

		expect(byCode(diagnostics, 'scalar-member-access')).toHaveLength(0);
	});

	it('does not treat member names in qualified chains as bare scalar receivers', () => {
		const src =
			'Sub Main()\n' +
			'    Dim Value As String\n' +
			'    Dim item As Variant\n' +
			'    item.Value.Length\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(0);
	});

	it('treats colon-separated declaration and member access statements independently', () => {
		const src = 'Sub Main()\n    Dim value As String: value.Length\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'scalar-member-access');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('value.');
	});

	it('does not flag Variant, Object, or unknown receiver types', () => {
		const src =
			'Sub Main()\n' +
			'    Dim flexible As Variant\n' +
			'    Dim item As Object\n' +
			'    Dim person As Person\n' +
			'    flexible.\n' +
			'    item.\n' +
			'    person.\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'scalar-member-access')).toHaveLength(0);
	});
});

describe('analyzeModule - object variable not set', () => {
	it('flags straight-line local object member access before Set', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    obj.ToString\n' +
			'    Dim ws As Worksheet\n' +
			'    ws.Range("A1").Value = 1\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');

		expectDiagnostics(src, hits, 'object-variable-not-set', [
			{ severity: 'error', span: 'obj' },
			{ span: 'ws' },
		]);
	});

	it('accepts object locals after Set and flags them again after Set Nothing', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = New Collection\n' +
			'    obj.ToString\n' +
			'    Set obj = Nothing\n' +
			'    obj.ToString\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'object-variable-not-set');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('obj');
	});

	it('lets a provable missing member diagnostic supersede runtime not-set', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim ws As Worksheet\n' +
			'    ws.DefinitelyMissingMember\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src);

		expect(byCode(diagnostics, 'object-variable-not-set')).toHaveLength(0);
		const memberHits = byCode(diagnostics, 'member-not-found');
		expect(memberHits).toHaveLength(1);
		expect(spanText(src, memberHits[0])).toBe('DefinitelyMissingMember');
	});

	it('keeps parameters module-level objects Static locals and helper-initialized locals quiet', () => {
		const src =
			'Private moduleObj As Object\n' +
			'Private Sub Initialize(ByRef target As Object)\n' +
			'End Sub\n' +
			'Public Sub T(ByVal param As Object)\n' +
			'    Static cached As Object\n' +
			'    Dim initialized As Object\n' +
			'    Initialize initialized\n' +
			'    moduleObj.ToString\n' +
			'    param.ToString\n' +
			'    cached.ToString\n' +
			'    initialized.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('stays quiet after branch-local initialization makes straight-line state unknown', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
			'    obj.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('ignores inactive conditional-compilation member access', () => {
		const src =
			'#Const Enabled = False\n' +
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'#If Enabled Then\n' +
			'    obj.ToString\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});
});
