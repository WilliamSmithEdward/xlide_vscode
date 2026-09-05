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

	it('treats an object passed to a call in any position as possibly Set (#70 class)', () => {
		const src =
			'Private Function TryGet(ByRef target As Object) As Boolean\n' +
			'    Set target = New Collection\n' +
			'    TryGet = True\n' +
			'End Function\n' +
			'Public Sub T(ByVal factory As Object)\n' +
			'    Dim a As Object\n' +
			'    Dim b As Object\n' +
			'    Dim c As Object\n' +
			'    Dim d As Object\n' +
			'    Dim n As Long\n' +
			'    If TryGet(a) Then a.ToString\n' +
			'    n = TryGet(b)\n' +
			'    b.ToString\n' +
			'    Debug.Print TryGet(c)\n' +
			'    c.ToString\n' +
			'    factory.Build d\n' +
			'    d.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('treats an object passed to a call inside a branch as possibly Set after the merge', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        TryGet obj\n' +
			'    End If\n' +
			'    obj.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('treats a With-relative call and a named argument as passing the object', () => {
		const src =
			'Public Sub T(ByVal factory As Object)\n' +
			'    Dim a As Object\n' +
			'    Dim b As Object\n' +
			'    With factory\n' +
			'        .Build a\n' +
			'    End With\n' +
			'    a.ToString\n' +
			'    Build target:=b\n' +
			'    b.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('still flags member access inside a call argument, an Is test, and a read-only intrinsic', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Dim n As Long\n' +
			'    n = Load(obj.Name)\n' +
			'    If obj Is Nothing Then Debug.Print TypeName(obj)\n' +
			'    obj.ToString\n' +
			'End Sub\n';

		const hits = byCode(analyzeModule(src), 'object-variable-not-set');
		const lineOf = (offset: number): number => src.slice(0, offset).split('\n').length;
		expect(hits.map((hit) => [spanText(src, hit), lineOf(hit.span.start)])).toEqual([
			['obj', 4],
			['obj', 6],
		]);
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

	it('flags Nothing member access inside a balanced If arm (branch-merge coverage)', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        obj.ToString\n' +
			'    Else\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
			'End Sub\n';

		expectDiagnostic(src, analyzeModule(src), 'object-variable-not-set', { span: 'obj' });
	});

	it('does not flag a branch-arm access when the object was Set before the If', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = New Collection\n' +
			'    If Ready Then\n' +
			'        obj.ToString\n' +
			'    Else\n' +
			'        obj.ToString\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('stays quiet after a balanced If sets the object on every arm', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        Set obj = New Collection\n' +
			'    Else\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
			'    obj.ToString\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('does not enter a one-armed If (no else stays conservative)', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    If Ready Then\n' +
			'        obj.ToString\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('falls back to conservative flow when On Error is present', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    On Error Resume Next\n' +
			'    If Ready Then\n' +
			'        obj.ToString\n' +
			'    Else\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'object-variable-not-set')).toHaveLength(0);
	});

	it('falls back to conservative flow when a GoTo/label is present', () => {
		const src =
			'Public Sub T()\n' +
			'    Dim obj As Object\n' +
			'    GoTo Skip\n' +
			'Skip:\n' +
			'    If Ready Then\n' +
			'        obj.ToString\n' +
			'    Else\n' +
			'        Set obj = New Collection\n' +
			'    End If\n' +
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
