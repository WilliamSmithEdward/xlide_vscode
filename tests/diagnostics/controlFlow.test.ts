// Diagnostics tests: controlFlow rule family.
// Split verbatim from tests/vbaDiagnostics.test.ts (audit #107).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule, projectClassMembers, type ProjectTestModule } from './helpers';

describe('analyzeModule - Exit statement matches procedure', () => {
	it('flags Exit Function inside a Sub', () => {
		const src = 'Sub T()\n    Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
		expect(hits[0].severity).toBe('error');
	});

	it('flags mismatched Exit statements after numeric line labels', () => {
		const src = 'Sub T()\n10 Exit Function\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'exit-wrong-proc');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Exit Function');
	});

	it('flags Exit Sub inside a Function', () => {
		const src = 'Function F() As Long\n    Exit Sub\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('flags Exit Sub inside a Property Let', () => {
		// Property Let/Set reject Exit Sub and Exit Function (oracle-verified);
		// only Property Get is lenient.
		const src = 'Property Let Name(v As String)\n    Exit Sub\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('flags Exit Function inside a Property Set', () => {
		const src = 'Property Set Obj(v As Object)\n    Exit Function\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(1);
	});

	it('accepts Exit Function inside a Property Get (the stdEnumerator pattern)', () => {
		// VBE special-cases Property Get: it is value-returning, so Exit Function
		// compiles there (oracle-verified).
		const src = 'Property Get Item() As Variant\n    Exit Function\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
	});

	it('accepts Exit Sub inside a Property Get', () => {
		const src = 'Property Get Name() As String\n    Exit Sub\nEnd Property\n';
		expect(byCode(analyzeModule(src), 'exit-wrong-proc')).toHaveLength(0);
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

describe('analyzeModule - procedure labels', () => {
	it('accepts forward and backward GoTo and GoSub labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'    GoTo done\n' +
			'start:\n' +
			'    GoSub done\n' +
			'    GoTo start\n' +
			'done:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('flags missing GoTo, GoSub, Resume, and On Error labels', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingGoTo\n' +
			'    GoSub MissingGoSub\n' +
			'    Resume MissingResume\n' +
			'    On Error GoTo MissingHandler\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual([
			'MissingGoTo',
			'MissingGoSub',
			'MissingResume',
			'MissingHandler',
		]);
		expect(hits.every((hit) => hit.severity === 'error')).toBe(true);
	});

	it('keeps labels scoped to their enclosing procedure', () => {
		const src =
			'Sub A()\n' +
			'    GoTo Done\n' +
			'End Sub\n' +
			'Sub B()\n' +
			'Done:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Done');
	});

	it('flags duplicate named labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'StartHere:\n' +
			'    Debug.Print "first"\n' +
			'StartHere:\n' +
			'    Debug.Print "second"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-label');

		expectDiagnostic(src, hits, 'duplicate-label', { severity: 'error', span: 'StartHere' });
	});

	it('flags duplicate normalized numeric labels in the same procedure', () => {
		const src =
			'Sub T()\n' +
			'010 Debug.Print "first"\n' +
			'10 Debug.Print "second"\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'duplicate-label');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('10');
	});

	it('allows the same label name in separate procedures', () => {
		const src =
			'Sub A()\n' +
			'Done:\n' +
			'End Sub\n' +
			'Sub B()\n' +
			'Done:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'duplicate-label')).toHaveLength(0);
	});

	it('accepts explicit On Error forms that do not name a label', () => {
		const src =
			'Sub T()\n' +
			'    On Error Resume Next\n' +
			'    On Error GoTo 0\n' +
			'    On Error GoTo -1\n' +
			'    Resume\n' +
			'    Resume Next\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('accepts On Error disable forms after line numbers and inside single-line If statements', () => {
		const src =
			'Sub T(ByVal flag As Boolean)\n' +
			'10 On Error GoTo 0\n' +
			'20 If flag Then On Error GoTo 0 Else On Error GoTo Handler\n' +
			'30 On Error GoTo -1\n' +
			'Handler:\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('validates On n GoTo and On n GoSub label lists', () => {
		const src =
			'Sub T(ByVal n As Long)\n' +
			'    On n GoTo First, MissingJump\n' +
			'    On n GoSub First, MissingSub\n' +
			'First:\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingJump', 'MissingSub']);
	});

	it('validates labels on both sides of a single-line If Else statement', () => {
		const src = 'Sub T(ByVal flag As Boolean)\n    If flag Then GoTo MissingA Else GoTo MissingB\nEnd Sub\n';
		const hits = byCode(analyzeModule(src), 'undefined-label');
		expect(hits.map((hit) => spanText(src, hit))).toEqual(['MissingA', 'MissingB']);
	});

	it('accepts numeric line labels and references', () => {
		const src =
			'Sub T()\n' +
			'    GoTo 10\n' +
			'10:\n' +
			'    On 1 GoTo 10\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('does not use labels from inactive conditional-compilation branches', () => {
		const src =
			'Sub T()\n' +
			'    GoTo MissingWhenInactive\n' +
			'#If VBA7 Then\n' +
			'MissingWhenInactive:\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				conditionalCompilation: { compilerConstants: { VBA7: false } },
			}),
			'undefined-label',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('MissingWhenInactive');
	});

	it('does not treat inactive duplicate labels as active duplicates', () => {
		const src =
			'Sub T()\n' +
			'StartHere:\n' +
			'#If VBA7 Then\n' +
			'StartHere:\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'duplicate-label',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - statement context', () => {
	it('flags #ElseIf after #Else in the same conditional-compilation block', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#Else\n' +
			'    Debug.Print 1\n' +
			'#ElseIf True Then\n' +
			'    Debug.Print 2\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { severity: 'error', span: '#ElseIf' });
	});

	it('flags duplicate #Else branches in a conditional-compilation block', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#Else\n' +
			'    Debug.Print 1\n' +
			'#Else\n' +
			'    Debug.Print 2\n' +
			'#End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: '#Else' });
	});

	it('accepts conditional-compilation #ElseIf branches before #Else and in nested blocks', () => {
		const src =
			'Sub T()\n' +
			'#If False Then\n' +
			'    Debug.Print 0\n' +
			'#ElseIf True Then\n' +
			'    Debug.Print 1\n' +
			'#Else\n' +
			'    #If True Then\n' +
			'        Debug.Print 2\n' +
			'    #ElseIf False Then\n' +
			'        Debug.Print 3\n' +
			'    #End If\n' +
			'#End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'else-branch-order')).toHaveLength(0);
	});

	it('flags ElseIf after Else in the same If block', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    Else\n' +
			'        Debug.Print 1\n' +
			'    ElseIf True Then\n' +
			'        Debug.Print 2\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: 'ElseIf' });
	});

	it('flags duplicate Else branches in the same If block', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    Else\n' +
			'        Debug.Print 1\n' +
			'    Else\n' +
			'        Debug.Print 2\n' +
			'    End If\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'else-branch-order');

		expectDiagnostic(src, hits, 'else-branch-order', { span: 'Else' });
	});

	it('accepts ElseIf branches before Else and nested If blocks after Else', () => {
		const src =
			'Sub T()\n' +
			'    If False Then\n' +
			'        Debug.Print 0\n' +
			'    ElseIf True Then\n' +
			'        Debug.Print 1\n' +
			'    Else\n' +
			'        If True Then\n' +
			'            Debug.Print 2\n' +
			'        ElseIf False Then\n' +
			'            Debug.Print 3\n' +
			'        End If\n' +
			'    End If\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'else-branch-order')).toHaveLength(0);
	});

	it('does not report inactive If branch-order diagnostics', () => {
		const src =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    If False Then\n' +
			'    Else\n' +
			'    ElseIf True Then\n' +
			'    End If\n' +
			'#End If\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'else-branch-order',
			),
		).toHaveLength(0);
	});

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

	it('accepts line-numbered Case statements inside Select Case', () => {
		const src =
			'Sub T()\n' +
			'10 Select Case x\n' +
			'20 Case 1, 2\n' +
			'30 x = 3\n' +
			'40 Case Else\n' +
			'50 x = 4\n' +
			'60 End Select\n' +
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

	it('uses the With receiver for unknown source-backed class members', () => {
		const person = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Delete\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'member-not-found',
		);
		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
	});

	it('uses the With receiver for source-backed class member assignment rules', () => {
		const person =
			'Private mAge As Integer\n' +
			'Public Property Get Age() As Integer\n' +
			'    Age = mAge\n' +
			'End Property\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Age = 2\n' +
			'    End With\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
				]),
			}),
			'readonly-member-assignment',
		);
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('Age');
	});

	it('uses the With receiver for source-backed class member call rules', () => {
		const person = 'Public Sub Save(ByVal Count As Long)\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        .Save "bad"\n' +
			'        .Save()\n' +
			'    End With\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
			]),
		});
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(1);
		expect(byCode(diagnostics, 'call-statement-forbids-parens')).toHaveLength(1);
	});

	it('uses nested With receivers from outer leading-dot expressions', () => {
		const person = 'Public Property Get Child() As Child\nEnd Property\n';
		const child =
			'Public Property Get Age() As Integer\nEnd Property\n' +
			'Public Sub Save(ByVal Count As Long)\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    With p\n' +
			'        With .Child\n' +
			'            .Delete\n' +
			'            .Age = 2\n' +
			'            .Save "bad"\n' +
			'        End With\n' +
			'    End With\n' +
			'End Sub\n';
		const diagnostics = analyzeModule(src, {
			projectClassMembers: projectClassMembers([
				{ moduleName: 'Person', moduleKind: 'class', source: person },
				{ moduleName: 'Child', moduleKind: 'class', source: child },
			]),
		});
		expectDiagnostic(src, diagnostics, 'member-not-found', { span: 'Delete' });
		expect(byCode(diagnostics, 'readonly-member-assignment')).toHaveLength(1);
		expect(byCode(diagnostics, 'argument-type-mismatch')).toHaveLength(1);
	});

	it('uses parenthesized member receivers for source-backed class diagnostics', () => {
		const person = 'Public Property Get Child() As Child\nEnd Property\n';
		const child = 'Public Sub Save()\nEnd Sub\n';
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    (p.Child).Delete\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeModule(src, {
				projectClassMembers: projectClassMembers([
					{ moduleName: 'Person', moduleKind: 'class', source: person },
					{ moduleName: 'Child', moduleKind: 'class', source: child },
				]),
			}),
			'member-not-found',
		);

		expectDiagnostic(src, hits, 'member-not-found', { span: 'Delete' });
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

	it('flags a Next variable that does not match the active For loop', () => {
		const src =
			'Sub T()\n' +
			'    Dim i As Long\n' +
			'    Dim j As Long\n' +
			'    For i = 1 To 3\n' +
			'        Debug.Print i\n' +
			'    Next j\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'next-variable-mismatch');

		expectDiagnostic(src, hits, 'next-variable-mismatch', { severity: 'error', span: 'j' });
	});

	it('flags a Next variable that does not match the active For Each loop', () => {
		const src =
			'Sub T()\n' +
			'    For Each item In items\n' +
			'    Next other\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'next-variable-mismatch');

		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('other');
	});

	it('accepts omitted, matching, and nested Next variables', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n' +
			'        For j = 1 To 3\n' +
			'        Next J\n' +
			'    Next i\n' +
			'    For k = 1 To 3\n' +
			'    Next\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'next-variable-mismatch')).toHaveLength(0);
	});

	it('does not report Next variable mismatches from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    For i = 1 To 3\n' +
			'#If VBA7 Then\n' +
			'    Next j\n' +
			'#Else\n' +
			'    Next i\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'next-variable-mismatch',
			),
		).toHaveLength(0);
	});

	it('flags a For Each control variable declared as a scalar intrinsic type', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Long\n' +
			'    For Each item In Array(1, 2, 3)\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-control-variable-type');

		expectDiagnostic(src, hits, 'for-each-control-variable-type', {
			severity: 'error',
			span: 'item',
		});
	});

	it('flags an array variable used as a For Each control variable', () => {
		const src =
			'Sub T()\n' +
			'    Dim values() As Variant\n' +
			'    For Each values In Array(1, 2, 3)\n' +
			'        Debug.Print values\n' +
			'    Next values\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-control-variable-type');

		expectDiagnostic(src, hits, 'for-each-control-variable-type', { span: 'values' });
	});

	it('uses visible exported scalar and array globals for For Each control variables', () => {
		const caller =
			'Sub T()\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'    For Each SharedValues In Array(1, 2, 3)\n' +
			'    Next SharedValues\n' +
			'End Sub\n';
		const diagnostics = analyzeProjectModule(caller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedItem As Long\n' +
					'Public SharedValues() As Variant\n',
			},
		], 'Caller');
		const hits = byCode(diagnostics, 'for-each-control-variable-type');

		expectDiagnostics(caller, hits, 'for-each-control-variable-type', [
			{ span: 'SharedItem', message: 'Long' },
			{ span: 'SharedValues', message: 'array variable' },
		]);
	});

	it('keeps local Variant shadows and ambiguous exported For Each controls quiet', () => {
		const shadowCaller =
			'Sub T()\n' +
			'    Dim SharedItem As Variant\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedItem As Long\n' },
		], 'Caller'), 'for-each-control-variable-type')).toHaveLength(0);

		const ambiguousCaller =
			'Sub T()\n' +
			'    For Each SharedItem In Array(1, 2, 3)\n' +
			'    Next SharedItem\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedItem As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedItem As String\n' },
		], 'Caller'), 'for-each-control-variable-type')).toHaveLength(0);
	});

	it('accepts Variant, Object, and host object For Each control variables', () => {
		const src =
			'Sub T()\n' +
			'    Dim value As Variant\n' +
			'    For Each value In Array(1, 2, 3)\n' +
			'    Next value\n' +
			'    Dim obj As Object\n' +
			'    For Each obj In items\n' +
			'    Next obj\n' +
			'    Dim ws As Worksheet\n' +
			'    For Each ws In ThisWorkbook.Worksheets\n' +
			'    Next ws\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'for-each-control-variable-type')).toHaveLength(0);
	});

	it('flags a project UDT used as a For Each control variable when type metadata is available', () => {
		const modules: ProjectTestModule[] = [
			{
				moduleName: 'Types',
				moduleKind: 'standard',
				source: 'Public Type TItem\n    Value As Long\nEnd Type\n',
			},
			{ moduleName: 'Module1', moduleKind: 'standard', source: '' },
		];
		const src =
			'Sub T()\n' +
			'    Dim item As TItem\n' +
			'    For Each item In items\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(src, modules, 'Module1'),
			'for-each-control-variable-type',
		);

		expectDiagnostic(src, hits, 'for-each-control-variable-type', { span: 'item' });
	});

	it('does not report For Each control variable type diagnostics from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Long\n' +
			'#If VBA7 Then\n' +
			'    For Each item In Array(1, 2, 3)\n' +
			'    Next item\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-control-variable-type',
			),
		).toHaveLength(0);
	});

	it('flags a For Each source declared as a scalar intrinsic type', () => {
		const src =
			'Private ModuleValue As String\n' +
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim Value As Long\n' +
			'    For Each item In Value\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'    For Each item In ModuleValue\n' +
			'        Debug.Print item\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(analyzeModule(src), 'for-each-source-type');

		expectDiagnostics(src, hits, 'for-each-source-type', [
			{ severity: 'error', span: 'Value' },
			{ span: 'ModuleValue' },
		]);
	});

	it('uses visible exported scalar globals for For Each source types', () => {
		const caller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		const hits = byCode(
			analyzeProjectModule(caller, [
				{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
			], 'Caller'),
			'for-each-source-type',
		);

		expectDiagnostic(caller, hits, 'for-each-source-type', { span: 'SharedValue' });
	});

	it('keeps exported arrays Variants local shadows and ambiguous For Each sources quiet', () => {
		const acceptedCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValues\n' +
			'    Next item\n' +
			'    For Each item In SharedFlexible\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(acceptedCaller, [
			{
				moduleName: 'Globals',
				source:
					'Public SharedValues() As Long\n' +
					'Public SharedFlexible As Variant\n',
			},
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);

		const shadowCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim SharedValue() As Long\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(shadowCaller, [
			{ moduleName: 'Globals', source: 'Public SharedValue As Long\n' },
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);

		const ambiguousCaller =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    For Each item In SharedValue\n' +
			'    Next item\n' +
			'End Sub\n';
		expect(byCode(analyzeProjectModule(ambiguousCaller, [
			{ moduleName: 'GlobalsA', source: 'Public SharedValue As Long\n' },
			{ moduleName: 'GlobalsB', source: 'Public SharedValue As String\n' },
		], 'Caller'), 'for-each-source-type')).toHaveLength(0);
	});

	it('accepts known array/object-like and unresolved For Each sources', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim values() As Long\n' +
			'    Dim maybeValues As Variant\n' +
			'    Dim obj As Object\n' +
			'    For Each item In values\n' +
			'    Next item\n' +
			'    For Each item In maybeValues\n' +
			'    Next item\n' +
			'    For Each item In obj\n' +
			'    Next item\n' +
			'    For Each item In obj.Items\n' +
			'    Next item\n' +
			'End Sub\n';

		expect(byCode(analyzeModule(src), 'for-each-source-type')).toHaveLength(0);
	});

	it('uses the active conditional branch when checking For Each source types', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'#If VBA7 Then\n' +
			'    Dim Values As Long\n' +
			'#Else\n' +
			'    Dim Values() As Long\n' +
			'#End If\n' +
			'    For Each item In Values\n' +
			'    Next item\n' +
			'End Sub\n';

		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: true } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(1);
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(0);
	});

	it('does not report For Each source diagnostics from inactive branches', () => {
		const src =
			'Sub T()\n' +
			'    Dim item As Variant\n' +
			'    Dim Value As Long\n' +
			'#If VBA7 Then\n' +
			'    For Each item In Value\n' +
			'    Next item\n' +
			'#End If\n' +
			'End Sub\n';
		expect(
			byCode(
				analyzeModule(src, {
					conditionalCompilation: { compilerConstants: { VBA7: false } },
				}),
				'for-each-source-type',
			),
		).toHaveLength(0);
	});
});

describe('analyzeModule - control-flow no-diagnostic boundary controls (rule audit backfill)', () => {
	it('stays quiet for an If without Then in an inactive #If 0 branch', () => {
		const src =
			'Sub T()\n' +
			'#If 0 Then\n' +
			'    If x > 0\n' +
			'        x = 1\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'if-missing-then')).toHaveLength(0);
	});

	it('stays quiet for Exit For and Exit Do inside an inactive #If 0 branch', () => {
		const src =
			'Sub T()\n' +
			'#If 0 Then\n' +
			'    Exit For\n' +
			'    Exit Do\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'exit-outside-block')).toHaveLength(0);
	});
});

describe('analyzeModule - realtime incomplete-expression recovery (Priority 3 closure)', () => {
	// Codes that would indicate the incomplete active line cascaded past its own
	// statement: block balance broke (the closer rules) or a statement-context
	// error spilled onto the recovered statement. None may fire on recovery.
	const CASCADE_CODES = [
		'unmatched-block-closer',
		'missing-block-closer',
		'if-missing-then',
		'case-outside-select',
		'else-without-if',
		'member-access-outside-with',
		'exit-outside-block',
	] as const;

	const expectNoCascade = (src: string): void => {
		const diags = analyzeModule(src);
		for (const code of CASCADE_CODES) {
			expect(byCode(diags, code), `expected no ${code}`).toHaveLength(0);
		}
	};

	it('does not cascade a trailing-operator line into the recovered next statement', () => {
		const src = 'Sub T()\n' + '    x = 1 +\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('keeps a trailing-operator line soft, with no block-closer cascade into End Sub', () => {
		const src = 'Sub T()\n' + '    x = 1 +\n' + 'End Sub\n';
		expect(byCode(analyzeModule(src), 'missing-block-closer')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'unmatched-block-closer')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('recovers a trailing-operator line at the next procedure boundary', () => {
		const src =
			'Sub T()\n' + '    x = 1 +\n' + 'End Sub\n' + 'Sub U()\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('recovers a partial call after the next statement without a block cascade', () => {
		const src = 'Sub T()\n' + '    Foo(1,\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('recovers a partial call at the next procedure boundary', () => {
		const src =
			'Sub T()\n' + '    Foo(1,\n' + 'End Sub\n' + 'Sub U()\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
	});

	it('confines a partial member access to its own line and recovers the next statement', () => {
		const src = 'Sub T()\n' + '    obj.\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('confines a partial member access and recovers at the next procedure boundary', () => {
		const src =
			'Sub T()\n' + '    obj.\n' + 'End Sub\n' + 'Sub U()\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('recovers a partial member access inside a With block without breaking block balance', () => {
		const src =
			'Sub T()\n' +
			'    With Range("A1")\n' +
			'        obj.\n' +
			'        .Value = 1\n' +
			'    End With\n' +
			'End Sub\n';
		expectNoCascade(src);
		const hits = byCode(analyzeModule(src), 'invalid-expression-syntax');
		expect(hits).toHaveLength(1);
		expect(spanText(src, hits[0])).toBe('.');
	});

	it('leaves an incomplete string to the lexer without a control-flow cascade', () => {
		const src = 'Sub T()\n' + '    x = "abc\n' + '    y = 2\n' + 'End Sub\n';
		expectNoCascade(src);
		expect(byCode(analyzeModule(src), 'invalid-expression-syntax')).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'unterminated-string')).toHaveLength(1);
	});
});

describe('analyzeModule - bare numeric line labels (undefined-label FP fix)', () => {
	it('treats a bare numeric line label as defined (no false undefined-label)', () => {
		const src =
			'Sub T()\n' +
			'    Dim x As Long\n' +
			'    x = 1\n' +
			'    On x GoTo 100, 200\n' +
			'    Exit Sub\n' +
			'100\n' +
			'    Debug.Print "a"\n' +
			'200\n' +
			'    Debug.Print "b"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(0);
	});

	it('still flags a genuinely undefined numeric label (no over-suppression)', () => {
		const src =
			'Sub T()\n' +
			'    GoTo 999\n' +
			'100\n' +
			'    Debug.Print "a"\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), 'undefined-label')).toHaveLength(1);
	});
});

describe('analyzeModule - reserved keyword in If condition', () => {
	const code = 'if-reserved-keyword-in-condition';

	it('flags a duplicate If keyword where the condition belongs', () => {
		const src = 'Sub T()\n    If If True Then\n    End If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), code)).toHaveLength(1);
	});

	it('flags a second Then keyword in the block header', () => {
		const src = 'Sub T()\n    If True Then Then\n    End If\nEnd Sub\n';
		expect(byCode(analyzeModule(src), code)).toHaveLength(1);
	});

	it('does not flag a well-formed block If / ElseIf / Else', () => {
		const src =
			'Sub T(ByVal a As Boolean, ByVal b As Boolean)\n' +
			'    If a Then\n' +
			'    ElseIf b Then\n' +
			'    Else\n' +
			'    End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), code)).toHaveLength(0);
	});

	it('does not flag legal expression keywords (And/Not/TypeOf/Is)', () => {
		const src =
			'Sub T(ByVal a As Boolean, ByVal b As Boolean, o As Object)\n' +
			'    If a And Not b Then\n' +
			'    End If\n' +
			'    If TypeOf o Is Collection Then\n' +
			'    End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), code)).toHaveLength(0);
	});

	it('does not flag identifiers that merely contain if/then', () => {
		const src =
			'Sub T(ByVal ifCount As Long, ByVal thenFlag As Boolean)\n' +
			'    If ifCount > 0 And thenFlag Then\n' +
			'    End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), code)).toHaveLength(0);
	});

	it('does not flag single-line If or conditional-compilation directives', () => {
		const single = 'Sub T(ByVal x As Boolean)\n    If x Then Exit Sub\nEnd Sub\n';
		expect(byCode(analyzeModule(single), code)).toHaveLength(0);
		const directive =
			'Sub T()\n' +
			'#If VBA7 Then\n' +
			'    Dim x As Long\n' +
			'#End If\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(directive), code)).toHaveLength(0);
	});
});
