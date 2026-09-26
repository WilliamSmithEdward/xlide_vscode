// Block structure the analyzer used to misread (issues #128 to #131). Each
// sample compiles and runs in Excel 16.0 (build 20326), measured through
// pyVBAharness on 2026-09-26; the analyzer must report no compile error on it.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const errors = (src: string): string[] => analyzeVbaModuleSource({ source: src, moduleName: 'Module1', moduleType: 'standard' }).diagnostics
	.filter((d) => d.severity === 'error')
	.map((d) => `${d.code}: ${d.message}`);

describe('a block written in a one-line If tail (issue #128)', () => {
	it('opens and closes on that line for With, For, Do, While, Select and an Else tail', () => {
		const src = [
			'Option Explicit',
			'Function Main() As Variant',
			'    Dim a As Boolean, i As Long, n As Long, k As Long, c As New Collection',
			'    a = True',
			'    If a Then With c: .Add 1: .Add 2: End With',
			'    If a Then For i = 1 To 3: n = n + i: Next i',
			'    If a Then Do: n = n + 1: Loop Until n > 3',
			'    If a Then While n < 10: n = n + 1: Wend',
			'    If a Then Select Case k: Case 0: n = n + 1: Case Else: n = 0: End Select',
			'    If Not a Then n = -1 Else For i = 1 To 2: n = n + 1: Next',
			'    Main = c.Count & ":" & n',
			'End Function',
			'',
		].join('\n');
		expect(errors(src)).toEqual([]);
	});
});

describe('Else with a colon (issue #129)', () => {
	it('is the Else of its If, not a label, twice in one procedure', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n    If n = 1 Then\n        Main = 1\n    Else:\n        Main = 2\n    End If\n    If n = 2 Then\n        Main = Main + 10\n    Else:\n        Main = Main + 20\n    End If\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'duplicate-label')).toHaveLength(0);
		expect(errors(src)).toEqual([]);
	});
});

describe('a block opener or closer written in both arms of an #If (issue #130)', () => {
	it('counts End If, End With and a Select Case header once', () => {
		const closerInBothArms = 'Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n    If n = 0 Then\n        If n = 0 Then\n            Main = "inner"\n#If VBA7 Then\n        End If\n#Else\n        End If\n#End If\n    Else\n        Main = "outer else"\n    End If\nEnd Function\n';
		expect(errors(closerInBothArms)).toEqual([]);
		const openerInBothArms = 'Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n    If n = 0 Then\n        If n = 1 Then\n            Main = "one"\n        Else\n#If VBA7 Then\n            If n = 0 Then\n#Else\n            If n = 2 Then\n#End If\n                Main = "zero"\n            End If\n        End If\n    Else\n        Main = "other"\n    End If\nEnd Function\n';
		expect(errors(openerInBothArms)).toEqual([]);
		const withInBothArms = 'Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection\n    With c\n        .Add 1\n#If VBA7 Then\n    End With\n#Else\n    End With\n#End If\n    With c\n        .Add 2\n    End With\n    Main = c.Count\nEnd Function\n';
		expect(errors(withInBothArms)).toEqual([]);
		const selectInBothArms = 'Option Explicit\nFunction Main() As Variant\n    Dim n As Long\n#If VBA7 Then\n    Select Case n\n#Else\n    Select Case n + 1\n#End If\n        Case 0\n            Main = 7\n        Case Else\n            Main = 8\n    End Select\nEnd Function\n';
		expect(errors(selectInBothArms)).toEqual([]);
		expect(byCode(analyzeModule(selectInBothArms), 'missing-return-assignment')).toHaveLength(0);
	});

	it('still reports a genuine second Else and a genuine stray End If', () => {
		const twoElse = 'Option Explicit\nSub T()\n    If True Then\n    Else\n    Else\n    End If\nEnd Sub\n';
		expect(byCode(analyzeModule(twoElse), 'else-branch-order')).toHaveLength(1);
		const stray = 'Option Explicit\nSub T()\n    Dim n As Long\n    n = 1\n    End If\nEnd Sub\n';
		expect(errors(stray).some((line) => line.includes('End If'))).toBe(true);
	});
});

describe('line numbers ahead of statements (issue #131)', () => {
	it('do not hide a block keyword from the structural pass', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim i As Long, n As Long\n10  For i = 1 To 3\n        n = n + i\n    Next i\n20  If n > 1 Then\n        n = n * 2\n    End If\n30  Do While n > 10\n        n = n - 1\n    Loop\n    Main = n\nEnd Function\n';
		expect(errors(src)).toEqual([]);
		const numberedClosers = 'Option Explicit\nSub T()\n    Dim i As Long\n    For i = 1 To 3\n10  Next i\n    Select Case i\n20  End Select\nEnd Sub\n';
		expect(errors(numberedClosers)).toEqual([]);
	});

	it('do not turn a numbered one-line If with an = condition into an assignment', () => {
		const src = 'Option Explicit\nPublic Sub Report(ByVal count As Long)\n10    On Error GoTo Report_Error\n20    If count = 0 Then MsgBox "No records found", vbInformation\n30    If count = 1 Then Debug.Print "one"\n40    If count = 2 Then Worksheets(1).Range("A1").ClearContents\n50    Exit Sub\nReport_Error:\n60    MsgBox "Error " & Err.Number & " at line " & Erl\nEnd Sub\n';
		expect(errors(src)).toEqual([]);
	});
});
