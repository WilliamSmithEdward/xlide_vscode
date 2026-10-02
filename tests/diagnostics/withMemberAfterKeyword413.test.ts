// Diagnostics tests: a With member read after a keyword, `If .Count > 3`,
// is no statement calling the property (issue #413, a regression from
// #266). Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CLASS = 'Option Explicit\nPublic Property Get State(ByVal k As String) As Long\n    State = 5\nEnd Property\nPublic Property Get Count() As Long\n    Count = 5\nEnd Property\nPublic Sub SetState(ByVal k As String, ByVal v As Long)\nEnd Sub\n';

function propertyUse(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1, i As Long\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, { moduleName: 'Class1', source: CLASS, type: 'class' }], 'Module1')
		.filter((diag) => diag.code === 'invalid-property-use')
		.map((diag) => diag.message);
}

describe('a With member after a keyword', () => {
	it('is read, not called', () => {
		for (const lines of [
			['With c', '    If .State("q") > 3 Then .SetState "q", 3', 'End With'],
			['With c', '    If .Count > 3 Then Main = 1', 'End With'],
			['With c', '    Debug.Print .Count', 'End With'],
			['With c', '    If False Then', '    ElseIf .Count > 3 Then', '        Main = 1', '    End If', 'End With'],
			['With c', '    Select Case 4', '    Case .Count', '        Main = 1', '    End Select', 'End With'],
			['With c', '    For i = 1 To .Count', '    Next', 'End With'],
		]) {
			expect(propertyUse(...lines), lines.join(' / ')).toEqual([]);
		}
	});

	it('still reports a property called as a statement', () => {
		expect(propertyUse('c.Count')).toEqual([expect.stringContaining("'c.Count' is a property")]);
	});
});
