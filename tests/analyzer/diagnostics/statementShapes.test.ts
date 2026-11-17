// End, Stop, Error, On Error and On ... GoTo in a shape the VBE refuses, each
// a Syntax error, or "Expected: end of statement" on a procedure's own line. Measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20527) through the VBE oracle; the cases are the statement_shape_* fixtures
// in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule, diagnosticMetadataForCode } from '../../../src/analyzer';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: { source: string }[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('statement_shape_'));

const shapes = (source: string, host = 'excel') => byCode(analyzeModule(source, { host }), 'malformed-statement');
const compileErrors = (source: string) => analyzeModule(source, { host: 'excel' })
	.filter((diagnostic) => diagnosticMetadataForCode(diagnostic.code ?? '')?.diagnosticKind === 'compile-error')
	.map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`);

function procedure(...lines: string[]): string {
	return ['Option Explicit', 'Sub Main()', '    Dim v As Variant, n As Long', ...lines, 'End Sub', ''].join('\r\n');
}

describe('End, Stop, Error and On statements in a shape the VBE refuses', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((kase) => kase.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((kase) => [kase.id, kase] as const))('%s', (_id, kase) => {
		expect(shapes(kase.modules[0].source, kase.host).length).toBe(kase.expected === 'rejected' ? 1 : 0);
	});

	it.each([
		['    End 1', 'End 1', "'End' stops the program and takes nothing after it, so '1' cannot follow it."],
		['    If v Then End Else Beep', 'End Else', "so 'Else' cannot follow it"],
		['    End While', 'End While', 'a While loop ends at Wend'],
		['    Stop n', 'Stop n', "'Stop' takes no argument."],
		['    Error', 'Error', 'needs an error number'],
		['    Error 5, 6', 'Error 5, 6', 'takes one error number'],
		['    On Error GoTo -2', 'On Error GoTo -2', 'one line label, line number, 0 or -1'],
		['    On Error Resume 0', 'On Error Resume 0', 'takes only Next'],
		['    On Error Resume Next 1', 'On Error Resume Next 1', 'and nothing after it'],
		['    On Error GoTo 0 1', 'On Error GoTo 0 1', 'and nothing after it'],
		['    On n GoSub A, , B\r\nA:\r\nB:', 'GoSub A, , B', 'in each place of its list'],
	])('reports %j', (line, span, message) => {
		const src = procedure(line);
		expectDiagnostic(src, shapes(src), 'malformed-statement', { span, message: `${message}` });
	});

	it('reads the statement after a colon, a label, a line number and in nested one-line Ifs', () => {
		for (const line of [
			'    Beep: Stop 1', 'L: Error', '10 End 1', '    If v Then If n Then Stop 1', '    If v Then Beep Else On Error Resume', '    Beep: Error: Beep',
			'    If v Then Beep: Error Else Beep', '    If v Then Beep: Beep Else Stop 1', '    If v Then Beep: Beep Else End 1',
			'    If v Then Beep: End Else Beep', '    If v Then Beep Else Beep: Stop 1',
			'    If v Then\r\n        Beep\r\n    Else If n Then Stop 1\r\n    End If', '    If v Then\r\n        Beep\r\n    Else Stop 1\r\n    End If',
			'    If v Then\r\n        Beep\r\n    Else Error\r\n    End If', '    If v Then\r\n        Beep\r\n    Else On Error Resume Next 1\r\n    End If',
		]) {
			expect(shapes(procedure(line)), line).toHaveLength(1);
		}
	});

	it('reads an ElseIf line, a procedure\'s own line and a one-line If before a block If header', () => {
		for (const src of [
			procedure('    If False Then', '        Beep', '    ElseIf True Then Stop 1', '    End If'),
			procedure('    If False Then', '        Beep', '    ElseIf True Then Error', '    End If'),
			['Option Explicit', 'Sub Other(): Stop 1: End Sub', ''].join('\r\n'),
			procedure('    If True Then Stop 1: If True Then', '        Beep', '    End If'),
		]) {
			expect(shapes(src), src).toHaveLength(1);
		}
		for (const src of [
			procedure('    If False Then', '        Beep', '    ElseIf True Then Stop', '    End If'),
			procedure('    If False Then', '        Beep', '    ElseIf True Then', '        Stop', '    End If'),
			['Option Explicit', 'Sub Other(): Stop: End Sub', ''].join('\r\n'),
		]) {
			expect(compileErrors(src), src).toEqual([]);
		}
	});

	it('leaves a reserved word after GoTo or Resume to the jump-target check, one report each', () => {
		for (const line of ['    On Error GoTo Next 1', '    On Error Resume Stop', '    On Error Resume If', '    If v Then On Error Resume Else Beep', '    If v Then Beep: On Error Resume Else Beep', '    If v Then On Error GoTo Else Beep']) {
			const found = shapes(procedure(line));
			expect(found, line).toHaveLength(1);
			expect(found[0].message, line).toContain('cannot name a line label');
		}
	});

	it('leaves the forms that compile alone', () => {
		for (const line of [
			'    End', "    End ' note", '    If v Then End', '    If v Then Beep Else End', '    If v Then End: Beep',
			'    Stop', '    If v Then Stop Else Beep', 'Stop: Beep',
			'    Error 5', '    Error (5)', '    If v Then Error 5 Else Beep', 'Error: Beep', '10 Error: Beep', '    Error _\r\n        5', '    Error$ 5',
			'    v = Error', '    v = Error(5)', '    Debug.Print Error',
			'    On Error GoTo 0', '    On Error GoTo -1', '    On Error GoTo - 1', '    On Error GoTo L\r\nL:', '    On Error GoTo 10\r\n10 Beep',
			'    On Error Resume Next', '    On Local Error Resume Next', '    If v Then On Error Resume Next Else Beep',
			'    On n GoTo A, B\r\nA:\r\nB:', '    On n GoSub A\r\n    Exit Sub\r\nA:\r\n    Return',
			'    Dim r As Object\r\n    Set r = Nothing\r\n    If v Then r.End(1).Select', '    With Nothing\r\n        .End 1\r\n        .Stop 1\r\n    End With',
			'    If v Then Beep: Stop Else Beep', '    If v Then\r\n        End\r\n    End If',
			'    If v Then\r\n        Beep\r\n    Else If n Then Stop Else Beep\r\n    End If', '    If v Then\r\n        Beep\r\n    Else Stop\r\n    End If',
		]) {
			expect(compileErrors(procedure(line)), line).toEqual([]);
		}
	});

	it('leaves an inactive #If arm alone', () => {
		expect(shapes(procedure('#If False Then', '    End 1', '    Stop 1', '    Error', '#End If'))).toEqual([]);
	});
});
