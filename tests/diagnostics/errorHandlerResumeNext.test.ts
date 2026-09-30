// On Error Resume Next inside a running error handler (issue #199). Each
// procedure was run in Excel 16.0 through pyVBAharness: the raising ones
// raise the error named to the caller, and the others return normally.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function runtimeErrors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As String\n    Dim x As Double, i As Integer\n${body}\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1' }).diagnostics
		.filter((d) => d.severity === 'error')
		.map((d) => /Run-time error '(\d+)'/.exec(d.message)?.[1] ?? d.code ?? '');
}

const RAISE = '    On Error GoTo Handler\n    Err.Raise 5\n    Exit Function\nHandler:\n';

describe('On Error Resume Next in a running error handler (issue #199)', () => {
	const RAISES: ReadonlyArray<readonly [string, string, string]> = [
		['the handler', `${RAISE}    On Error Resume Next\n    x = 1 / 0`, '11'],
		['Err.Clear first', `${RAISE}    Err.Clear\n    On Error Resume Next\n    x = 1 / 0`, '11'],
		['a conversion', `${RAISE}    On Error Resume Next\n    i = CInt("abc")`, '13'],
		['On Error GoTo 0 first', `${RAISE}    On Error GoTo 0\n    On Error Resume Next\n    x = 1 / 0`, '11'],
		['an Exit inside an If first', `${RAISE}    If Err.Number = 6 Then Exit Function\n    On Error Resume Next\n    x = 1 / 0`, '11'],
		['a second On Error GoTo', `${RAISE}    On Error GoTo Second\n    x = 1 / 0\n    Exit Function\nSecond:\n    Main = "second"`, '11'],
		['a line-number handler', '    On Error GoTo 100\n    Err.Raise 5\n    Exit Function\n100 On Error Resume Next\n    x = 1 / 0', '11'],
		['a handler below an End If', '    On Error GoTo Handler\n    If x = 0 Then\n        Err.Raise 5\n    End If\n    Exit Function\nHandler:\n    On Error Resume Next\n    x = 1 / 0', '11'],
		['a label nothing names in the handler', `${RAISE}Note:\n    On Error Resume Next\n    x = 1 / 0`, '11'],
		['numbered lines in the handler', '10  On Error GoTo 100\n20  Err.Raise 5\n30  Exit Function\n100 Main = "h"\n110 On Error Resume Next\n120 x = 1 / 0', '11'],
		['On Local Error in the handler', '    On Local Error GoTo Handler\n    Err.Raise 5\n    Exit Function\nHandler:\n    On Local Error Resume Next\n    x = 1 / 0', '11'],
	];
	it.each(RAISES)('reports what %s raises', (_name, body, number) => {
		expect(runtimeErrors(body)).toEqual([number]);
	});

	const SURVIVES: ReadonlyArray<readonly [string, string]> = [
		['On Error GoTo -1 ends the handler', `${RAISE}    On Error GoTo -1\n    On Error Resume Next\n    x = 1 / 0`],
		['Resume leaves it', `${RAISE}    Resume Retry\nRetry:\n    On Error Resume Next\n    x = 1 / 0`],
		['no handler runs', '    On Error Resume Next\n    x = 1 / 0'],
		['the label is entered by falling into it', '    On Error GoTo Handler\n    Main = "a"\nHandler:\n    On Error Resume Next\n    x = 1 / 0'],
		['Resume Next returns from the handler', '    On Error GoTo Handler\n    Err.Raise 5\n    On Error Resume Next\n    x = 1 / 0\n    Exit Function\nHandler:\n    Resume Next'],
		['a line-numbered Resume Next', '10  On Error Resume Next\n20  x = 1 / 0'],
		['a labelled Resume Next', 'Start: On Error Resume Next\n    x = 1 / 0'],
		['On Local Error Resume Next', '    On Local Error Resume Next\n    x = 1 / 0'],
		// Labels a GoTo names are entered with no error pending.
		['a GoTo also enters the handler', '    On Error GoTo Handler\n    If x = 0 Then GoTo Handler\n    Err.Raise 5\n    Exit Function\nHandler:\n    On Error Resume Next\n    x = 1 / 0'],
		['a GoTo enters a label in the handler', '    On Error GoTo Handler\n    If x = 0 Then GoTo Done\n    Err.Raise 5\n    Exit Function\nHandler:\n    Main = "h"\nDone:\n    On Error Resume Next\n    x = 1 / 0'],
	];
	it.each(SURVIVES)('stays quiet where %s', (_name, body) => {
		expect(runtimeErrors(body)).toEqual([]);
	});
});
