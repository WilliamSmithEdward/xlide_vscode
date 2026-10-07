import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string, extra = '', declarations = ''): string[] {
	const source = `Option Explicit\n${declarations}\nSub Main()\n${body}\nEnd Sub\n${extra}`;
	return analyzeModule(source).filter(d => d.severity === 'error').map(d => d.code);
}

describe('indirect changes to runtime state', () => {
	it.each([
		'If Fill(ws) Then\nDebug.Print ws.Range("A1").Find("x").Address\nEnd If',
		'Do While Fill(ws)\nDebug.Print ws.Range("A1").Find("x").Address\nExit Do\nLoop',
	])('does not carry sheet emptiness through a block header: %s', operation => {
		expect(errors(`Dim ws As Worksheet\nSet ws = Worksheets.Add\n${operation}`, 'Function Fill(ws As Worksheet) As Boolean\nws.Range("A1").Value = "x"\nFill = True\nEnd Function')).toEqual([]);
	});

	it('does not infer a shared RegExp pattern across helper calls', () => {
		expect(errors('Set rx = CreateObject("VBScript.RegExp")\nrx.Pattern = "x"\nChangePattern\nDebug.Print rx.Execute("y")(0).Value', 'Sub ChangePattern()\nrx.Pattern = "y"\nEnd Sub', 'Dim rx As Object')).toEqual([]);
	});

	it('does not carry file absence between filesystem instances', () => {
		expect(errors('Dim fso As Object, other As Object, ts As Object\nSet fso = CreateObject("Scripting.FileSystemObject")\nSet other = CreateObject("Scripting.FileSystemObject")\nSet ts = fso.CreateTextFile("xlide-audit.tmp")\nts.Close\nfso.DeleteFile "xlide-audit.tmp"\nSet ts = other.CreateTextFile("xlide-audit.tmp")\nts.Close\nDebug.Print fso.GetFile("xlide-audit.tmp").Size')).toEqual([]);
	});

	it('does not retain file absence across a helper using only the path', () => {
		expect(errors('Dim fso As Object, ts As Object\nSet fso = CreateObject("Scripting.FileSystemObject")\nSet ts = fso.CreateTextFile("xlide-audit.tmp")\nts.Close\nfso.DeleteFile "xlide-audit.tmp"\nRestoreFile\nDebug.Print fso.GetFile("xlide-audit.tmp").Size', 'Sub RestoreFile()\nDim writer As Object\nSet writer = CreateObject("Scripting.FileSystemObject")\nwriter.CreateTextFile("xlide-audit.tmp").Close\nEnd Sub')).toEqual([]);
	});

	it.each(['ok = Restore()', 'ok = RestoreValue'])('does not retain deleted settings across an assignment: %s', operation => {
		expect(errors(`Dim ok As Boolean\nSaveSetting "xlideAudit", "s", "k", "v"\nDeleteSetting "xlideAudit", "s", "k"\n${operation}\nDeleteSetting "xlideAudit", "s", "k"`, 'Function Restore() As Boolean\nSaveSetting "xlideAudit", "s", "k", "v"\nRestore = True\nEnd Function\nProperty Get RestoreValue() As Boolean\nSaveSetting "xlideAudit", "s", "k", "v"\nRestoreValue = True\nEnd Property')).toEqual([]);
	});

	it('retains definite local pattern and consecutive setting checks', () => {
		expect(errors('Dim rx As Object\nSet rx = CreateObject("VBScript.RegExp")\nrx.Pattern = "x"\nDebug.Print rx.Execute("y")(0).Value')).toEqual(['collection-index-out-of-range']);
		expect(errors('DeleteSetting "xlideAudit", "s", "k"\nDeleteSetting "xlideAudit", "s", "k"')).toEqual(['runtime-argument-value']);
	});
});
