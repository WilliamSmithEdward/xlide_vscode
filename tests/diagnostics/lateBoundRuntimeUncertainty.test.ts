import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeProjectModule } from './helpers';

const SETUP = 'Dim fso As Object, ts As Object\nSet fso = CreateObject("Scripting.FileSystemObject")\nSet ts = fso.CreateTextFile("xlide-audit.tmp")\nts.Close\nfso.DeleteFile "xlide-audit.tmp"\n';

describe('late-bound state across project calls and error handlers', () => {
	it.each([
		['RestoreFile', 'Public Sub RestoreFile()\n', '\nEnd Sub'],
		['Helpers.RestoreFile', 'Public Sub RestoreFile()\n', '\nEnd Sub'],
		['ok = RestoreFile()', 'Public Function RestoreFile() As Boolean\n', '\nRestoreFile = True\nEnd Function'],
	])('forgets deleted files across a cross-module call: %s', (call, header, footer) => {
		const caller = `Option Explicit\nSub Main()\nDim ok As Boolean\n${SETUP}${call}\nDebug.Print fso.GetFile("xlide-audit.tmp").Size\nEnd Sub`;
		const helper = `Option Explicit\n${header}Dim writer As Object\nSet writer = CreateObject("Scripting.FileSystemObject")\nwriter.CreateTextFile("xlide-audit.tmp").Close${footer}`;
		const findings = analyzeProjectModule(caller, [{ moduleName: 'Helpers', source: helper }], 'Caller');
		expect(findings.filter(d => d.severity === 'error')).toEqual([]);
	});

	it('does not assume a failed CreateTextFile replaced a readable stream', () => {
		const source = 'Option Explicit\nSub Main()\nDim fso As Object, ts As Object\nSet fso = CreateObject("Scripting.FileSystemObject")\nSet ts = fso.CreateTextFile("xlide-audit.tmp")\nts.WriteLine "x"\nts.Close\nSet ts = fso.OpenTextFile("xlide-audit.tmp", 1)\nOn Error Resume Next\nSet ts = fso.CreateTextFile("xlide-audit.tmp", False)\nOn Error GoTo 0\nDebug.Print ts.ReadAll\nEnd Sub';
		expect(analyzeModule(source).filter(d => d.severity === 'error')).toEqual([]);
	});

	it('retains mode checks outside error-handled procedures', () => {
		const source = 'Option Explicit\nSub Main()\nDim fso As Object, ts As Object\nSet fso = CreateObject("Scripting.FileSystemObject")\nSet ts = fso.CreateTextFile("xlide-audit.tmp")\nDebug.Print ts.ReadAll\nEnd Sub';
		expect(analyzeModule(source).filter(d => d.severity === 'error').map(d => d.code)).toEqual(['file-mode-mismatch']);
	});
});
