// Diagnostics tests: a standard module's member used against its kind,
// through the module's name or bare (issue #423). Each case was compiled
// through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

function errors(body: string, module2: string, class1?: string): string[] {
	const main = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	const modules = [
		{ moduleName: 'Module1', source: main },
		{ moduleName: 'Module2', source: `Option Explicit\n${module2}\n` },
		...(class1 ? [{ moduleName: 'Class1', source: `Option Explicit\n${class1}\n`, moduleKind: 'class' as const }] : []),
	];
	return analyzeProjectModule(main, modules, 'Module1')
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code);
}

const SUB = 'Public Sub M()\nEnd Sub';
const GET = 'Public Property Get M() As Long\n    M = 1\nEnd Property';

describe("a standard module's member used against its kind (issue #423)", () => {
	it('reports the compile errors the VBE gives', () => {
		expect(errors('Module2.M = 9', SUB)).toEqual(['assignment-to-procedure-name']);
		expect(errors('Main = Module2.M', SUB)).toEqual(['sub-used-as-value']);
		expect(errors('Module2.M = 9', 'Public Function M() As Long\nEnd Function')).toEqual(['assignment-to-procedure-name']);
		expect(errors('Module2.M = 9', 'Public Declare PtrSafe Function M Lib "kernel32" Alias "GetTickCount" () As Long')).toEqual(['assignment-to-procedure-name']);
		expect(errors('Module2.M', 'Public M As Long')).toEqual(['non-callable-call']);
		expect(errors('Call Module2.M', 'Public Const M As Long = 1')).toEqual(['non-callable-call']);
		expect(errors('Module2.M', 'Public Enum E\n    M = 1\nEnd Enum')).toEqual(['non-callable-call']);
		expect(errors('Module2.M', GET)).toEqual(['invalid-property-use']);
		expect(errors('Call M', GET)).toEqual(['invalid-property-use']);
		expect(errors('M = 9', GET)).toEqual(['readonly-member-assignment']);
		expect(errors('Module2.M 5', 'Public Property Let M(ByVal v As Long)\nEnd Property')).toEqual(['invalid-property-use']);
		expect(errors('M = 9', 'Public Type M\n    x As Long\nEnd Type')).toEqual(['undeclared-variable']);
		expect(errors('Main = M', 'Public Type M\n    x As Long\nEnd Type')).toEqual(['undeclared-variable']);
		expect(errors('Dim c As New Class1\n    c.M = 9', 'Public Sub Other()\nEnd Sub', 'Public Function M() As Long\nEnd Function')).toEqual(['assignment-to-procedure-name']);
	});

	it('stays quiet on what compiles', () => {
		expect(errors('Module2.M', SUB)).toEqual([]);
		expect(errors('Call Module2.M', SUB)).toEqual([]);
		expect(errors('Main = Module2.M', GET)).toEqual([]);
		expect(errors('Main = M', GET)).toEqual([]);
		expect(errors('Debug.Print Module2.M', GET)).toEqual([]);
		expect(errors('Module2.M = 9', 'Public Function M() As Variant\nEnd Function')).toEqual([]);
		expect(errors('Module2.M = 9', 'Public M As Long')).toEqual([]);
		expect(errors('Dim Module2 As New Class1\n    Module2.M = 9', SUB, 'Public M As Long')).toEqual([]);
		// It compiles, and raises 424 when it runs: M returns Empty (issue #414).
		expect(errors('Dim c As New Class1\n    c.M = 9', 'Public Sub Other()\nEnd Sub', 'Public Function M() As Variant\nEnd Function')).toEqual(['variant-value-misuse']);
	});
});
