// Diagnostics tests: statement words, library procedures and Application
// members named bare where a value goes (issue #318). Measured in Excel,
// Word and PowerPoint 16.0 64-bit (2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string, host?: 'word', explicit = true): string[] {
	const src = `${explicit ? 'Option Explicit\n' : ''}Function Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src, { knownIdentifiers: new Set<string>(), ...(host ? { host } : {}) }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a statement word read as a value', () => {
	it('is a Syntax error', () => {
		for (const word of ['Tab', 'Spc', 'Input', 'Print', 'Write', 'Shared', 'Len', 'Array']) {
			expect(found(`Main = TypeName(${word})`), word).toEqual([expect.stringMatching(/^reserved-keyword-in-expression: .*Syntax error/)]);
		}
		expect(found('Main = Tab(5)')).toEqual([expect.stringMatching(/Syntax error/)]);
	});

	it('is Argument not optional for Seek', () => {
		expect(found('Main = Seek')).toEqual([expect.stringMatching(/^argument-count: .*Argument not optional/)]);
	});

	it('runs in a Print list or with its argument list', () => {
		expect(found('Debug.Print "a"; Tab(5); "b"\n    Main = 1')).toEqual([]);
		expect(found('Debug.Print "a"; Spc(2); "b"\n    Main = 1')).toEqual([]);
		expect(found('Main = Len("ab") + Len("c")')).toEqual([]);
		expect(found('Main = Environ$("TEMP")')).toEqual([]);
	});
});

describe('a library procedure named bare where a value goes', () => {
	it('is Argument not optional when it needs an argument', () => {
		for (const word of ['Left', 'Mid', 'Kill', 'Width', 'ChDir', 'Shell', 'MsgBox', 'Format', 'Environ']) {
			expect(found(`Main = TypeName(${word})`), word).toEqual([expect.stringMatching(/^argument-count: .*Argument not optional/)]);
		}
	});

	it('is Expected Function or variable for a statement that takes none', () => {
		for (const word of ['Beep', 'Reset', 'Randomize']) {
			expect(found(`Main = TypeName(${word})`), word).toEqual([expect.stringMatching(/^sub-used-as-value: .*Expected Function or variable/)]);
		}
	});

	it('gives its value when every argument is optional', () => {
		for (const word of ['Now', 'Timer', 'DoEvents', 'CurDir', 'Rnd', 'Error', 'FreeFile', 'Date', 'Time', 'Erl']) {
			expect(found(`Main = TypeName(${word})`), word).toEqual([]);
		}
	});

	it('is the module\'s where the module declares the name', () => {
		expect(found('Dim Left As Long\n    Left = 2\n    Main = Left')).toEqual([]);
	});
});

describe('Line, Name and Application members that Global lacks', () => {
	it('are Variable not defined under Option Explicit', () => {
		for (const word of ['Line', 'Name', 'ScreenUpdating', 'Caption', 'Version']) {
			expect(found(`Main = ${word}`), word).toEqual([expect.stringMatching(/^undeclared-variable: /)]);
		}
		expect(found('Main = Name', undefined, false)).toEqual([]);
	});

	it('follow the host\'s own Global: Word has Name, not Caption', () => {
		expect(found('Main = Name', 'word')).toEqual([]);
		expect(found('Main = Caption', 'word')).toEqual([expect.stringMatching(/^undeclared-variable: /)]);
	});

	it('leave Global\'s members and the statements themselves alone', () => {
		expect(found('Main = ActiveSheet.Name')).toEqual([]);
		expect(found('Dim s As String\n    Open "x" For Input As #1\n    Line Input #1, s\n    Close #1\n    Main = s')).toEqual([]);
	});
});
