// Diagnostics tests: procedures named in strings (issue #408). Measured in
// Excel 16.0 64-bit (build 20430, 2026-10-02) through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const MODULE_TAIL = 'Public Function Twice(ByVal n As Long) As Long\n    Twice = n * 2\nEnd Function\nPublic Sub NoArgs()\nEnd Sub\n'
	+ 'Public Function Opt(Optional ByVal n As Long) As Long\n    Opt = n\nEnd Function\n';
const CLASS = 'Option Explicit\nPublic Fld As Long\nPublic Property Get P() As Long\n    P = 1\nEnd Property\n'
	+ 'Public Function F(ByVal x As Long) As Long\n    F = x\nEnd Function\nPublic Sub S()\nEnd Sub\n'
	+ 'Public Property Get Q() As Long\n    Q = 1\nEnd Property\nPublic Property Let Q(ByVal v As Long)\nEnd Property\n';

function errors(line: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1, k As New Collection\n    k.Add 5\n    ${line}\nEnd Function\n${MODULE_TAIL}`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Class1', source: CLASS, type: 'class' },
	], 'Module1', { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a procedure called by name', () => {
	it('raises for a name, a count or a call type that does not fit', () => {
		for (const [line, error] of [
			['Run "Nope"', '1004'], ['Application.Run "NoArgs", 1', '450'], ['Main = Application.Run("Twice")', '449'],
			['Main = Application.Run("Opt", 1, 2)', '450'], ['CallByName c, "P", VbLet, 5', '451'],
			['Main = CallByName(c, "F", VbMethod)', '449'], ['Main = CallByName(c, "F", VbMethod, 1, 2)', '450'],
			['Main = CallByName(k, "Count", VbGet)', '438'], ['Main = CallByName(k, "Item", VbGet, 1)', '438'], ['CallByName k, "Add", VbLet, 1', '438'],
		]) {
			expect(errors(line), line).toEqual([expect.stringMatching(new RegExp(`^runtime-member-not-found: .*'${error}'`))]);
		}
	});

	it('is quiet where it runs', () => {
		for (const line of [
			'Main = Application.Run("Twice", 3)', 'Main = Application.Run("Module1.Twice", 3)', 'Main = Application.Run("twice", 3)',
			'Main = Run("Twice", 2)', 'Run "NoArgs"', 'Main = Application.Run("Opt")', 'CallByName c, "S", VbMethod',
			'Main = CallByName(c, "P", VbGet)', 'Main = CallByName(c, "Fld", VbGet)', 'CallByName c, "Q", VbLet, 5', 'CallByName c, "Fld", VbLet, 5',
			'Main = CallByName(k, "Count", VbMethod)', 'Main = CallByName(k, "Item", VbMethod, 1)',
		]) {
			expect(errors(line), line).toEqual([]);
		}
	});
});
