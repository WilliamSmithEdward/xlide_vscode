import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

/** True when the module reports the "expected end of statement" juxtaposition error. */
function flagsJuxtaposition(body: string): boolean {
	const src = `Sub T()\n    Dim n As Long\n    Dim arr As Variant\n    ${body}\nEnd Sub\n`;
	const r = analyzeVbaModuleSource({ source: src, moduleName: 'Module1' });
	return r.diagnostics.some((d) => /expected end of statement/i.test(d.message));
}

describe('juxtaposed-values diagnostic', () => {
	it.each([
		'n = 1 n 1',
		'n = 1 MsgBox("hello") 1',
		'n = 1 1',
		'n = Foo() bar',
		'n = "a" 1',
	])('flags juxtaposition: %s', (body) => {
		expect(flagsJuxtaposition(body)).toBe(true);
	});

	it.each([
		'n = 1 + 1',
		'n = Foo(1)',
		'n = MsgBox (1)',     // a call written with a space before the paren
		'n = arr(1)(2)',      // jagged-array / call-chain access
		'n = Count&',         // Long type-declaration suffix
		'n = a.b',            // member access
		'n = "x" & "y"',
		'n = IIf(1, 2, 3)',
		'n = (1 = 2)',        // parenthesized comparison
		'n = New Collection', // New <Type>
		'MsgBox "hello"',     // implicit call statement (no '=')
		'Debug.Print n',      // call statement
		'Foo n, 1',           // call statement with args
	])('does not flag valid: %s', (body) => {
		expect(flagsJuxtaposition(body)).toBe(false);
	});
});
