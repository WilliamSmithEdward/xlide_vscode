// github.com/WilliamSmithEdward/xlide_vscode/issues/1347.
//
// A module #Const of a built-in replaces its value even when that value is
// not known here: `#Const Win64 = UNDEFINED_X` then `#If Win64` takes the
// #Else arm in Excel 16.0 (no project arguments set), where XLIDE kept the
// built-in Win64 = 1 and took the #If arm. With the project arguments
// omitted, an undefined name may be one of them (#102), so both arms are
// unknown; with them supplied, the name is 0 and the arm is decided.

import { describe, expect, it } from 'vitest';
import { createConditionalActivityTracker, parseModule } from '../src/analyzer';
import { documentOutlineSymbolsForSource } from '../src/vbaSymbolPresentation';

const module = (name: string, value: string): string =>
	`#Const ${name} = ${value}\nSub Main()\n#If ${name} Then\n    Debug.Print 1\n#Else\n    Debug.Print 2\n#End If\nEnd Sub\n`;

/** The activity of the #If arm and the #Else arm. */
function arms(source: string, env?: Parameters<typeof createConditionalActivityTracker>[1]): string[] {
	const tracker = createConditionalActivityTracker(parseModule(source), env)!;
	return ['Debug.Print 1', 'Debug.Print 2'].map((text) => {
		const at = source.indexOf(text);
		return tracker.activityForSpan({ start: at, end: at + 1 });
	});
}

describe('a #Const of a built-in whose value is not known (#1347)', () => {
	it.each(['Win64', 'VBA7'])('leaves both arms unknown for %s when the project arguments are not known', (name) => {
		expect(arms(module(name, 'UNDEFINED_X'))).toEqual(['unknown', 'unknown']);
		expect(arms(module(name, 'MISSING + 1'))).toEqual(['unknown', 'unknown']);
	});

	it('decides the arm Excel compiles once the arguments are known', () => {
		// Measured: issue1347_win64_undefined_compile and
		// issue1347_win64_undefined_plus_one_compile.
		expect(arms(module('Win64', 'UNDEFINED_X'), { projectConstants: {} })).toEqual(['inactive', 'active']);
		expect(arms(module('Win64', 'UNDEFINED_X + 1'), { projectConstants: {} })).toEqual(['active', 'inactive']);
		expect(arms(module('Win64', 'MISSING + 1'), { projectConstants: { MISSING: -1 } })).toEqual(['inactive', 'active']);
	});

	it('still replaces a built-in with a value it knows', () => {
		// Measured: issue1347_win64_zero_compile and issue1347_vba7_zero_compile.
		expect(arms(module('Win64', '0'))).toEqual(['inactive', 'active']);
		expect(arms(module('VBA7', '0'))).toEqual(['inactive', 'active']);
	});

	it('keeps an unknown #Const unknown with the arguments supplied, rather than reading it as 0', () => {
		// 1 / 0 has no value; Win64 is neither its built-in 1 nor an undefined 0.
		expect(arms(module('Win64', '1 / 0'), { projectConstants: {} })).toEqual(['unknown', 'unknown']);
	});

	it('lists the #Else arm\'s procedure in the Outline, which reads no project arguments', () => {
		const source = '#Const Win64 = UNDEFINED_X\n#If Win64 Then\nSub A()\nEnd Sub\n#Else\nSub B()\nEnd Sub\n#End If\n';
		expect(documentOutlineSymbolsForSource('Module1', 'standard', source).map((s) => s.name)).toContain('B');
	});
});
