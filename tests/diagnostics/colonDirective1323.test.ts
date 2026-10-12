// A colon after a compiler directive (issue #1323). Measured in Excel 16.0 on
// 2026-10-11: an #If ... Then, #ElseIf ... Then, #Else, #End If or #Const line
// a statement colon follows, with nothing, a comment or code after it, is no
// directive at all. Where it is compiled it is refused; in a skipped arm it is
// skipped with the arm; either way it opens, divides, closes and defines
// nothing.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

// No project constants, so an undefined name in #If is 0, as in Excel.
const errors = (source: string): string[] => analyzeVbaModuleSource({
	source, moduleName: 'Module1', knownIdentifiers: new Set<string>(), conditionalCompilation: { projectConstants: {} },
}).diagnostics.filter((d) => d.severity === 'error').map((d) => d.code);
const fn = (...lines: string[]): string => ['Function Main() As Variant', ...lines, 'End Function', ''].join('\n');
const ZERO = '    Main = 1 / 0';

describe('a colon after a compiler directive (issue #1323)', () => {
	it.each([
		fn('#If True Then:', '    Main = 1', '#End If'),
		fn('#If True Then :', '    Main = 1', '#End If'),
		fn("#If True Then: ' note", '    Main = 1', '#End If'),
		fn('#If True Then', '    Main = 1', '#End If:'),
		fn('#If True Then', '    Main = 1', "#End If: ' note"),
		fn('#If True Then', '    Main = 1', '#Else:', '    Main = 2', '#End If'),
		fn('#If True Then', '    Main = 1', '#Else: Main = 2', '#End If'),
		fn('#If False Then', "#If True Then: ' note", '#End If', '#End If', '    Main = 1'),
		'#Const A = 1:\n' + fn('    Main = 1'),
		"#Const A = 1: ' note\n" + fn('    Main = 1'),
		'#Const A = 1: #Const B = 2\n' + fn('    Main = 1'),
	])('reports what Excel refuses: %s', (source) => {
		expect(errors(source)).not.toEqual([]);
	});

	it.each([
		fn('#If False Then', '#If True Then', '#Else: Main = 2', '#End If', '#End If', '    Main = 1'),
		'#If False Then\n#Const A = 1: x = 1\n#End If\n' + fn('    Main = 1'),
		'#If False Then\n#Const A = 1: Rem note\n#End If\n' + fn('    Main = 1'),
		fn('    Main = 1', '#If #1/1/2000 12:00:00 AM# < 1 Then', '    Main = 2', '#End If'),
	])('stays quiet where Excel compiles: %s', (source) => {
		expect(errors(source)).toEqual([]);
	});

	it('takes no new arm at an #Else or #ElseIf a colon follows', () => {
		for (const line of ['#Else:', "#Else: ' note", '#Else: Rem note', "#ElseIf True Then: ' note", '#ElseIf True Then:']) {
			const source = fn('    Main = 1', '#If False Then', '    Main = 2', line, ZERO, '#End If');
			expect(errors(source), line).toEqual([]);
		}
		// Without the colon the arm is taken, and its division by zero found.
		expect(errors(fn('    Main = 1', '#If False Then', '    Main = 2', '#Else', ZERO, '#End If'))).toContain('division-by-zero');
	});

	it('defines nothing with a #Const a colon follows, even in a skipped arm', () => {
		const later = fn('    Main = 1', '#If A = 1 Then', ZERO, '#End If');
		expect(errors('#If False Then\n#Const A = 1: Rem note\n#End If\n' + later)).toEqual([]);
		expect(errors('#If False Then\n#Const A = 1: x = 1\n#End If\n' + later)).toEqual([]);
		// A plain #Const in a skipped arm still defines A (issue #192).
		expect(errors('#If False Then\n#Const A = 1\n#End If\n' + later)).toContain('division-by-zero');
	});
});
