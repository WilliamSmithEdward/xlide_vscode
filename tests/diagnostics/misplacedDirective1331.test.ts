// A compiler directive that does not start its logical line (issue #1331).
// Measured in Excel 16.0 on 2026-10-12: a `#` is a directive only where it
// starts its logical line, after nothing but blanks and line continuations.
// After a statement's or a label's colon, another directive's colon, or a
// line number it is none: where it is compiled it is refused, and either way
// it opens, divides, closes and defines nothing, even in a skipped arm.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

// No project constants, so an undefined name in #If is 0, as in Excel.
const errors = (source: string): string[] => analyzeVbaModuleSource({
	source, moduleName: 'Module1', knownIdentifiers: new Set<string>(), conditionalCompilation: { projectConstants: {} },
}).diagnostics.filter((d) => d.severity === 'error').map((d) => d.code);
const fn = (...lines: string[]): string => ['Function Main() As Variant', ...lines, 'End Function', ''].join('\n');
const ZERO = '    Main = 1 / 0';

describe('a compiler directive that does not start its logical line (issue #1331)', () => {
	it.each([
		['after a statement', fn('    x = 1: #If True Then', '    Main = 1', '#End If')],
		['across a continuation', fn('    x = 1: _', '#If True Then', '    Main = 1', '#End If')],
		['after a label', fn('L1: #If True Then', '    Main = 1', '#End If')],
		['after a line number', fn('10 #If True Then', '    Main = 1', '#End If')],
		['#End If after a statement', fn('#If True Then', '    Main = 1', '    x = 1: #End If')],
		['#End If after a line number', fn('#If True Then', '    Main = 1', '10 #End If')],
		['#Else after a statement', fn('#If True Then', '    Main = 1', '    x = 1: #Else', '    Main = 2', '#End If')],
		['#Else after a line number', fn('#If True Then', '    Main = 1', '10 #Else', '    Main = 2', '#End If')],
		['#Const after a statement', fn('    x = 1: #Const A = 1', '    Main = 1')],
		['#Const after a line number', fn('10 #Const A = 1', '    Main = 1')],
		['#If after a directive', '#Const B = 1: #If False Then\n#End If\n' + fn('    Main = 1')],
		['#If in a skipped arm', fn('    Main = 1', '#If False Then', '    x = 1: #If True Then', '#End If', '#End If')],
		['#If after a line number in a skipped arm', fn('    Main = 1', '#If False Then', '10 #If True Then', '#End If', '#End If')],
		['#End If in a skipped arm', fn('    Main = 1', '#If False Then', '#If True Then', '    x = 1: #End If', '#End If')],
	])('reports what Excel refuses: %s', (_what, source) => {
		expect(errors(source)).not.toEqual([]);
	});

	it.each([
		['indented', fn('    Main = 1', '    #If True Then', '    Main = 2', '    #End If')],
		['tab-indented', fn('    Main = 1', '\t#If True Then', '    Main = 2', '\t#End If')],
		['after a continuation with nothing before it', fn('    Main = 1', ' _', '#If True Then', '    Main = 2', '#End If')],
		['#Else after a statement, in a skipped arm', fn('    Main = 1', '#If False Then', '    x = 1: #Else', ZERO, '#End If')],
	])('stays quiet where Excel compiles: %s', (_what, source) => {
		expect(errors(source)).toEqual([]);
	});

	it('defines nothing with a #Const that does not start its line, even in a skipped arm', () => {
		const later = fn('    Main = 1', '#If A = 1 Then', ZERO, '#End If');
		for (const line of ['Private y As Long: #Const A = 1', 'Private y As Long: _\n#Const A = 1', 'L1: #Const A = 1', '10 #Const A = 1', '#Const B = 1: #Const A = 1', '#Else: #Const A = 1']) {
			expect(errors(`#If False Then\n${line}\n#End If\n${later}`), line).toEqual([]);
		}
		// A plain #Const in a skipped arm still defines A (issue #192).
		expect(errors(`#If False Then\n#Const A = 1\n#End If\n${later}`)).toContain('division-by-zero');
	});

	it('reports the line once, and once only after a directive\'s colon', () => {
		const once = (source: string) => analyzeVbaModuleSource({ source, moduleName: 'Module1', knownIdentifiers: new Set<string>() })
			.diagnostics.filter((d) => d.code === 'malformed-statement' || d.code === 'directive-trailing-statement');
		expect(once(fn('    x = 1: #Const A = 1', '    Main = 1'))).toHaveLength(1);
		expect(once('#Const B = 1: #Const A = 1\n' + fn('    Main = 1'))).toHaveLength(1);
	});
});
