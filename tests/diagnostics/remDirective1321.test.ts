// `Rem` on a compiler-directive line (issue #1321). The VBE does not read Rem
// as a comment there. Measured in 64-bit Excel 16.0 (2026-10-12): an #If,
// #ElseIf, #Else or #End If line holding a Rem is no directive at all, so it
// opens, continues and closes nothing, even in a skipped arm, and in a
// compiled arm it is an error; a #Const line holding one is an error where it
// is compiled.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeVbaStructure } from '../../src/vbaStructuralDiagnostics';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { stripVba } from '../../src/vbaSourceScan';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const OPTIONS = { knownIdentifiers: new Set<string>(), conditionalCompilation: { projectConstants: {} } };
const analyze = (src: string) => analyzeModule(src, OPTIONS);
const malformed = (src: string) => byCode(analyze(src), 'malformed-statement');
const structure = (src: string) => analyzeVbaStructure(src).map((d) => d.message);
const undeclared = (src: string) =>
	byCode(analyze(src), 'undeclared-variable').map((d) => /'(\w+)'/.exec(d.message)?.[1]);

const fn = (...lines: string[]) => ['Option Explicit', 'Function Main() As Variant', ...lines, '    Main = 1', 'End Function', ''].join('\n');

describe('an #If, #ElseIf, #Else or #End If line holding a Rem (issue #1321)', () => {
	it.each([
		['#If Win64 Then Rem note', 'opens no #If block'],
		['#If True Rem note Then', 'opens no #If block'],
		['#If 1Rem note Then', 'opens no #If block'],
		['#If True _\n    Rem note Then', 'opens no #If block'],
		['#If True Then REM note', 'opens no #If block'],
	])('reports %s, and the #Else and #End If after it have no #If', (line, effect) => {
		const src = fn(line, '    a = 1', '#Else', '    b = 2', '#End If');
		expectDiagnostic(src, analyze(src), 'malformed-statement', { message: ['Rem does not start a comment', effect, "with ' instead"] });
		expect(structure(src)).toEqual(["'#Else' has no matching '#If'.", "'#End If' has no matching '#If'."]);
	});

	it('reports a Rem #Else or #ElseIf in a compiled arm, which then runs on into the next line', () => {
		for (const line of ['#Else Rem note', '#ElseIf False Rem note Then', '#ElseIf False Then Rem note']) {
			const src = fn('#If True Then', '    a = 1', line, '    b = 2', '#End If');
			expectDiagnostic(src, analyze(src), 'malformed-statement', { span: line.replace(/(Rem).*/, '$1'), message: 'starts no new arm' });
			expect(undeclared(src)).toEqual(['a', 'b']);
			expect(structure(src)).toEqual([]);
		}
	});

	it('reports a Rem #End If, and leaves its #If open', () => {
		for (const line of ['#End If Rem note', '#EndIf Rem note']) {
			const src = fn('#If True Then', '    a = 1', line);
			expectDiagnostic(src, analyze(src), 'malformed-statement', { message: 'closes no #If block' });
			expect(structure(src)).toEqual(["Missing '#End If' for '#If'."]);
		}
	});

	it('stays quiet in a skipped arm, where the line is skipped with the arm', () => {
		for (const line of ['#Else Rem note', '#ElseIf True Rem note Then']) {
			const src = fn('#If False Then', '    a = 1', line, '    b = 2', '#End If');
			expect(malformed(src)).toHaveLength(0);
			expect(undeclared(src)).toEqual([]);
			expect(structure(src)).toEqual([]);
		}
	});

	it('still pairs the #If blocks around a Rem line in a skipped arm as the VBE does', () => {
		expect(structure(fn('#If False Then', '#If True Rem note Then', '#End If', '#End If'))).toEqual(["'#End If' has no matching '#If'."]);
		expect(structure(fn('#If False Then', '#If True Then', '#End If Rem note', '#End If'))).toEqual(["Missing '#End If' for '#If'."]);
		expect(malformed(fn('#If False Then', '#If True Rem note Then', '#End If', '#End If'))).toHaveLength(0);
	});

	it('reports a Rem line at module level too', () => {
		const src = ['#If Win64 Then Rem note', 'Private x As Long', '#End If', 'Function Main() As Variant', '    Main = 1', 'End Function', ''].join('\n');
		expectDiagnostic(src, analyze(src), 'malformed-statement', { message: 'opens no #If block' });
		expect(structure(src)).toEqual(["'#End If' has no matching '#If'."]);
	});

	it('finds the Rem lines under CRLF and bare CR line endings', () => {
		const lf = fn('#If False Then', '#If True Rem note Then', '#End If', '#End If');
		for (const eol of ['\r\n', '\r']) {
			expect(structure(lf.replace(/\n/g, eol))).toEqual(["'#End If' has no matching '#If'."]);
		}
	});

	it('counts only a directive that starts its line, so the parser and analyzeVbaStructure agree', () => {
		// After a colon, or after a continued line, the line is not read as a
		// Rem directive by either (directive-trailing-statement owns the colon).
		const rem = (src: string) => malformed(src).filter((d) => d.message.startsWith('Rem does not'));
		for (const src of [
			fn('#If True Then: #Else Rem note', '#End If'),
			fn('x = 1 _', ': #If True Rem note Then', '#End If'),
		]) {
			expect(rem(src)).toHaveLength(0);
			expect(structure(src)).toEqual([]);
		}
	});

	it('leaves Remark, a "Rem" string, a member named Rem, an apostrophe comment and Rem after the line alone', () => {
		const quiet = [
			fn('#If Remark Then', '#End If'),
			fn('#If "Rem" = "Rem" Then', '#End If'),
			fn("#If True Then ' Rem note", '#End If'),
			fn('#If True Then', '    Rem note', '#End If'),
			fn('#If True Then', '#Else', '#End If ' + "' Rem"),
			fn('#If x.Rem Then', '#End If'),
			fn('#If x!Rem Then', '#End If'),
		];
		for (const src of quiet) {
			expect(malformed(src)).toHaveLength(0);
			expect(structure(src)).toEqual([]);
		}
	});
});

describe('the text scanner reads Rem as the lexer does (issue #1321)', () => {
	it('takes Rem after a digit or _ as a comment, and leaves it inside a name', () => {
		expect(stripVba('#If 1Rem note Then')).toBe('#If 1             ');
		expect(stripVba('#If True _Rem note Then')).toBe('#If True _             ');
		for (const line of ['If [1Rem] Then', 'If [_Rem] Then', 'If x1Rem Then', 'If 1Remé Then', 'If ÉRem Then']) {
			expect(stripVba(line)).toBe(line);
			expect(structure(fn(line, 'End If'))).toEqual([]);
		}
	});

	it('does not take the _ left before a blanked Rem as a line continuation', () => {
		for (const line of ['    x = y _Rem note', '    x = y _ Rem note', '#If X _Rem note Then']) {
			expect(structure(['Sub A()', line, 'End Sub', ''].join('\n'))).toEqual([]);
		}
	});

	it('still joins real continuations, and runs a comment on through its ` _`', () => {
		const sub = (...lines: string[]) => ['Sub A()', ...lines, 'End Sub', ''].join('\n');
		for (const eol of ['\n', '\r\n', '\r']) {
			// A comment ending in ` _` swallows the next line, End If included.
			expect(structure(sub('    If x Then', "    ' note _", '    End If', '    End If').replace(/\n/g, eol))).toEqual([]);
			expect(structure(sub('    If x _', '        And y _', '        Then', '    End If').replace(/\n/g, eol))).toEqual([]);
			expect(structure(sub('    s = "a _"', '    If x Then', '    End If').replace(/\n/g, eol))).toEqual([]);
		}
	});
});

describe('a #Const line holding a Rem (issue #1321)', () => {
	const mod = (line: string) => [line, 'Function Main() As Variant', '    Main = 1', 'End Function', ''].join('\n');

	it.each([
		['#Const A = True Rem note', 'Expected: end of statement'],
		['#Const A = True _\n    Rem note', 'Expected: end of statement'],
		['#Const A = Rem note', 'Expected: expression'],
		['#Const Rem = 1', 'Expected: identifier'],
	])('reports %s: %s', (line, error) => {
		const src = mod(line);
		expectDiagnostic(src, analyze(src), 'malformed-statement', { message: ['#Const line', error] });
	});

	it('stays quiet in a skipped arm, and for an apostrophe comment or a "Rem" value', () => {
		expect(malformed(mod('#If False Then\n#Const A = True Rem note\n#End If'))).toHaveLength(0);
		expect(malformed(mod("#Const A = True ' Rem note"))).toHaveLength(0);
		expect(malformed(mod('#Const A = "Rem"'))).toHaveLength(0);
		expect(malformed(mod('#Const Remark = 1'))).toHaveLength(0);
	});
});

describe('a reserved word as a #Const name (issue #1321)', () => {
	// Measured: one name of each kind of reserved word is "Expected:
	// identifier" (the oracle holds all 56 names measured), and Error, Win64,
	// Object and [Rem] compile.
	const mod = (line: string) => [line, 'Function Main() As Variant', '    Main = 1', 'End Function', ''].join('\n');
	const names = (line: string) => byCode(analyze(mod(line)), 'malformed-statement');

	it('reports a reserved name, once even with a Rem on the line', () => {
		for (const name of ['Empty', 'True', 'If', 'Not', 'String', 'Debug', 'Me']) {
			const src = mod(`#Const ${name} = 1`);
			expectDiagnostic(src, analyze(src), 'malformed-statement', { span: name, message: ['reserved word', 'Expected: identifier'] });
		}
		const src = mod('#Const Empty = 1 Rem note');
		expectDiagnostic(src, analyze(src), 'malformed-statement', { span: 'Empty', message: 'Expected: identifier' });
	});

	it('leaves names the VBE takes alone, and a reserved name in a skipped arm', () => {
		for (const name of ['Error', 'Win64', 'Object', '[Rem]']) {
			expect(names(`#Const ${name} = 1`)).toHaveLength(0);
		}
		expect(names('#If False Then\n#Const Empty = 1\n#End If')).toHaveLength(0);
	});
});

describe('typing a Rem onto a directive line (issue #1321)', () => {
	type Result = ReturnType<typeof analyzeVbaModuleSource>;
	const run = (source: string, state?: Result['rulesIncrementalState']) => analyzeVbaModuleSource({
		source, moduleName: 'Module1', knownIdentifiers: new Set<string>(), rulesIncremental: { state, fingerprint: ['fp'] },
	});
	const key = (r: Result) => r.diagnostics.map((d) => `${d.code}:${d.span.start}:${d.span.end}:${d.message}`).sort().join('\n');

	it('gives the same diagnostics incrementally as a full pass, both ways', () => {
		const before = fn('#If True Then', '    a = 1', '#Else', '    b = 2', '#End If');
		for (const after of [before.replace('#Else', '#Else Rem note'), before.replace('#If True Then', '#If True Then Rem note')]) {
			const previous = run(before);
			const edited = run(after, previous.rulesIncrementalState);
			expect(key(edited)).toBe(key(run(after)));
			expect(edited.diagnostics.some((d) => d.code === 'malformed-statement' && d.message.startsWith('Rem does not'))).toBe(true);
			expect(previous.diagnostics.some((d) => d.code === 'malformed-statement')).toBe(false);
			expect(key(run(before, edited.rulesIncrementalState))).toBe(key(previous));
		}
	});
});
