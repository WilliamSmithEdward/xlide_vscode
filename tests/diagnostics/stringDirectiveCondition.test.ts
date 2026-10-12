// A string as a compiler-directive condition. Measured in 64-bit Excel 16.0
// (Office Home 2024, regional format en-AU, decimal point ".", 2026-10-12),
// with the VBE oracle in run and compile mode:
//
//  - The VBE reads the string as VBA converts a string to a Boolean (#188):
//    #If "0", "00", 0 & 0, "False", "0.0" and #If S after #Const S = "0" take
//    the #Else arm; "1", "2", "-1", "True", "true", "&H10", "1e3", "2.5" and
//    "1,000" the #If arm. #ElseIf "0" after a False arm and a nested
//    #If "0" take the #Else arm.
//  - "", "abc", " True " and #If S after #Const S = "x" are "Type mismatch",
//    and so is a string an operator cannot convert: Not "abc", "abc" = 1,
//    "abc" = #1/1/2020#, Null - "abc", "abc" * Null,
//    "a" & "b" as the whole condition, UNDEFINED_X - "abc". The VBE refuses
//    these in an arm it skips too, and in an #ElseIf after a taken arm.
//  - #Const A = "abc" + 1 is "Type mismatch" even when nothing reads it, and
//    in a skipped arm; #Const S = "x" alone compiles.
//  - Beside a date a string is read as a date: "1/1/2020" = #1/1/2020# is
//    True. An operator reads "&H10" as 16: "&H10" And 1 takes the #Else arm.

import { describe, it, expect, vi } from 'vitest';
import * as lexer from '../../src/analyzer/lexer/tokenize';
import { invalidConditionDirectives } from '../../src/analyzer/conditional/conditionalCompilation';
import { analyzeModule, createConditionalActivityTracker, parseModule } from '../../src/analyzer';
import { ProjectIndex } from '../../src/analyzer/symbols/projectIndex';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule } from '../../src/vbaProjectAnalysis';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const OPTIONS = { knownIdentifiers: new Set<string>(), conditionalCompilation: { projectConstants: {} } };
const mismatches = (src: string) => byCode(analyzeModule(src, OPTIONS), 'directive-type-mismatch');

/** The #If arm and the #Else arm of a two-arm module. */
const arms = (condition: string, setup = '') =>
	`${setup}Function Main() As Variant\n#If ${condition} Then\n    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n`;

function activities(src: string): string[] {
	const tracker = createConditionalActivityTracker(parseModule(src), OPTIONS.conditionalCompilation);
	expect(tracker).toBeDefined();
	return ['Main = 1', 'Main = 2'].map((text) => {
		const at = src.indexOf(text);
		expect(at).toBeGreaterThanOrEqual(0);
		return tracker!.activityForSpan({ start: at, end: at + 1 });
	});
}

describe('the arm a string condition takes', () => {
	it.each([
		['"0"', ''], ['"00"', ''], ['0 & 0', ''], ['"False"', ''], ['"0.0"', ''],
		['S', '#Const S = "0"\n'], ['" 1 " - 1', ''], ['"&H10" And 1', ''],
	])('#If %s %s takes the #Else arm', (condition, setup) => {
		expect(activities(arms(condition, setup))).toEqual(['inactive', 'active']);
	});

	it.each(['"1"', '"2"', '"-1"', '"True"', '"true"', '"&H10"', '"1e3"', '"2.5"', '"1,000"', 'Not "0"', '"1" And 1'])(
		'#If %s takes the #If arm',
		(condition) => {
			expect(activities(arms(condition))).toEqual(['active', 'inactive']);
		},
	);

	it('reads #ElseIf "0" and a nested #If "0" the same way', () => {
		const elseIf = 'Function Main() As Variant\n#If False Then\n    Main = 0\n#ElseIf "0" Then\n    Main = 1\n#Else\n    Main = 2\n#End If\nEnd Function\n';
		expect(activities(elseIf)).toEqual(['inactive', 'active']);
		const nested = 'Function Main() As Variant\n#If True Then\n#If "0" Then\n    Main = 1\n#Else\n    Main = 2\n#End If\n#End If\nEnd Function\n';
		expect(activities(nested)).toEqual(['inactive', 'active']);
	});

	it('analyzes the arm the VBE compiles, and not the other', () => {
		const src = 'Option Explicit\nSub S()\n#If "0" Then\n    skipped = 1\n#Else\n    compiled = 1\n#End If\nEnd Sub\n';
		const names = byCode(analyzeModule(src, OPTIONS), 'undeclared-variable').map((d) => src.slice(d.span.start, d.span.end));
		expect(names).toEqual(['compiled']);
	});

	it('leaves the arm unknown for a string no locale converts, and for a number it cannot value', () => {
		// "" and "abc" are measured errors; the rest are numbers whose
		// value depends on what XLIDE cannot see (a currency symbol, an
		// exponent past a Double, two decimal points in some locales), so
		// neither arm is chosen.
		for (const condition of ['""', '"abc"', '" True "', '"$5"', '"1e400"', '"1.5.5"', '"&H100000000"', '"&H100000000" And 1']) {
			expect(activities(arms(condition)), condition).toEqual(['unknown', 'unknown']);
		}
	});
});

describe('directive-type-mismatch', () => {
	const fn = (...lines: string[]) => ['Function Main() As Variant', ...lines, '    Main = 1', 'End Function', ''].join('\n');

	it.each([
		['#If "abc" Then', '"abc"'],
		['#If "" Then', '""'],
		['#If " True " Then', '" True "'],
		['#If Not "abc" Then', '"abc"'],
		['#If "abc" = 1 Then', '"abc"'],
		['#If "a" & "b" Then', '"ab"'],
		['#If UNDEFINED_X - "abc" Then', '"abc"'],
		['#If "abc" = #1/1/2020# Then', '"abc"'],
		['#If #1/1/2020# + "abc" Then', '"abc"'],
		['#If Null - "abc" Then', '"abc"'],
		['#If "abc" * Null Then', '"abc"'],
	])('flags %s', (directive, shown) => {
		const src = fn(directive, '#End If');
		expectDiagnostic(src, analyzeModule(src, OPTIONS), 'directive-type-mismatch', {
			span: directive, severity: 'error', message: [shown, 'Type mismatch'],
		});
	});

	it('flags #If S after #Const S = "x", and not #Const S = "x" alone', () => {
		const read = `#Const S = "x"\n${fn('#If S Then', '#End If')}`;
		expectDiagnostic(read, analyzeModule(read, OPTIONS), 'directive-type-mismatch', { span: '#If S Then' });
		expect(mismatches(`#Const S = "x"\n${fn()}`)).toHaveLength(0);
	});

	it('flags a #Const value an operator cannot convert, even unread or in a skipped arm', () => {
		const unread = `#Const A = "abc" + 1\n${fn()}`;
		expectDiagnostic(unread, analyzeModule(unread, OPTIONS), 'directive-type-mismatch', {
			span: '#Const A = "abc" + 1', message: ['#Const value', 'Type mismatch'],
		});
		expect(mismatches(`#If False Then\n#Const A = "abc" + 1\n#End If\n${fn()}`)).toHaveLength(1);
	});

	it('flags a skipped arm and an #ElseIf after a taken arm, unlike Null', () => {
		for (const condition of ['"abc"', 'Not "abc"', '"a" & "b"', 'UNDEFINED_X - "abc"']) {
			const src = fn('#If False Then', `#If ${condition} Then`, '#End If', '#End If');
			expectDiagnostic(src, analyzeModule(src, OPTIONS), 'directive-type-mismatch', {
				span: `#If ${condition} Then`, message: 'Type mismatch',
			});
		}
		expect(mismatches(`#Const S = "x"\n${fn('#If False Then', '#If S Then', '#End If', '#End If')}`)).toHaveLength(1);
		expect(mismatches(fn('#If True Then', '#ElseIf "abc" Then', '#End If'))).toHaveLength(1);
		// Null stays limited to lines the VBE evaluates (issue #208).
		const nulls = byCode(analyzeModule(fn('#If False Then', '#If Null Then', '#End If', '#End If'), OPTIONS), 'null-directive-condition');
		expect(nulls).toHaveLength(0);
	});

	it('stays quiet for strings and operators the VBE converts', () => {
		for (const condition of ['"0"', '"1"', '"True"', '"false"', '"2.5"', '"1,000"', '"&H10" And 1', '" 1 " - 1', 'Not "0"', '"1" And 1', '"2" - 2']) {
			expect(mismatches(fn(`#If ${condition} Then`, '#End If')), condition).toHaveLength(0);
		}
	});

	it('reads a string beside a date as a date, and leaves Null its own diagnostic', () => {
		const date = arms('"1/1/2020" = #1/1/2020#');
		expect(mismatches(date)).toHaveLength(0);
		// Excel takes the #If arm; how a date string reads depends on the
		// locale, so XLIDE leaves the arm unknown rather than guess.
		expect(activities(date)).toEqual(['unknown', 'unknown']);
		// "abc" = Null and Null + "abc" are "Invalid use of Null" in Excel.
		for (const condition of ['"abc" = Null', 'Null + "abc"']) {
			const src = fn(`#If ${condition} Then`, '#End If');
			expect(mismatches(src), condition).toHaveLength(0);
			expect(byCode(analyzeModule(src, OPTIONS), 'null-directive-condition'), condition).toHaveLength(1);
		}
	});

	it('converts a string only beside a literal in a comparison, as a Variant otherwise', () => {
		// Excel 16.0: 1 = "True" and "abc" = True are a Type mismatch;
		// "True" = True is True; UNDEFINED_X = "abc", UNDEFINED_X = "" and
		// UNDEFINED_X = "0" compile, also in a skipped arm.
		const flagged: [string, string][] = [['1 = "True"', 'as a number'], ['"abc" = True', 'as True or False'],
			['"abc" = -1', 'as a number'], ['"abc" = #1/1/2020#', 'as a date'], ['(1) = "abc"', 'as a number'],
			['"abc" + (1)', 'as a number']];
		for (const [condition, as] of flagged) {
			const [diagnostic, ...rest] = mismatches(arms(condition));
			expect(rest, condition).toEqual([]);
			expect(diagnostic?.message, condition).toContain(as);
		}
		for (const condition of ['"True" = True', 'UNDEFINED_X = "abc"', 'UNDEFINED_X = ""', 'UNDEFINED_X = "0"']) {
			expect(mismatches(arms(condition)), condition).toHaveLength(0);
			expect(mismatches(fn('#If False Then', `#If ${condition} Then`, '#End If', '#End If')), condition).toHaveLength(0);
		}
		// A #Const holds a Variant too.
		expect(mismatches(arms('A = "abc"', '#Const A = 1\n'))).toHaveLength(0);
		// The arms: UNDEFINED_X = "0" is True, as in Excel. A string beside a
		// name or a Boolean is not evaluated, so "True" = True is unknown.
		expect(activities(arms('UNDEFINED_X = "0"'))).toEqual(['active', 'inactive']);
		expect(activities(arms('"True" = True'))).toEqual(['unknown', 'unknown']);
	});

	it('reports nothing for a string the host supplied, until a #Const gives the name a value', () => {
		// How the VBE reads a non-numeric project argument was not measured.
		const host = { knownIdentifiers: new Set<string>(), conditionalCompilation: { projectConstants: { FLAG: 'x' } } };
		const read = fn('#If FLAG Then', '#End If');
		expect(byCode(analyzeModule(read, host), 'directive-type-mismatch')).toHaveLength(0);
		const redefined = `#Const FLAG = "y"\n${read}`;
		expect(byCode(analyzeModule(redefined, host), 'directive-type-mismatch')).toHaveLength(1);
		// A #Const that gives no value leaves the host's string, in either table.
		const failed = `#Const FLAG = 1 / 0\n${read}`;
		expect(byCode(analyzeModule(failed, host), 'directive-type-mismatch')).toHaveLength(0);
		const compiler = { knownIdentifiers: new Set<string>(), conditionalCompilation: { compilerConstants: { FLAG: 'x' }, projectConstants: {} } };
		expect(byCode(analyzeModule(read, compiler), 'directive-type-mismatch')).toHaveLength(0);
		expect(byCode(analyzeModule(failed, compiler), 'directive-type-mismatch')).toHaveLength(0);
	});

	it('reads a skipped #If after a #Const holding a string in a skipped arm', () => {
		const src = fn('#If False Then', '#Const S = "x"', '#If S Then', '#End If', '#End If');
		expectDiagnostic(src, analyzeModule(src, OPTIONS), 'directive-type-mismatch', { span: '#If S Then', message: 'as True or False' });
	});

	it('stays quiet when concatenation depends on an unknown constant', () => {
		// Without the project's arguments, MISSING may be anything.
		const src = fn('#If (MISSING & "x") - 1 Then', '#End If');
		expect(byCode(analyzeModule(src), 'directive-type-mismatch')).toHaveLength(0);
	});

	it('keeps string operations and comparisons quiet', () => {
		for (const condition of ['VBA7 And "x" = "x"', '"x" = Empty', 'Empty = "x"', '"x" + "y" = "xy"', 'S & "y" = "xy"']) {
			expect(mismatches(arms(condition, '#Const S = "x"\n')), condition).toHaveLength(0);
		}
	});

	it('reads a string every arithmetic and logical operator converts', () => {
		// Implementation regressions for each place a string is converted;
		// the measured forms are above.
		for (const condition of ['"abc" Mod 2', '"abc" \\ 2', '"abc" * 2', '"abc" / 2', '-"abc"', '+"abc"', '"abc" ^ 2', '2 ^ -"abc"',
			'"abc" And 1', '"abc" Or 1', '"abc" Xor 1', '"abc" Eqv 1', '"abc" Imp 1', '1 + "abc"', '"abc" < 1']) {
			expect(mismatches(arms(condition)), condition).toHaveLength(1);
		}
		for (const condition of ['"6" Mod 4', '"6" \\ 4', '"6" * 1', '-"1"', '"True" Or 0', '"1" + 1 = 2', '"2" < 3']) {
			expect(mismatches(arms(condition)), condition).toHaveLength(0);
		}
	});

	it('reports nothing for an expression it cannot read', () => {
		for (const condition of ['("abc" - 1', '"abc" -', '"abc" 1', '"abc" - Not', '"abc" - "oops']) {
			expect(mismatches(arms(condition)), condition).toHaveLength(0);
		}
	});

	it('names the directive and quotes the string as VBA writes it', () => {
		const long = 'x'.repeat(50);
		const [diagnostic] = mismatches(arms(`"a""b" & "${long}"`));
		expect(diagnostic.message).toContain(`"a""b${'x'.repeat(37)}..."`);
		expect(mismatches(fn('#If False Then', '#ElseIf "abc" Then', '#End If'))[0].message).toContain('#ElseIf condition');
	});
});

describe('the project index', () => {
	const helper = (condition: string) =>
		`#If ${condition} Then\nPublic Sub Helper(ByVal x As Long)\nEnd Sub\n#Else\nPublic Sub Helper()\nEnd Sub\n#End If\n`;
	const caller = 'Sub Run()\n    Helper\n    Helper 1\nEnd Sub\n';

	/** The argument-count findings in the calling module, by line. */
	function argumentCountLines(project: ProjectIndex): string[] {
		return analyzeVbaModuleSource({
			source: caller, moduleName: 'Module2', moduleKind: 'standard',
			...projectAnalysisOptionsForModule(project, 'Module2'),
		}).diagnostics
			.filter((d) => d.code === 'argument-count')
			.map((d) => caller.slice(caller.lastIndexOf('\n', d.span.start) + 1, caller.indexOf('\n', d.span.start)).trim());
	}

	it('keeps the declarations of the arm the VBE compiles', () => {
		// Excel compiled and ran Module2 calling Helper with no argument.
		const project = buildVbaProjectIndex([
			{ moduleName: 'Module1', source: helper('"0"') },
			{ moduleName: 'Module2', source: caller },
		]);
		expect(argumentCountLines(project)).toEqual(['Helper 1']);
	});

	it('follows an edit of the string, and of the #Const it reads', () => {
		const project = buildVbaProjectIndex([
			{ moduleName: 'Module1', source: helper('"0"') },
			{ moduleName: 'Module2', source: caller },
		]);
		project.setModule({ moduleName: 'Module1', moduleKind: 'standard', source: helper('"1"') });
		expect(argumentCountLines(project)).toEqual(['Helper']);
		project.setModule({ moduleName: 'Module1', moduleKind: 'standard', source: `#Const S = "0"\n${helper('S')}` });
		expect(argumentCountLines(project)).toEqual(['Helper 1']);
	});
});

describe('the cost of reading skipped lines', () => {
	it('lexes each skipped line once, even after a #Const holding a string', () => {
		// With a string constant any later condition may be a string, so a
		// skipped #ElseIf is read too: once each, never once per earlier line.
		const count = 1000;
		const source = '#Const S = "x"\n#If True Then\nPublic Taken As Long\n' +
			'#ElseIf Not Missing Then\n'.repeat(count) + '#End If\n';
		const module = parseModule(source);
		const spy = vi.spyOn(lexer, 'tokenize');
		try {
			expect(invalidConditionDirectives(module, OPTIONS.conditionalCompilation)).toEqual([]);
			expect(spy.mock.calls.length).toBeGreaterThanOrEqual(count);
			expect(spy.mock.calls.length).toBeLessThanOrEqual(count + 4);
		} finally {
			spy.mockRestore();
		}
	});
});
