// Diagnostics tests: runtime errors that come through a name in a string, a
// Null a built-in returns, a Case Is value, a Scripting.Dictionary and an
// Excel cell reference (issue #243). Each raising case was measured in Excel
// 16.0 (build 20326, 2026-09-30); each quiet neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const CLASS = 'Option Explicit\nPublic Function Hello() As String\n    Hello = "hi"\nEnd Function\nPublic Sub Go()\nEnd Sub\nPrivate Function Hidden() As Long\nEnd Function\n';

function project(main: string, extra: { moduleName: string; source: string; type?: 'class' | 'standard' }[] = []) {
	return analyzeProjectModule(main, [{ moduleName: 'Module1', source: main }, { moduleName: 'Class1', source: CLASS, type: 'class' }, ...extra], 'Module1');
}

describe('CallByName', () => {
	it.each([
		['Dim c As New Class1', 'Main = CallByName(c, "NoSuch", VbMethod)', '"NoSuch"', "'438'"],
		['Dim c As New Class1', 'Main = CallByName(c, "Hidden", VbMethod)', '"Hidden"', "'438'"],
		['Dim c As New Collection', 'Main = CallByName(c, "NoSuch", VbGet)', '"NoSuch"', "'438'"],
		['Dim c As New Class1', 'CallByName c, "Hello", VbLet, 1', 'VbLet', "'450'"],
		['Dim c As New Class1', 'CallByName c, "Go", VbLet, 1', 'VbLet', "'450'"],
		['Dim c As New Class1', 'Main = CallByName(c, "Hello", VbGet)', 'VbGet', "'450'"],
	])('reports %s then %s', (decl, call, span, error) => {
		const src = wrap(decl, call);
		expectDiagnostic(src, byCode(project(src), 'runtime-member-not-found'), 'runtime-member-not-found', { span, message: error });
	});

	it('stays quiet on a member that fits, in any case, and on a Collection\'s own members', () => {
		for (const lines of [
			['Dim c As New Class1', 'Main = CallByName(c, "Hello", VbMethod)'],
			['Dim c As New Class1', 'Main = CallByName(c, "HELLO", VbMethod)'],
			['Dim c As New Collection', 'CallByName c, "Add", VbMethod, 1'],
			['Dim c As New Collection', 'Main = CallByName(c, "Item", VbMethod, 1)'],
		]) {
			expect(byCode(project(wrap(...lines)), 'runtime-member-not-found'), lines.join('; ')).toHaveLength(0);
		}
	});
});

describe('Application.Run', () => {
	const helpers = 'Public Sub Helper()\nEnd Sub\nPrivate Sub PrivHelper()\nEnd Sub\n';

	it('reports a name no standard or document module has, or only a class has', () => {
		for (const name of ['NoSuchMacro', 'Module1.NoSuch', 'Hello']) {
			const src = `${wrap(`Application.Run "${name}"`)}${helpers}`;
			expectDiagnostic(src, byCode(project(src), 'runtime-member-not-found'), 'runtime-member-not-found', { span: `"${name}"`, message: "'1004'" });
		}
	});

	it('stays quiet on a Sub or Function of any module, Private ones included', () => {
		const other = { moduleName: 'Module2', source: 'Option Explicit\nPrivate Sub Hidden()\nEnd Sub\n' };
		for (const name of ['Helper', 'PrivHelper', 'Module1.Helper', 'Hidden', 'Module2.Hidden', 'Book1.xlsm!Anything']) {
			const src = `${wrap(`Application.Run "${name}"`)}${helpers}`;
			expect(byCode(project(src, [other]), 'runtime-member-not-found'), name).toHaveLength(0);
		}
	});

	it('is not judged without the project, or outside Excel', () => {
		const src = wrap('Application.Run "NoSuchMacro"');
		expect(byCode(analyzeModule(src), 'runtime-member-not-found')).toHaveLength(0);
		const word = analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }], 'Module1', { host: 'word' });
		expect(byCode(word, 'runtime-member-not-found')).toHaveLength(0);
	});
});

describe('Null from Choose, Switch and IIf into a typed variable', () => {
	it.each([
		['Dim n As Long', 'n = Choose(5, 1, 2)', 'Choose(5, 1, 2)'],
		['Dim s As String', 's = Choose(0, "a")', 'Choose(0, "a")'],
		['Dim n As Long', 'n = Switch(False, 1)', 'Switch(False, 1)'],
		['Dim n As Long', 'n = IIf(False, 1, Null)', 'IIf(False, 1, Null)'],
	])('reports %s then %s', (decl, statement, span) => {
		const src = wrap(decl, statement);
		expectDiagnostic(src, byCode(analyzeModule(src), 'assignment-type-mismatch'), 'assignment-type-mismatch', { span, message: "'94'" });
	});

	it('stays quiet on a choice in range, or a Variant target', () => {
		for (const lines of [['Dim n As Long', 'n = Choose(2, 1, 2)'], ['Dim v As Variant', 'v = Choose(5, 1, 2)'], ['Dim n As Long', 'n = IIf(True, 1, Null)'], ['Dim n As Long', 'n = Switch(True, 1)'], ['Dim n As Long', 'n = Switch(False, 1, True, 2)']]) {
			expect(byCode(analyzeModule(wrap(...lines)), 'assignment-type-mismatch'), lines.join('; ')).toHaveLength(0);
		}
		// The module's own Choose is not VBA's.
		const own = `${wrap('Dim n As Long', 'n = Choose(5, 1, 2)')}Private Function Choose(ByVal i As Long, ByVal a As Long, ByVal b As Long) As Long\nEnd Function\n`;
		expect(byCode(analyzeModule(own), 'assignment-type-mismatch')).toHaveLength(0);
	});
});

describe('Case Is with a string against a number', () => {
	it('reports Case Is > "abc" under Select Case 5, and leaves Case Is > 3', () => {
		const src = wrap('Select Case 5', '    Case Is > "abc"', '        Main = 1', 'End Select');
		expectDiagnostic(src, analyzeModule(src), 'string-arithmetic-coercion', { span: '"abc"', message: "'13'" });
		expect(byCode(analyzeModule(wrap('Select Case 5', '    Case Is > 3', '        Main = 1', 'End Select')), 'string-arithmetic-coercion')).toHaveLength(0);
	});
});

describe('Scripting.Dictionary from CreateObject', () => {
	const make = ['Dim d As Object', 'Set d = CreateObject("Scripting.Dictionary")'];

	it('reports a key added twice, a key removed that is not there, and Keys past the end', () => {
		const dup = wrap(...make, 'd.Add "a", 1', 'd.Add "a", 2');
		expectDiagnostic(dup, byCode(analyzeModule(dup), 'collection-key-in-use'), 'collection-key-in-use', { span: '"a"', message: "'457'" });
		const missing = wrap(...make, 'd.Remove "a"');
		expectDiagnostic(missing, analyzeModule(missing), 'collection-key-not-found', { span: '"a"', message: "'32811'" });
		const past = wrap(...make, 'd.Add "a", 1', 'Main = d.Keys()(3)');
		expectDiagnostic(past, analyzeModule(past), 'collection-index-out-of-range', { span: '3', message: ['0 to 0', "'9'"] });
		const justPast = wrap(...make, 'd.Add "a", 1', 'Main = d.Items()(1)');
		expectDiagnostic(justPast, analyzeModule(justPast), 'collection-index-out-of-range', { span: '1' });
		const cleared = wrap(...make, 'd.Add "a", 1', 'd.RemoveAll', 'd.Remove "a"');
		expectDiagnostic(cleared, analyzeModule(cleared), 'collection-key-not-found', { span: '"a"' });
	});

	it('takes keys as written, and a read adds the key it does not find', () => {
		const quiet = [
			[...make, 'd.Add "a", 1', 'd.Add "A", 2'],
			[...make, 'd.Add "1", 1', 'd.Add 1, 2'],
			[...make, 'Main = IsEmpty(d("missing"))', 'd.Remove "missing"'],
			[...make, 'd("k") = 1', 'Main = d.Keys()(0)'],
			[...make, 'd.Add "a", 1', 'd.RemoveAll', 'd.Add "a", 2'],
		];
		for (const lines of quiet) {
			const diags = analyzeModule(wrap(...lines));
			for (const code of ['collection-key-in-use', 'collection-key-not-found', 'collection-index-out-of-range']) {
				expect(byCode(diags, code), `${lines.join('; ')}: ${code}`).toHaveLength(0);
			}
		}
		// CompareMode is followed since issue #349: an empty Dictionary still
		// has no "a" to remove (measured in Excel 16.0, 32811).
		const compared = wrap(...make, 'd.CompareMode = 1', 'd.Remove "a"');
		expectDiagnostic(compared, analyzeModule(compared), 'collection-key-not-found', { span: '"a"', message: "'32811'" });
	});
});

describe('Excel cell references past the grid', () => {
	it.each([
		['Cells(1, "AAAA")', '"AAAA"', "'13'"],
		['Cells(1, "XFE")', '"XFE"', "'13'"],
		['Cells(1, "")', '""', "'13'"],
		['Cells(1, "A1")', '"A1"', "'1004'"],
		['Cells(1, "5")', '"5"', "'1004'"],
		['Range("A1:")', '"A1:"', "'1004'"],
		['Range(":A1")', '":A1"', "'1004'"],
	])('reports %s', (call, span, error) => {
		const src = wrap(`Main = ${call}.Address`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'host-argument-out-of-range'), 'host-argument-out-of-range', { span, message: error });
	});

	it('stays quiet on a column Excel has', () => {
		for (const column of ['"XFD"', '"a"', '"$A"', '"AB"']) {
			expect(byCode(analyzeModule(wrap(`Main = Cells(1, ${column}).Address`)), 'host-argument-out-of-range'), column).toHaveLength(0);
		}
	});
});
