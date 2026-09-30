// Diagnostics tests: the follow-up cases issue #219 gathered from #96, #114,
// #119 and #121. Each raising case was measured in Excel 16.0 (build 20326,
// 2026-09-30) and raises the error named every time; each quiet neighbour
// runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function count(code: string, ...lines: string[]): number {
	return byCode(analyzeModule(wrap(...lines)), code).length;
}

describe('a Private member of a document module (issue #114)', () => {
	const sheet = 'Private Function Secret() As String\nEnd Function\nPrivate m As Long\nPublic Function Open1() As String\nEnd Function\n';

	it.each([
		['Main = Sheet1.Secret()', 'Secret'],
		['Sheet1.Secret', 'Secret'],
		['Main = Sheet1.m', 'm'],
		['With Sheet1\n        Main = .Secret()\n    End With', 'Secret'],
	])('flags %s from another module', (statement, member) => {
		const src = wrap(statement);
		const diags = analyzeProjectModule(src, [{ moduleName: 'Sheet1', source: sheet, type: 'document' }, { moduleName: 'Module1', source: src }], 'Module1');
		expectDiagnostic(src, diags, 'member-not-found', { span: member, message: ['Sheet1', 'Private'] });
	});

	it('stays quiet on a Public member and on a late-bound sheet', () => {
		const src = wrap('Main = Sheet1.Open1()', 'Main = Worksheets(1).Secret()', 'Main = Sheet1.Name', 'Sheet1.Range("A1").Value = 1');
		const diags = analyzeProjectModule(src, [{ moduleName: 'Sheet1', source: sheet, type: 'document' }, { moduleName: 'Module1', source: src }], 'Module1');
		expect(byCode(diags, 'member-not-found')).toHaveLength(0);
	});

	it('flags Me.Secret() inside the sheet, and keeps a bare Secret() quiet', () => {
		const own = `${sheet}Public Function T1() As String\n    T1 = Me.Secret()\n    T1 = Secret()\nEnd Function\n`;
		for (const host of [undefined, 'excel']) {
			const diags = analyzeProjectModule(own, [{ moduleName: 'Sheet1', source: own, type: 'document' }], 'Sheet1', { moduleKind: 'document', host });
			expectDiagnostic(own, diags, 'member-not-found', { span: 'Secret', message: 'Private' });
		}
	});
});

describe('Switch with an odd number of arguments (issue #96)', () => {
	it.each(['Main = Switch(True, 1, False)', 'Main = Switch(True)', 'Main = Switch(True, 1, False, 2, True)'])('flags %s', (statement) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { message: ['pairs', "Run-time error '5'"] });
	});

	it('stays quiet on pairs', () => {
		expect(count('runtime-argument-value', 'Main = Switch(True, 1, False, 2)')).toBe(0);
	});
});

describe('division by an Empty Variant or a zero conversion (issue #119)', () => {
	it.each([
		[['Dim v As Variant', 'Main = 5 / v'], "'11'"],
		[['Dim v', 'Main = 5 / v'], "'11'"],
		[['Dim v As Variant', 'Main = v / v'], "'6'"],
		[['Dim v As Variant', 'Main = 10 \\ v'], "'11'"],
		[['Dim v As Variant', 'Main = 10 Mod v'], "'11'"],
		[['Main = 10 / CLng(0)'], "'11'"],
		[['Main = 10 / (CLng(0))'], "'11'"],
		[['Main = 10 / CLng(0.4)'], "'11'"],
		[['Main = 10 / CDbl(0)'], "'11'"],
		[['Main = 0 / CLng(0)'], "'6'"],
	])('flags %j', (lines, error) => {
		const src = wrap(...lines);
		expectDiagnostic(src, analyzeModule(src), 'division-by-zero', { message: error });
	});

	it.each([
		[['Dim v As Variant', 'Main = v / 5']],
		[['Dim v As Variant', 'v = 2', 'Main = 10 / v']],
		[['Dim v As Variant', 'If Timer > 0 Then v = 2', 'Main = 5 / v']],
		[['Dim v As Variant', 'Fill v', 'Main = 5 / v']],
		[['Dim v As Variant, x As Variant', 'For Each x In Array(1)', '    Main = 5 / x', 'Next']],
		[['Main = 10 / CDbl(0.4)']],
		[['Main = 10 / CLng(0.6)']],
	])('stays quiet on %j', (lines) => {
		const src = `${wrap(...lines)}Sub Fill(ByRef x As Variant)\n    x = 1\nEnd Sub\n`;
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(0);
	});
});

describe('Collection Add arguments (issue #121)', () => {
	const add = (...lines: string[]): string => wrap('Dim c As New Collection', ...lines, 'Main = c.Count');

	it.each([
		[['c.Add "x", 5'], 'collection-add-argument', "'13'"],
		[['c.Add "x", 1.5'], 'collection-add-argument', "'13'"],
		[['c.Add "x", True'], 'collection-add-argument', "'13'"],
		[['c.Add Item:="x", Key:=5'], 'collection-add-argument', "'13'"],
		[['Dim k As Variant', 'c.Add "x", k'], 'collection-add-argument', "'13'"],
		[['c.Add "a"', 'c.Add "b", , 1, 1'], 'collection-add-argument', "'5'"],
		[['c.Add "x", Before:=1'], 'collection-index-out-of-range', "'5'"],
		[['c.Add "x", , 1'], 'collection-index-out-of-range', "'5'"],
		[['c.Add "x", After:=1'], 'collection-index-out-of-range', "'5'"],
		[['c.Add "a"', 'c.Add "x", Before:=2'], 'collection-index-out-of-range', "'9'"],
		[['c.Add "a"', 'c.Add "x", After:=2'], 'collection-index-out-of-range', "'9'"],
		[['c.Add "a"', 'c.Add "x", Before:=0'], 'collection-index-out-of-range', "'9'"],
	])('flags %j', (lines, code, error) => {
		const src = add(...lines);
		expectDiagnostic(src, analyzeModule(src), code, { message: error });
	});

	it.each([
		[['c.Add "x", "5"']],
		[['c.Add "a"', 'c.Add "x", Before:=1']],
		[['c.Add "a"', 'c.Add "x", After:=1']],
		[['Dim k As Variant', 'k = "key"', 'c.Add "x", k']],
	])('stays quiet on %j', (lines) => {
		const diags = analyzeModule(add(...lines));
		expect([...byCode(diags, 'collection-add-argument'), ...byCode(diags, 'collection-index-out-of-range')]).toHaveLength(0);
	});
});

describe('a scalar or Empty Variant where an array goes (issue #121)', () => {
	it.each([
		[['Dim v As Variant', 'v = 5', 'Erase v'], 'Erase'],
		[['Dim v As Variant', 'v = "abc"', 'Erase v'], 'Erase'],
		[['Dim v As Variant', 'Erase v'], 'Empty'],
		[['Dim v As Variant', 'v = 5', 'ReDim Preserve v(2)'], 'ReDim Preserve'],
		[['Dim v As Variant', 'ReDim Preserve v(2)', 'Main = UBound(v)'], 'Empty'],
		[['Dim v As Variant, x As Variant', 'v = 5', 'For Each x In v', 'Next'], 'For Each'],
		[['Dim v As Variant, x As Variant', 'v = "abc"', 'For Each x In v', 'Next'], 'For Each'],
		[['Dim v As Variant, x As Variant', 'For Each x In v', 'Next'], 'Empty'],
	])('flags %j', (lines, text) => {
		const src = wrap(...lines);
		expectDiagnostic(src, analyzeModule(src), 'variant-value-misuse', { span: 'v', message: [text, "'13'"] });
	});

	it.each([
		[['Dim v As Variant', 'v = 5', 'ReDim v(2)', 'Main = UBound(v)']],
		[['Dim v As Variant, x As Variant', 'v = Split("a,b", ",")', 'For Each x In v', 'Next']],
		[['Dim v As Variant', 'v = Array(1)', 'ReDim Preserve v(2)']],
		[['Dim v As Variant', 'Main = 1', 'If Main = 2 Then v = Array(1)', 'ReDim Preserve v(2)']],
		[['Dim v As Variant', 'Again:', 'ReDim Preserve v(2)', 'v = 1', 'If Main = 0 Then GoTo Again']],
	])('stays quiet on %j', (lines) => {
		expect(byCode(analyzeModule(wrap(...lines)), 'variant-value-misuse')).toHaveLength(0);
	});
});

describe('New assigned without Set (issue #121)', () => {
	it.each([
		[['Dim x As Variant', 'x = New Collection'], "'450'"],
		[['Main = New Collection'], "'450'"],
	])('flags %j', (lines, error) => {
		const src = wrap(...lines);
		expectDiagnostic(src, analyzeModule(src), 'object-default-value', { span: 'New Collection', message: error });
	});

	it('flags a class with no default member, 438', () => {
		const src = wrap('Dim x As Variant', 'x = New Class1');
		const diags = analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, { moduleName: 'Class1', source: 'Public Name As String\n', type: 'class' }], 'Module1');
		expectDiagnostic(src, diags, 'object-default-value', { span: 'New Class1', message: "'438'" });
	});

	it('stays quiet on Set', () => {
		expect(count('object-default-value', 'Dim x As Variant', 'Set x = New Collection')).toBe(0);
	});
});
