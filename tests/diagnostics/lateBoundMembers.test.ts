// Diagnostics tests: a member the VBE binds at run time on an object whose
// class and member list are known (issue #121). Measured in Excel 16.0 (build
// 20326, 2026-09-26): each raising sample compiles and raises 438; each quiet
// one runs.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'runtime-member-not-found';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('runtime-member-not-found (issue #121)', () => {
	it('flags a name Excel.Application has not got and that is no worksheet function', () => {
		const src = wrap('Application.Zzq', 'Main = 1');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'Zzq', message: ["'438'", 'worksheet function'] });
		expect(byCode(analyzeModule(src), 'member-not-found')).toHaveLength(0);
	});

	it('stays quiet for Application members, worksheet functions, and other hosts', () => {
		const src = wrap('Main = Application.Match(1, Array(1), 0)', 'Main = Application.Version', 'Application.Volatile', 'Main = Application.WorksheetFunction.Sum(1)');
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
		const word = wrap('Application.Zzq', 'Main = 1');
		expect(byCode(analyzeModule(word, { hostModel: getWordObjectModel() }), CODE)).toHaveLength(0);
	});

	it('flags a member a Collection has not got on an Object set to New Collection', () => {
		const src = wrap('Dim o As Object', 'Set o = New Collection', 'o.Foo', 'Main = o.Count');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'Foo', message: ['Collection', "'438'"] });
	});

	it('stops following the variable after a block, a reassignment or a pass elsewhere', () => {
		const src = wrap(
			'Dim o As Object, p As Object, q As Object',
			'Set o = New Collection',
			'If Main Then Set o = Nothing',
			'o.Foo',
			'Set p = New Collection',
			'Set p = CreateObject("Scripting.Dictionary")',
			'p.Foo',
			'Set q = New Collection',
			'Reshape q',
			'q.Foo',
			'Main = 1',
		) + 'Sub Reshape(ByRef x As Object)\nEnd Sub\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});
