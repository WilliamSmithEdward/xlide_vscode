// github.com/WilliamSmithEdward/xlide_vscode/issues/68. A module that names a
// member after one of the host's globals used to lose it: `Rows` is an Excel
// global, so a class declaring `Public Property Get rows() As Widget` had
// `rows.Where(p)` measured against `Excel.Range` and reported member-not-found
// on legal code. A module VARIABLE of the same name already won, which is what
// pinned the fault to the procedure path rather than to the name.

import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';

/** The reported member-not-found messages, which is all this is about. */
function memberNotFound(source: string, moduleType: 'class' | 'standard' | 'document' = 'class'): string[] {
	return analyzeModule(source, { moduleName: 'Klass', moduleType })
		.filter((d) => d.code === 'member-not-found')
		.map((d) => d.message);
}

function moduleNaming(declaration: string, end: string): string {
	return [
		'Option Explicit',
		'',
		declaration,
		end,
		'',
		'Public Function Where(ByVal p As Variant) As Widget',
		'    Set Where = Nothing',
		'End Function',
		'',
		'Public Function Use(ByVal p As Variant) As Widget',
		'    Set Use = rows.Where(p)',
		'End Function',
	].join('\r\n');
}

describe('a module member named after a host global', () => {
	it('wins the bare name from a Property Get, as it already did from a variable', () => {
		expect(memberNotFound(moduleNaming('Public Property Get rows() As Widget', 'End Property'))).toEqual([]);
	});

	it('wins it from a Function too', () => {
		expect(memberNotFound(moduleNaming('Public Function rows() As Widget', 'End Function'))).toEqual([]);
	});

	it('still wins it as a variable, which never lost it', () => {
		const source = [
			'Option Explicit',
			'Public rows As Widget',
			'',
			'Public Function Where(ByVal p As Variant) As Widget',
			'    Set Where = Nothing',
			'End Function',
			'',
			'Public Function Use(ByVal p As Variant) As Widget',
			'    Set Use = rows.Where(p)',
			'End Function',
		].join('\r\n');
		expect(memberNotFound(source)).toEqual([]);
	});

	it('wins it in every module kind, since the shadowing is the module\'s', () => {
		for (const moduleType of ['class', 'standard', 'document'] as const) {
			expect(memberNotFound(moduleNaming('Public Property Get rows() As Widget', 'End Property'), moduleType), moduleType)
				.toEqual([]);
		}
	});

	it('shadows with a Sub as well, which has no value but is still the name', () => {
		// `rows.Where` is not legal against a Sub either, but the module owns
		// the name: answering for Excel's Range would be answering the wrong
		// question.
		const source = [
			'Option Explicit',
			'Public Sub rows()',
			'End Sub',
			'',
			'Public Function Use(ByVal p As Variant) As Variant',
			'    Use = rows.Where(p)',
			'End Function',
		].join('\r\n');
		expect(memberNotFound(source)).toEqual([]);
	});

	it('covers the other names every workbook collides with', () => {
		for (const name of ['rows', 'columns', 'cells', 'selection', 'names', 'sheets', 'application']) {
			const source = [
				'Option Explicit',
				`Public Property Get ${name}() As Widget`,
				'End Property',
				'',
				'Public Function Where(ByVal p As Variant) As Widget',
				'    Set Where = Nothing',
				'End Function',
				'',
				'Public Function Use(ByVal p As Variant) As Widget',
				`    Set Use = ${name}.Where(p)`,
				'End Function',
			].join('\r\n');
			expect(memberNotFound(source), name).toEqual([]);
		}
	});
});

describe('what the shadowing must not swallow', () => {
	it("binds the procedure's own return type, rather than merely going quiet", () => {
		// The shadowed name answers as a Workbook now, so a member Workbook has
		// not got is still reported, and reported against Workbook. Silence
		// here would mean the fix had only suppressed the check.
		const source = [
			'Option Explicit',
			'Public Property Get rows() As Workbook',
			'End Property',
			'',
			'Public Sub Use()',
			'    rows.NotAMember',
			'End Sub',
		].join('\r\n');
		expect(memberNotFound(source)).toEqual([
			"Method or data member not found: 'Excel.Workbook.NotAMember'.",
		]);
	});

	it('leaves an unshadowed host global to the host', () => {
		// Nothing here declares `Rows`, so Excel still answers for it - and
		// still reports a member Range has not got.
		const source = [
			'Option Explicit',
			'Public Sub Use()',
			'    Rows.NotAMember',
			'End Sub',
		].join('\r\n');
		expect(memberNotFound(source)).toEqual([
			"Method or data member not found: 'Excel.Range.NotAMember'.",
		]);
	});

	it('leaves a local of the same name winning over the module member', () => {
		const source = [
			'Option Explicit',
			'Public Property Get rows() As Widget',
			'End Property',
			'',
			'Public Sub Use()',
			'    Dim rows As Excel.Range',
			'    rows.NotAMember',
			'End Sub',
		].join('\r\n');
		expect(memberNotFound(source)).toEqual([
			"Method or data member not found: 'Excel.Range.NotAMember'.",
		]);
	});
});
