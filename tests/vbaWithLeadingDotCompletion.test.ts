import { describe, it, expect } from 'vitest';
import { resolveMemberCompletions } from '../src/analyzer';

/** Member names offered just after the dot in the first occurrence of `marker`. */
function memberNames(src: string, marker: string): string[] {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error('marker not found: ' + marker);
	}
	return resolveMemberCompletions(src, idx + marker.length, {}).map((m) => m.name);
}

const withApplication = (body: string): string =>
	['Sub Foo()', '    With Application', body, '    End With', 'End Sub'].join('\n');

describe('implicit-With leading-dot member completion in expression positions', () => {
	it('completes a leading dot after `For Each ... In` (the reported case)', () => {
		const src = withApplication('        Dim wb As Workbook\n        For Each wb In .');
		const names = memberNames(src, 'In .');
		expect(names).toContain('Workbooks');
		expect(names.length).toBeGreaterThan(10);
	});

	it('completes a leading dot on the RHS of `Set x =`', () => {
		const src = withApplication('        Dim x As Workbook\n        Set x = .');
		expect(memberNames(src, 'Set x = .')).toContain('Workbooks');
	});

	it('completes a leading dot inside a call argument list', () => {
		const src = withApplication('        Debug.Print TypeName(.');
		expect(memberNames(src, 'TypeName(.')).toContain('Workbooks');
	});

	it('still completes a leading dot on its own statement line', () => {
		expect(memberNames(withApplication('        .'), '        .')).toContain('Workbooks');
	});

	it('completes a chained leading dot in expression position (`.Workbooks.`)', () => {
		const src = withApplication('        Dim n As Long\n        n = .Workbooks.');
		expect(memberNames(src, '.Workbooks.')).toContain('Count');
	});

	it('does not treat an explicit (unresolved) receiver as an implicit-With dot', () => {
		const src = withApplication('        Dim q As Long\n        q = bogusVar.');
		expect(memberNames(src, 'bogusVar.')).toEqual([]);
	});

	it('does not treat a bracketed identifier [Foo] before a dot as implicit-With', () => {
		const src = withApplication('        Dim q As Long\n        q = [Foo].');
		expect(memberNames(src, '[Foo].')).toEqual([]);
	});
});
