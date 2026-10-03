// Diagnostics tests: what a form's designer puts in its MultiPage and lists,
// against the form's own code (issue #315). Measured on 2026-10-02 in 64-bit
// Excel 16.0 (build 20430) on a form F1 built through VBIDE, its code called
// from Module1 as `Dim f As New F1: Main = f.Probe()`.

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CONTROLS = [
	{ name: 'T1', type: 'MSForms.TextBox' },
	{ name: 'L1', type: 'MSForms.ListBox', listStartsEmpty: true },
	{ name: 'Cm', type: 'MSForms.ComboBox', listStartsEmpty: true },
	{ name: 'Mp', type: 'MSForms.MultiPage', pages: ['Page1', 'Page2'] },
];

const CALLER = 'Option Explicit\nPublic Function Main() As Variant\n    Dim f As New F1\n    Main = f.Probe()\nEnd Function\n';

function raised(body: string, extra = '', others: string[] = []): string[] {
	const src = `Option Explicit\n${extra}Public Function Probe() As Variant\n    ${body}\n    If IsEmpty(Probe) Then Probe = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'F1', source: src, moduleKind: 'userform', implicitMembers: CONTROLS },
		{ moduleName: 'Module1', source: CALLER },
		...others.map((source, n) => ({ moduleName: `Other${n}`, source })),
	], 'F1', { moduleKind: 'userform', implicitMembers: CONTROLS })
		.filter((d) => d.severity === 'error')
		.map((d) => /Run-time error '(-?\d+)'/.exec(d.message)?.[1] ?? d.code);
}

describe('a MultiPage\'s pages, from the designer (issue #315)', () => {
	it('reports a page past the count or a name it lacks', () => {
		for (const body of ['Probe = Mp.Pages(9).Caption', 'Probe = Mp.Pages(2).Caption', 'Probe = Mp.Pages(-1).Caption', 'Probe = Mp.Pages("Page9").Caption', 'Probe = Me.Mp.Pages(9).Caption']) {
			expect(raised(body), body).toEqual(['5']);
		}
		expect(raised('Mp.Value = 5')).toEqual(['380']);
		expect(raised('Mp.Value = 2')).toEqual(['380']);
	});

	it('matches a page name in any case, and leaves Value = -1 alone', () => {
		for (const body of ['Probe = Mp.Pages("page2").Caption', 'Probe = Mp.Pages("PAGE1").Caption', 'Mp.Value = -1']) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('takes the pages it has, and leaves alone a procedure that changes them', () => {
		for (const body of ['Probe = Mp.Pages(1).Caption', 'Probe = Mp.Pages("Page2").Caption', 'Probe = Mp.Pages.Count', 'Mp.Value = 1',
			'Mp.Pages.Add\n    Probe = Mp.Pages(2).Caption', 'Dim ps As Object\n    Set ps = Mp.Pages\n    ps.Add\n    Probe = Mp.Pages(2).Caption',
			'With Mp\n        .Pages.Add\n    End With\n    Probe = Mp.Pages(2).Caption']) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('leaves alone a MultiPage named elsewhere', () => {
		expect(raised('Probe = Mp.Pages(2).Caption', 'Private Sub UserForm_Initialize()\n    Mp.Pages.Add\nEnd Sub\n')).toEqual([]);
		expect(raised('Probe = Mp.Pages(2).Caption', '', ['Public Sub Grow(f As Object)\n    f.Mp.Pages.Add\nEnd Sub\n'])).toEqual([]);
		expect(raised('Probe = Mp.Pages(2).Caption', '', ['Public Sub Grow(f As Object)\n    f.Controls(3).Pages.Add\nEnd Sub\n'])).toEqual([]);
	});
});

describe('a list the code fills, from empty (issue #315)', () => {
	it('reports ListIndex, Selected and List past the items added', () => {
		const cases: Array<[string, string]> = [
			['L1.AddItem "a"\n    L1.ListIndex = 5', '380'],
			['L1.ListIndex = 0', '380'],
			['L1.ListIndex = -2', '380'],
			['L1.AddItem "a"\n    L1.Clear\n    L1.ListIndex = 0', '380'],
			['L1.AddItem "a"\n    L1.RemoveItem 0\n    L1.ListIndex = 0', '380'],
			['L1.List = Array("a", "b")\n    L1.ListIndex = 2', '380'],
			['Cm.AddItem "a"\n    Cm.ListIndex = 5', '380'],
			['Me.L1.AddItem "a"\n    Me.L1.ListIndex = 5', '380'],
			['L1.AddItem "a"\n    L1.Selected(5) = True', '380'],
			['L1.AddItem "a"\n    Probe = L1.List(5)', '381'],
			['L1.AddItem "a"\n    L1.ListIndex = L1.ListCount', '380'],
			['Cm.AddItem "a"\n    Cm.ListIndex = Cm.ListCount', '380'],
			['L1.List = Array()\n    L1.ListIndex = 0', '380'],
			['L1.AddItem "a"\n    L1.Selected(-1) = True', '380'],
			['L1.AddItem "a"\n    Probe = L1.List(-1)', '381'],
		];
		for (const [body, error] of cases) {
			expect(raised(body), body).toEqual([error]);
		}
	});

	it('takes an index inside the list', () => {
		for (const body of ['L1.AddItem "a"\n    L1.ListIndex = 0', 'L1.ListIndex = -1', 'L1.AddItem "a"\n    L1.AddItem "b"\n    L1.ListIndex = 1',
			'L1.List = Array("a", "b")\n    L1.ListIndex = 1', 'L1.AddItem "a"\n    Probe = L1.List(0)', 'Cm.ListIndex = -1', 'L1.AddItem "a"\n    Probe = L1.Selected(0)']) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('does not count items a loop or a branch may add', () => {
		expect(raised('Dim i As Long\n    For i = 1 To 3\n        L1.AddItem "a"\n    Next\n    L1.ListIndex = 2')).toEqual([]);
		expect(raised('If Probe = 1 Then L1.AddItem "a"\n    L1.ListIndex = 0')).toEqual([]);
	});

	it('leaves alone a list filled anywhere else', () => {
		const init = 'Private Sub UserForm_Initialize()\n    Dim i As Long\n    For i = 1 To 10\n        L1.AddItem "x"\n    Next\nEnd Sub\n';
		expect(raised('L1.ListIndex = 5', init)).toEqual([]);
		expect(raised('L1.ListIndex = 5', '', ['Public Sub Fill(f As Object)\n    f.L1.AddItem "x"\nEnd Sub\n'])).toEqual([]);
		expect(raised('L1.ListIndex = 5', '', ['Public Sub Fill(f As Object)\n    Dim c As Object\n    For Each c In f.Controls\n        c.AddItem "x"\n    Next\nEnd Sub\n'])).toEqual([]);
		expect(raised('Fill L1\n    L1.ListIndex = 5', 'Private Sub Fill(lb As Object)\n    lb.AddItem "x"\nEnd Sub\n')).toEqual([]);
	});

	it('leaves alone a list with a RowSource, though ListCount still bounds it', () => {
		const bound = CONTROLS.map((c) => (c.name === 'L1' ? { name: 'L1', type: 'MSForms.ListBox' } : c));
		const src = 'Option Explicit\nPublic Function Probe() As Variant\n    L1.ListIndex = 5\n    L1.ListIndex = L1.ListCount\nEnd Function\n';
		const codes = analyzeProjectModule(src, [{ moduleName: 'F1', source: src, moduleKind: 'userform', implicitMembers: bound }], 'F1', { moduleKind: 'userform', implicitMembers: bound })
			.filter((d) => d.severity === 'error')
			.map((d) => /Run-time error '(-?\d+)'/.exec(d.message)?.[1]);
		expect(codes).toEqual(['380']);
	});
});
