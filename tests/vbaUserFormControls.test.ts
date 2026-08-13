import { describe, expect, it } from 'vitest';
import { parseUserFormControls } from '../src/vbaUserFormControls';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

// Issue #17. A form's controls are members of the form's class, declared by the
// designer rather than by any line of code, so code-behind referring to them is
// correct VBA. The declarations live in the .frm header's Begin blocks.
const FORM_HEADER = [
	'VERSION 5.00',
	'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm ',
	'   Caption         =   "Entry"',
	'   ClientHeight    =   3000',
	'   Begin Forms.ComboBox.1 RegionPick ',
	'      Height          =   24',
	'   End',
	'   Begin Forms.CheckBox.1 Taxable ',
	'      Caption         =   "Taxable"',
	'   End',
	'   Begin Forms.TextBox.1 NameBox ',
	'      Height          =   18',
	'   End',
	'End',
	'Attribute VB_Name = "EntryForm"',
	'',
].join('\r\n');

const CODE = [
	'Option Explicit',
	'',
	'Private Sub UserForm_Initialize()',
	'    RegionPick.AddItem "North"',
	'    Taxable.Value = True',
	'End Sub',
	'',
].join('\r\n');

describe('controls a form designer declares', () => {
	it('finds every control with its name and type', () => {
		const controls = parseUserFormControls(FORM_HEADER + CODE);
		expect(controls.map((c) => c.name)).toEqual(['RegionPick', 'Taxable', 'NameBox']);
		expect(controls.map((c) => c.type)).toEqual([
			'MSForms.ComboBox', 'MSForms.CheckBox', 'MSForms.TextBox',
		]);
	});

	it('keeps the designer prog id as written', () => {
		expect(parseUserFormControls(FORM_HEADER)[0].progId).toBe('Forms.ComboBox.1');
	});

	it('does not treat the form itself as one of its own controls', () => {
		expect(parseUserFormControls(FORM_HEADER).some((c) => c.name === 'EntryForm')).toBe(false);
	});

	it('finds controls nested inside a container', () => {
		const nested = [
			'VERSION 5.00',
			'Begin {GUID} EntryForm ',
			'   Begin Forms.Frame.1 GroupBox ',
			'      Begin Forms.OptionButton.1 FirstChoice ',
			'      End',
			'   End',
			'End',
			'',
		].join('\r\n');
		expect(parseUserFormControls(nested).map((c) => c.name))
			.toEqual(['GroupBox', 'FirstChoice']);
	});

	it('passes an unrecognised prog id through rather than guessing', () => {
		const custom = [
			'VERSION 5.00',
			'Begin {GUID} EntryForm ',
			'   Begin MSComctlLib.TreeView.2 Tree ',
			'   End',
			'End',
			'',
		].join('\r\n');
		expect(parseUserFormControls(custom)[0].type).toBe('MSComctlLib.TreeView.2');
	});

	it.each([
		['a standard module', 'Attribute VB_Name = "Module1"\r\nOption Explicit\r\n'],
		['a class module', 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\nOption Explicit\r\n'],
		['empty source', ''],
	])('returns nothing for %s', (_label, source) => {
		expect(parseUserFormControls(source)).toEqual([]);
	});

	it('stops at an unterminated header rather than inventing controls from code', () => {
		const broken = 'VERSION 5.00\r\nBegin {GUID} EntryForm \r\n' + CODE;
		expect(parseUserFormControls(broken)).toEqual([]);
	});
});

describe('a form code-behind referring to its controls', () => {
	const FORM = [
		'VERSION 5.00',
		'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm ',
		'   Caption         =   "Entry"',
		'   Begin Forms.ComboBox.1 RegionPick ',
		'   End',
		'   Begin Forms.CheckBox.1 Taxable ',
		'   End',
		'   Begin Forms.TextBox.1 NameBox ',
		'   End',
		'End',
		'Attribute VB_Name = "EntryForm"',
		'Option Explicit',
		'',
		'Private Sub UserForm_Initialize()',
		'    RegionPick.AddItem "North"',
		'    Taxable.Value = True',
		'    NameBox.Text = "hi"',
		'End Sub',
		'',
		'Private Sub NotAControl()',
		'    Debug.Print Mystery',
		'End Sub',
		'',
	].join('\r\n');

	function analyzeForm(): string[] {
		const project = buildVbaProjectIndex([
			{ moduleName: 'EntryForm', source: FORM, type: 'form' },
		]);
		const options = projectAnalysisOptionsForModule(project, 'EntryForm');
		// The editor shows the body; the analyzer sees the body, while the
		// controls come from the header the index still holds.
		const body = FORM.slice(FORM.indexOf('Option Explicit'));
		return analyzeVbaModuleSource({
			source: body,
			moduleName: 'EntryForm',
			moduleType: 'form',
			moduleKind: 'userform',
			...options,
		} as never).diagnostics.map((d) => `${d.code}:${d.message}`);
	}

	it('reports no undeclared variable for a control', () => {
		const found = analyzeForm();
		for (const control of ['RegionPick', 'Taxable', 'NameBox']) {
			expect(found.some((f) => f.includes(control)), control).toBe(false);
		}
	});

	it('still reports a name that is not a control', () => {
		// The whole point of the rule has to survive the fix.
		expect(analyzeForm().some((f) => f.startsWith('undeclared-variable') && f.includes('Mystery')))
			.toBe(true);
	});

	it('carries each control type, not only its name', () => {
		const project = buildVbaProjectIndex([
			{ moduleName: 'EntryForm', source: FORM, type: 'form' },
		]);
		const options = projectAnalysisOptionsForModule(project, 'EntryForm');
		expect(options.implicitMembers).toEqual([
			{ name: 'RegionPick', progId: 'Forms.ComboBox.1', type: 'MSForms.ComboBox' },
			{ name: 'Taxable', progId: 'Forms.CheckBox.1', type: 'MSForms.CheckBox' },
			{ name: 'NameBox', progId: 'Forms.TextBox.1', type: 'MSForms.TextBox' },
		]);
	});

	it('adds nothing for a module that is not a form', () => {
		const project = buildVbaProjectIndex([
			{ moduleName: 'Module1', source: 'Attribute VB_Name = "Module1"\r\nOption Explicit\r\n' },
		]);
		expect(projectAnalysisOptionsForModule(project, 'Module1').implicitMembers).toBeUndefined();
	});
});
