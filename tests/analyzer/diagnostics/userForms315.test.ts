// Diagnostics tests: UserForm code (issue #315). Measured in 64-bit Excel
// 16.0 on a form Wizard holding a TextBox Answer, a Label Prompt, a ListBox
// Choices and a MultiPage Pages.

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const CONTROLS = [
	{ name: 'Answer', type: 'MSForms.TextBox' },
	{ name: 'Prompt', type: 'MSForms.Label' },
	{ name: 'Choices', type: 'MSForms.ListBox' },
	{ name: 'Pages', type: 'MSForms.MultiPage' },
];

function insideForm(code: string): string[] {
	const src = `Option Explicit\n${code}`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Wizard', source: src, moduleKind: 'userform', implicitMembers: CONTROLS },
	], 'Wizard', { moduleKind: 'userform', implicitMembers: CONTROLS }).filter((d) => d.severity === 'error').map((d) => d.code);
}

function fromModule(code: string): string[] {
	const src = `Option Explicit\n${code}`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Wizard', source: 'Option Explicit\n', moduleKind: 'userform', implicitMembers: CONTROLS },
	], 'Module1').filter((d) => d.severity === 'error').map((d) => d.code);
}

describe('UserForm code the VBE runs (issue #315)', () => {
	it('takes a control Controls.Add named in the same procedure', () => {
		expect(insideForm('Public Function Probe() As Variant\n    Me.Controls.Add "Forms.TextBox.1", "Dyn"\n    Probe = Me.Controls("Dyn").Name\nEnd Function\n')).toEqual([]);
		expect(insideForm('Public Function Probe() As Variant\n    Me.Controls.Add bstrProgID:="Forms.TextBox.1", Name:="Dyn"\n    Probe = Me.Controls("Dyn").Name\nEnd Function\n')).toEqual([]);
		expect(insideForm('Public Function Probe(n As String) As Variant\n    Me.Controls.Add "Forms.TextBox.1", n\n    Probe = Me.Controls("Other").Name\nEnd Function\n')).toEqual([]);
	});

	it('does not assume other runtime control names are absent', () => {
		expect(insideForm('Public Function Probe() As Variant\n    Me.Controls.Add "Forms.TextBox.1", "Dyn"\n    Probe = Me.Controls("Nope").Name\nEnd Function\n')).toEqual([]);
	});

	it('knows UserForms, the loaded forms', () => {
		expect(fromModule('Function Main() As Variant\n    Main = UserForms.Count\nEnd Function\n')).toEqual([]);
	});
});

describe('UserForm code the VBE refuses (issue #315)', () => {
	it('is Invalid use of property on Set of a bare control', () => {
		expect(insideForm('Public Sub P()\n    Set Answer = Nothing\nEnd Sub\n')).toEqual(['set-requires-object']);
		expect(insideForm('Public Sub P()\n    Dim Answer As Object\n    Set Answer = Nothing\nEnd Sub\n')).toEqual([]);
	});

	it('does not infer runtime control indexes from the designer', () => {
		expect(insideForm('Public Function P() As Variant\n    P = Me.Controls(99).Name\nEnd Function\n')).toEqual([]);
		expect(insideForm('Public Function P() As Variant\n    P = Me.Controls(3).Name\nEnd Function\n')).toEqual([]);
		expect(insideForm('Public Function P() As Variant\n    Me.Controls.Add "Forms.TextBox.1", "Dyn"\n    P = Me.Controls(4).Name\nEnd Function\n')).toEqual([]);
	});

	it('raises 13 on a control of another class Set into a typed variable', () => {
		expect(insideForm('Public Sub P()\n    Dim t As MSForms.TextBox\n    Set t = Prompt\nEnd Sub\n')).toEqual(['assignment-object-type-mismatch']);
		expect(insideForm('Public Sub P()\n    Dim t As MSForms.TextBox\n    Set t = Answer\nEnd Sub\n')).toEqual([]);
		expect(insideForm('Public Sub P()\n    Dim t As MSForms.Control\n    Set t = Prompt\nEnd Sub\n')).toEqual([]);
	});
});
