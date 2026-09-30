// Diagnostics tests: UserForm code (issue #226). Each case was measured in
// 64-bit Excel 16.0 (build 20326, 2026-09-30) on a form F1 built through
// VBIDE: TextBox T1, ListBox L1, CheckBox C1, Frame Fr holding TextBox T2,
// and one control of each other class.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CONTROLS = [
	{ name: 'T1', type: 'MSForms.TextBox' },
	{ name: 'L1', type: 'MSForms.ListBox' },
	{ name: 'C1', type: 'MSForms.CheckBox' },
	{ name: 'Fr', type: 'MSForms.Frame' },
	{ name: 'T2', type: 'MSForms.TextBox' },
	{ name: 'CB', type: 'MSForms.CommandButton' },
	{ name: 'Lb', type: 'MSForms.Label' },
	{ name: 'Cm', type: 'MSForms.ComboBox' },
	{ name: 'Ob', type: 'MSForms.OptionButton' },
	{ name: 'Tg', type: 'MSForms.ToggleButton' },
	{ name: 'Sp', type: 'MSForms.SpinButton' },
	{ name: 'Sc', type: 'MSForms.ScrollBar' },
	{ name: 'Im', type: 'MSForms.Image' },
	{ name: 'Mp', type: 'MSForms.MultiPage' },
	{ name: 'Ts', type: 'MSForms.TabStrip' },
];

function fromModule(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim f As New F1\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'F1', source: 'Option Explicit\nPublic Field As Long\n', moduleKind: 'userform', implicitMembers: CONTROLS },
	], 'Module1').filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
}

function insideForm(code: string): string[] {
	const src = `Option Explicit\n${code}`;
	return analyzeProjectModule(src, [
		{ moduleName: 'F1', source: src, moduleKind: 'userform', implicitMembers: CONTROLS },
	], 'F1', { moduleKind: 'userform', implicitMembers: CONTROLS }).filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
}

describe("VBA's own UserForm events (issue #226)", () => {
	const handler = (signature: string, body = '    Debug.Print 1\n') => byCode(analyzeModule(`Option Explicit\nPrivate Sub ${signature}\n${body}End Sub\n`, { moduleKind: 'userform', implicitMembers: [] }), 'event-handler-signature');

	it.each([
		'UserForm_Initialize(ByVal x As Long)',
		'UserForm_Terminate(ByVal x As Long)',
		'UserForm_Resize(ByVal x As Long)',
		'UserForm_Activate(ByVal x As Long)',
		'UserForm_Deactivate(ByVal x As Long)',
		'UserForm_QueryClose(Cancel As Boolean, CloseMode As Integer)',
		'UserForm_QueryClose(ByVal Cancel As Integer, CloseMode As Integer)',
		'UserForm_QueryClose(Cancel As Integer, ByVal CloseMode As Integer)',
		'UserForm_QueryClose(Cancel As Integer)',
	])('flags %s', (signature) => {
		expect(handler(signature)).toHaveLength(1);
	});

	it.each(['UserForm_Initialize()', 'UserForm_Activate()', 'UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)'])('takes %s', (signature) => {
		expect(handler(signature)).toHaveLength(0);
	});

	it('leaves an empty handler alone, as the VBE does', () => {
		expect(handler('UserForm_Initialize(ByVal x As Long)', '')).toHaveLength(0);
	});
});

describe("a control's members through the form (issue #226)", () => {
	it.each(['f.T1.Nope', 'f.T1.ListCount', 'f.L1.Nope', 'f.C1.Nope', 'f.CB.Nope', 'f.Lb.Nope', 'f.Cm.Nope', 'f.Sp.Nope', 'f.Sc.Nope', 'f.Im.Nope', 'F1.T1.Nope', 'f.T2.Nope'])('flags Main = %s', (expression) => {
		const hits = fromModule(`Main = ${expression}`);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^member-not-found: /);
	});

	it.each([
		'f.T1.Text', 'f.T1.RowSource', 'f.T1.BoundValue', 'f.T1.FontBold', 'f.T1.InSelection', 'f.T1.Left', 'f.T1.ControlTipText',
		'f.L1.ListCount', 'f.C1.Value', 'f.CB.Caption', 'f.Sp.Max', 'f.Im.PictureSizeMode',
		'f.Fr.Nope', 'f.Ob.Nope', 'f.Tg.Nope', 'f.Mp.Nope', 'f.Ts.Nope', 'f.T1.Font.Nope',
	])('stays quiet on Main = %s', (expression) => {
		expect(fromModule(`Main = ${expression}`)).toHaveLength(0);
	});

	it('does not check a variable declared as the control class', () => {
		expect(fromModule('Dim t As MSForms.TextBox', 'Main = t.Nope')).toHaveLength(0);
	});

	it('checks Me.T1 and a bare T1 inside the form, unless a local takes the name', () => {
		expect(insideForm('Public Function P() As Variant\n    P = Me.T1.Nope\nEnd Function\n')).toHaveLength(1);
		expect(insideForm('Public Function P() As Variant\n    P = T1.Nope\nEnd Function\n')).toHaveLength(1);
		// The local T1 is still Nothing there, which object-variable-not-set says.
		expect(insideForm('Public Function P() As Variant\n    Dim T1 As Object\n    P = T1.Nope\nEnd Function\n').filter((hit) => hit.startsWith('member-not-found'))).toHaveLength(0);
	});
});

describe('Set on a control, and a control name the form lacks (issue #226)', () => {
	it.each(['Set f.T1 = Nothing', 'Set f.Fr = Nothing'])('flags %s at compile time', (line) => {
		const hits = fromModule(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^set-requires-object: .*Invalid use of property/);
	});

	it('flags f.Controls("Nope"), error -2147024809', () => {
		const hits = fromModule('Main = f.Controls("Nope").Text');
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain("'-2147024809'");
	});

	it.each(['Main = f.Controls("T1").Text', 'Main = f.Controls("t1").Name', 'Main = f.Controls("T2").Name'])('stays quiet on %s', (line) => {
		expect(fromModule(line)).toHaveLength(0);
	});

	it('flags Me.Controls("Nope") inside the form', () => {
		expect(insideForm('Public Function P() As Variant\n    P = Me.Controls("Nope").Name\nEnd Function\n')).toHaveLength(1);
	});
});
