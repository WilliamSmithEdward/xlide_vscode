import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { getVb6ObjectModel } from '../src/analyzer/host/vb6ObjectModel';
import { EMPTY_HOST_MODEL, hostObjectModelForToken } from '../src/analyzer/host/hostRegistry';
import {
	getHostEvents,
	getHostMembers,
	getHostType,
	resolveHostConstant,
	resolveHostGlobal,
} from '../src/analyzer/host/hostModel';
import {
	analyzeModule,
	resolveEventHandlerCompletions,
	resolveIdentifierCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
} from '../src/analyzer';
import { readVb6Modules } from '../src/vba/vb6/vb6Project';
import { frmMembers, parseFrmHeader } from '../src/vba/vb6/frmHeader';
import { blankDesignerHeader } from '../src/vba/moduleSource';

// Slice 3 of docs/roadmap_vb6_support.md: the vb6 host model. VBRUN is read
// from the type library inside msvbvm60.dll; VB (App, Screen, Printer,
// Clipboard, Global, Form, the intrinsic controls) is transcribed from
// twinBASIC's documentation because VB.OLB is not available. The model
// offers and describes; it never produces a red on its own.

const FIXTURES = path.join(__dirname, 'fixtures', 'vb6');

function members(source: string, marker: string, ctx: Record<string, unknown> = {}): string[] {
	const model = getVb6ObjectModel();
	return resolveMemberCompletions(source, source.indexOf(marker) + marker.length, { model, ...ctx } as never)
		.map((member) => member.name);
}

describe('the vb6 host model', () => {
	it('is registered for the vb6 token, as itself', () => {
		const model = hostObjectModelForToken('vb6');
		expect(model).toBe(getVb6ObjectModel());
		expect(model).not.toBe(EMPTY_HOST_MODEL);
		expect(model?.hostName).toBe('VB6');
	});

	it('declares every global against a type it actually carries', () => {
		const model = getVb6ObjectModel();
		for (const [name, qualified] of Object.entries(model.globals)) {
			expect(model.types[qualified], `${name} -> ${qualified}`).toBeDefined();
		}
		expect(resolveHostGlobal('App', model)).toBe('VB.App');
		expect(resolveHostGlobal('Screen', model)).toBe('VB.Screen');
		expect(resolveHostGlobal('Printer', model)).toBe('VB.Printer');
		expect(resolveHostGlobal('Clipboard', model)).toBe('VB.Clipboard');
	});

	it('proves nothing absent: no type claims exhaustiveness', () => {
		for (const [qualified, type] of Object.entries(getVb6ObjectModel().types)) {
			expect(type.exhaustive ?? false, qualified).toBe(false);
		}
	});

	it('carries both libraries under their own namespaces, each with its provenance', () => {
		const model = getVb6ObjectModel();
		const form = getHostType('VB.Form', model);
		expect(form?.provenance).toContain('twinbasic/documentation');
		expect(form?.provenance).toContain('docs/Reference/Default/VB/Form/index.md');
		const dataObject = getHostType('VBRUN.DataObject', model);
		expect(dataObject?.provenance).toContain('msvbvm60.dll resource 3');
		expect(getHostMembers('VBRUN.DataObject', model).map((m) => m.name)).toEqual(
			expect.arrayContaining(['GetData', 'SetData', 'GetFormat', 'Files']),
		);
		expect(model.aliases['dataobject']).toBe('VBRUN.DataObject');
		expect(model.aliases['form']).toBe('VB.Form');
	});

	it('answers App., Screen., Printer. and Clipboard. through the member resolver', () => {
		expect(members('Sub T()\n    App.\nEnd Sub\n', 'App.')).toEqual(
			expect.arrayContaining(['Path', 'EXEName', 'Title', 'PrevInstance', 'Major']),
		);
		expect(members('Sub T()\n    Screen.\nEnd Sub\n', 'Screen.')).toEqual(
			expect.arrayContaining(['ActiveForm', 'MousePointer', 'TwipsPerPixelX', 'Width']),
		);
		expect(members('Sub T()\n    Printer.\nEnd Sub\n', 'Printer.')).toEqual(
			expect.arrayContaining(['Print', 'EndDoc', 'NewPage', 'DeviceName']),
		);
		expect(members('Sub T()\n    Clipboard.\nEnd Sub\n', 'Clipboard.')).toEqual(
			expect.arrayContaining(['GetText', 'SetText', 'Clear']),
		);
	});

	it('answers Me. in a VB6 form as a VB.Form, and Forms as the collection', () => {
		const got = members('Sub T()\n    Me.\nEnd Sub\n', 'Me.', { meType: 'VB.Form' });
		expect(got).toEqual(expect.arrayContaining(['Caption', 'Show', 'Hide', 'Cls', 'Controls', 'WindowState']));
		// Events are not object members: Me.Load() is not a call.
		expect(got).not.toContain('Load');
		expect(got).not.toContain('QueryUnload');
		expect(members('Sub T()\n    Forms.\nEnd Sub\n', 'Forms.')).toEqual(expect.arrayContaining(['Count', 'Item']));
	});

	it('types a designer control by its prog id, including a control array element', () => {
		const controls = [
			{ name: 'Text1', type: 'VB.TextBox' },
			{ name: 'Command1', type: 'VB.CommandButton', array: true },
		];
		expect(members('Sub T()\n    Text1.\nEnd Sub\n', 'Text1.', { implicitMembers: controls })).toEqual(
			expect.arrayContaining(['Text', 'SelStart', 'SetFocus', 'MaxLength']),
		);
		expect(members('Sub T()\n    Command1(0).\nEnd Sub\n', 'Command1(0).', { implicitMembers: controls })).toEqual(
			expect.arrayContaining(['Caption', 'Default', 'Cancel', 'Value']),
		);
	});

	it('knows the intrinsic constants with their values', () => {
		const model = getVb6ObjectModel();
		expect(resolveHostConstant('vbKeyReturn', model)?.value).toBe(13);
		expect(resolveHostConstant('vbKeyReturn', model)?.type).toBe('KeyCodeConstants');
		expect(resolveHostConstant('vbModal', model)?.value).toBe(1);
		expect(resolveHostConstant('vbRSTypeTable', model)?.value).toBe(0);
		expect(model.enums?.FormWindowStateConstants).toBeDefined();
	});

	it('keeps its globals, Global members and constants out of the undeclared rule', () => {
		const src = [
			'Option Explicit',
			'Sub T()',
			'    Dim n As Long',
			'    n = vbKeyReturn',
			'    Debug.Print App.Path, Screen.Width, Printer.DeviceName',
			'    Set Me.Picture = LoadPicture(App.Path & "\\logo.bmp")',
			'    Unload Me',
			'End Sub',
			'',
		].join('\n');
		const findings = analyzeModule(src, { host: 'vb6', moduleKind: 'userform', knownIdentifiers: new Set<string>() } as never)
			.map((d) => d.message);
		expect(findings).toEqual([]);
		// The rule itself still works: an invented name stays a finding.
		const bogus = 'Option Explicit\nSub T()\n    Dim n As Long\n    n = vbNotAConstant\nEnd Sub\n';
		const flagged = analyzeModule(bogus, { host: 'vb6', knownIdentifiers: new Set<string>() });
		expect(flagged.some((d) => d.message.includes('vbNotAConstant'))).toBe(true);
	});

	it('offers a call tip for a Global function and a Form method', () => {
		const model = getVb6ObjectModel();
		const src = 'Sub T()\n    Set p = LoadPicture(App.Path)\nEnd Sub\n';
		const tip = resolveSignatureHelp(src, src.indexOf('(App') + 1, { model } as never);
		expect(tip?.label).toContain('LoadPicture(');
		const bare = resolveIdentifierCompletions('Sub T()\n    LoadPi\nEnd Sub\n', 'Sub T()\n    LoadPi'.length, {
			moduleName: 'Form1',
			model,
		}).find((item) => item.name === 'LoadPicture');
		expect(bare).toBeDefined();
	});

	it('carries a reserved member with its note, and events apart from members', () => {
		const model = getVb6ObjectModel();
		const maskColor = getHostMembers('VB.CommandButton', model).find((m) => m.name === 'MaskColor');
		expect(maskColor?.doc?.remarks).toMatch(/Reserved for compatibility with VB6/);
		const events = getHostEvents('VB.CommandButton', model).map((e) => e.name);
		expect(events).toEqual(expect.arrayContaining(['Click', 'MouseDown', 'KeyPress', 'GotFocus']));
		expect(getHostMembers('VB.CommandButton', model).map((m) => m.name)).not.toContain('Click');
		expect(getHostEvents('VB.Form', model).map((e) => e.name)).toEqual(
			expect.arrayContaining(['Load', 'Unload', 'QueryUnload', 'Resize', 'Activate']),
		);
	});
});

describe('a VB6 form offers its own event handler stubs', () => {
	const model = getVb6ObjectModel();
	const controls = [
		{ name: 'Text1', type: 'VB.TextBox' },
		{ name: 'Command1', type: 'VB.CommandButton', array: true },
	];
	function stubs(src: string, marker: string, meType = 'VB.Form') {
		return resolveEventHandlerCompletions(src, src.indexOf(marker) + marker.length, {
			moduleName: 'Form1',
			moduleKind: 'userform',
			host: 'vb6',
			model,
			meType,
			implicitMembers: controls,
		});
	}

	it('names the form handlers Form_*, with VB6 parameter lists', () => {
		const got = stubs('Option Explicit\nForm_\n', 'Form_');
		const byName = new Map(got.map((item) => [item.name, item.signature]));
		expect(byName.get('Form_Load')).toBe('Form_Load()');
		expect(byName.get('Form_QueryUnload')).toBe('Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)');
		expect(byName.get('Form_KeyPress')).toBe('Form_KeyPress(KeyAscii As Integer)');
		expect(byName.has('UserForm_Initialize')).toBe(false);
	});

	it('names an MDI form\'s handlers MDIForm_*', () => {
		const got = stubs('Option Explicit\nMDIForm_\n', 'MDIForm_', 'VB.MDIForm').map((item) => item.name);
		expect(got).toContain('MDIForm_Load');
		expect(got).not.toContain('Form_Load');
	});

	it('gives a control array\'s handlers Index As Integer first', () => {
		const got = stubs('Option Explicit\nCommand1_\n', 'Command1_');
		const byName = new Map(got.map((item) => [item.name, item.signature]));
		expect(byName.get('Command1_Click')).toBe('Command1_Click(Index As Integer)');
		expect(byName.get('Command1_MouseDown')).toBe(
			'Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)',
		);
		const text = new Map(stubs('Option Explicit\nText1_\n', 'Text1_').map((item) => [item.name, item.signature]));
		expect(text.get('Text1_Change')).toBe('Text1_Change()');
		expect(text.get('Text1_KeyPress')).toBe('Text1_KeyPress(KeyAscii As Integer)');
	});

	it('does not offer a stub for a handler the module already declares', () => {
		const got = stubs('Option Explicit\nPrivate Sub Form_Load()\nEnd Sub\nForm_\n', '\nForm_').map((item) => item.name);
		expect(got).not.toContain('Form_Load');
		expect(got).toContain('Form_Resize');
	});
});

describe('the fixture code-behind is a negative control', () => {
	// Real VB6 forms from four projects (tests/fixtures/vb6/*/NOTICE.md). With
	// the vb6 model in place, their code must draw no member or declaration
	// finding: every control, global and constant they touch is one the model
	// now knows, and the model never proves a member absent.
	const forms: { name: string; source: string; implicitMembers: { name: string; type: string }[] }[] = [];
	for (const vbp of ['RunAsTrustedInstaller/Project1.vbp', 'Diabetes-prediction-1.0/MCD_prj.vbp']) {
		for (const m of readVb6Modules(path.join(FIXTURES, vbp))) {
			if (m.type === 'userform' && m.source) {
				forms.push({ name: `${vbp}:${m.name}`, source: m.source, implicitMembers: m.implicitMembers ?? [] });
			}
		}
	}
	for (const rel of [
		'polyworks/frmPreferences.frm', 'polyworks/frmInfo.frm', 'audiostation/Form_OpenDialog.frm',
		'audiostation/Form_Settings_Record.frm', 'audiostation/Form_Track_Properties.frm',
	]) {
		const text = fs.readFileSync(path.join(FIXTURES, rel), 'latin1');
		const header = parseFrmHeader(text);
		forms.push({
			name: rel,
			source: blankDesignerHeader(text),
			implicitMembers: header ? frmMembers(header).map((m) => ({ name: m.name, type: m.type })) : [],
		});
	}

	// Header noise counts too: a form whose designer block leaked into the
	// analysis (audiostation's Form_OpenDialog opens with `Object =` OCX
	// references before `Begin`) answered 119 statements-outside-procedure.
	const WRONG = ['member-not-found', 'undeclared-variable', 'argument-count', 'statement-outside-procedure', 'option-after-declaration'];

	it.each(forms.map((form) => [form.name, form] as const))('%s draws no member, declaration or header finding', (_name, form) => {
		expect(form.implicitMembers.length).toBeGreaterThan(0);
		const findings = analyzeModule(form.source, {
			host: 'vb6',
			moduleName: form.name,
			moduleKind: 'userform',
			meType: 'VB.Form',
			implicitMembers: form.implicitMembers,
		} as never).filter((d) => WRONG.includes(d.code));
		expect(findings.map((d) => `${d.code} L${d.line}: ${d.message}`)).toEqual([]);
	});
});
