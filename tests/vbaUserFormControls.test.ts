import { describe, expect, it } from 'vitest';
import { parseUserFormControls } from '../src/vbaUserFormControls';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { resolveMemberCompletions, resolveHover, resolveSignatureHelp } from '../src/analyzer';
import { resolveMemberSurfaceAt } from '../src/analyzer/completion/memberAccess';

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

describe('a control offers its own type members', () => {
	// The second half of #17: knowing the NAME stops the false finding, knowing
	// the TYPE is what makes `RegionPick.` useful.
	const IMPLICIT = [
		{ name: 'RegionPick', type: 'MSForms.ComboBox' },
		{ name: 'Taxable', type: 'MSForms.CheckBox' },
		{ name: 'NameBox', type: 'MSForms.TextBox' },
	];

	function membersAfterDot(controlName: string): string[] {
		const source = `Private Sub UserForm_Initialize()\r\n    ${controlName}.\r\nEnd Sub\r\n`;
		const at = source.indexOf(`${controlName}.`) + controlName.length + 1;
		return resolveMemberCompletions(source, at, { implicitMembers: IMPLICIT } as never)
			.map((m) => m.name);
	}

	it('offers ComboBox members, including the one from the report', () => {
		const members = membersAfterDot('RegionPick');
		expect(members).toContain('AddItem');
		expect(members).toContain('ListIndex');
		expect(members.length).toBeGreaterThan(20);
	});

	it('offers CheckBox members', () => {
		expect(membersAfterDot('Taxable')).toContain('Value');
	});

	it('offers TextBox members', () => {
		expect(membersAfterDot('NameBox')).toContain('Text');
	});

	it('gives each control its own members, not a shared set', () => {
		// AddItem belongs to a ComboBox and not to a TextBox; a fix that typed
		// every control the same would pass the tests above and fail this one.
		expect(membersAfterDot('NameBox')).not.toContain('AddItem');
	});

	it('offers nothing for a name that is not a control', () => {
		const source = 'Private Sub T()\r\n    Mystery.\r\nEnd Sub\r\n';
		expect(resolveMemberCompletions(source, source.indexOf('Mystery.') + 8, {
			implicitMembers: IMPLICIT,
		} as never)).toEqual([]);
	});
});

describe('hover and signature help on a form s controls', () => {
	// Issue #19. Completion answered on both the control and its members while
	// hover answered nothing on either, and the call tip had no signature to
	// show for a control method.
	const IMPLICIT = [
		{ name: 'RegionPick', type: 'MSForms.ComboBox' },
		{ name: 'NameBox', type: 'MSForms.TextBox' },
	];
	const CTX = {
		moduleName: 'EntryForm',
		moduleKind: 'userform',
		meProjectType: 'EntryForm',
		meType: 'MSForms.UserForm',
		implicitMembers: IMPLICIT,
	};
	const SOURCE = [
		'Option Explicit',
		'',
		'Private Sub UserForm_Initialize()',
		'    RegionPick.AddItem "North"',
		'End Sub',
		'',
	].join('\r\n');

	it('describes the control itself, with the type the designer gave it', () => {
		const at = SOURCE.indexOf('RegionPick') + 3;
		const info = resolveHover(SOURCE, at, CTX as never);
		expect(info?.signature).toBe('RegionPick As MSForms.ComboBox');
		expect(info?.details.join(' ')).toContain('EntryForm');
	});

	it('describes a member of the control', () => {
		const at = SOURCE.indexOf('.AddItem') + 4;
		const info = resolveHover(SOURCE, at, CTX as never);
		expect(info?.signature).toContain('AddItem');
		expect(info?.details.join(' ')).toContain('ComboBox');
	});

	it('carries the call signature the forms metadata holds', () => {
		// Qualified by its owner, the way every other member hover reads.
		const at = SOURCE.indexOf('.AddItem') + 4;
		expect(resolveHover(SOURCE, at, CTX as never)?.signature)
			.toBe('ComboBox.AddItem([pvargItem As Variant], [pvargIndex As Variant])');
	});

	it('offers a call tip inside a control method call', () => {
		const source = 'Private Sub T()\r\n    RegionPick.AddItem "North"\r\nEnd Sub\r\n';
		const info = resolveSignatureHelp(source, source.indexOf('"North"'), CTX as never);
		expect(info?.label).toContain('AddItem');
		expect(info?.parameters.map((p) => p.label)).toEqual([
			'[pvargItem As Variant]', '[pvargIndex As Variant]',
		]);
	});

	it('says nothing about a name that is not a control', () => {
		const source = 'Private Sub T()\r\n    Mystery.Thing\r\nEnd Sub\r\n';
		expect(resolveHover(source, source.indexOf('Mystery') + 3, CTX as never)).toBeUndefined();
	});
});

describe('Me in a form', () => {
	// Issue #18, second half: control RECEIVERS were typed, but the form's own
	// member surface behind `Me.` was not enumerated at all.
	const IMPLICIT = [
		{ name: 'RegionPick', type: 'MSForms.ComboBox' },
		{ name: 'NameBox', type: 'MSForms.TextBox' },
	];
	const FORM_CTX = {
		meProjectType: 'EntryForm',
		meType: 'MSForms.UserForm',
		implicitMembers: IMPLICIT,
		projectClassMembers: [{
			name: 'EntryForm',
			kind: 'userform',
			members: [{ name: 'Cancelled', kind: 'property', returns: 'Boolean' }],
		}],
	};

	function membersAfter(expression: string, ctx: unknown = FORM_CTX): string[] {
		const source = `Private Sub UserForm_Initialize()\r\n    ${expression}\r\nEnd Sub\r\n`;
		const at = source.indexOf(expression) + expression.length;
		return resolveMemberCompletions(source, at, ctx as never).map((m) => m.name);
	}

	it('offers the controls the designer declared', () => {
		const members = membersAfter('Me.');
		expect(members).toContain('RegionPick');
		expect(members).toContain('NameBox');
	});

	it('offers the form module s own code', () => {
		expect(membersAfter('Me.')).toContain('Cancelled');
	});

	it('offers the MSForms.UserForm surface it inherits', () => {
		const members = membersAfter('Me.');
		expect(members).toContain('Caption');
		expect(members).toContain('Controls');
		expect(members).toContain('Repaint');
	});

	it('offers what VBA adds to a form that MSForms does not carry', () => {
		// Show and Hide are the two members form code uses most and neither is
		// in the Microsoft Forms type library; nor are Name, Tag or the
		// position properties. Verified on a live form instance in Excel.
		const members = membersAfter('Me.');
		for (const name of ['Show', 'Hide', 'Move', 'Name', 'Tag', 'Left', 'Top', 'Visible']) {
			expect(members, name).toContain(name);
		}
	});

	it('keeps the form s additions off a control', () => {
		// A ComboBox is not a form: offering Show on one would mean the two
		// surfaces had been merged rather than kept apart.
		const source = 'Private Sub T()\r\n    RegionPick.\r\nEnd Sub\r\n';
		const names = resolveMemberCompletions(source, source.indexOf('RegionPick.') + 11, {
			implicitMembers: IMPLICIT,
		} as never).map((m) => m.name);
		expect(names).toContain('AddItem');
		expect(names).not.toContain('Show');
	});

	it('chains through a control to that control s own members', () => {
		expect(membersAfter('Me.RegionPick.')).toContain('AddItem');
		expect(membersAfter('Me.NameBox.')).not.toContain('AddItem');
	});

	it('offers the same surface through the form s predeclared instance', () => {
		// `EntryForm.RegionPick` is how the same form is addressed by name.
		expect(membersAfter('EntryForm.')).toContain('RegionPick');
	});

	it('never claims the surface is complete', () => {
		// A form carries designer and extender members this does not enumerate,
		// so absence must not become a diagnostic about form code.
		const source = 'Private Sub T()\r\n    Me.\r\nEnd Sub\r\n';
		const surface = resolveMemberSurfaceAt(source, source.indexOf('Me.') + 3, FORM_CTX as never);
		expect(surface?.exhaustive).toBe(false);
	});

	it('offers another module s controls to nobody', () => {
		// The context carries one control list - the edited module's. Reaching a
		// different form by name must not hand out this form's controls.
		expect(membersAfter('OtherForm.', {
			...FORM_CTX,
			projectClassMembers: [
				...FORM_CTX.projectClassMembers,
				{ name: 'OtherForm', kind: 'userform', members: [] },
			],
		})).not.toContain('RegionPick');
	});
});
