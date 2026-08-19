import { describe, expect, it } from 'vitest';
import { parseUserFormControls } from '../src/vbaUserFormControls';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectEditorSymbolContextForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { collectImplicitMemberMethodTokens, resolveMemberCompletions, resolveHover, resolveSignatureHelp } from '../src/analyzer';
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

describe('semantic tokens for method calls on a form s controls', () => {
	// Issue #20. A resolved method call paints `function` (the color Len gets);
	// a property, an unresolved member, and anything in a module without the
	// control stay exactly as they are.
	const IMPLICIT = [
		{ name: 'RegionPick', type: 'MSForms.ComboBox' },
		{ name: 'Taxable', type: 'MSForms.CheckBox' },
		{ name: 'NameBox', type: 'MSForms.TextBox' },
	];
	const CTX = { implicitMembers: IMPLICIT, meType: 'MSForms.UserForm' };

	function tokensIn(line: string, ctx: unknown = CTX): { text: string; type: string }[] {
		const source = `Private Sub T()\r\n    ${line}\r\nEnd Sub\r\n`;
		return collectImplicitMemberMethodTokens(source, ctx as never)
			.map((t) => ({ text: source.slice(t.span.start, t.span.end), type: t.tokenType }));
	}

	it.each([
		['RegionPick.AddItem "North"', 'AddItem'],
		['NameBox.SetFocus', 'SetFocus'],
		['Me.Hide', 'Hide'],
		['Me.RegionPick.AddItem "North"', 'AddItem'],
	])('paints the method call in %s', (line, expected) => {
		expect(tokensIn(line)).toEqual([{ text: expected, type: 'function' }]);
	});

	it.each([
		['a property stays as it is', 'Taxable.Value = True'],
		['an unresolved member is never painted', 'RegionPick.NotAMember'],
		['a bare identifier is not a member access', 'AddItem = 1'],
		['another receiver s member is not this control s', 'other.RegionPick.AddItem "North"'],
	])('%s', (_label, line) => {
		expect(tokensIn(line)).toEqual([]);
	});

	it('paints nothing in a module without the control', () => {
		expect(tokensIn('AddItem = 1', {})).toEqual([]);
		expect(tokensIn('RegionPick.AddItem "North"', {})).toEqual([]);
	});

	it('a declaration shadowing the control name wins', () => {
		// Clear IS a ComboBox method, so only the shadow check keeps this quiet.
		const source = [
			'Private Sub T()',
			'    Dim RegionPick As Collection',
			'    RegionPick.Clear',
			'End Sub',
			'',
		].join('\r\n');
		expect(collectImplicitMemberMethodTokens(source, CTX as never)).toEqual([]);
	});

	it('Me outside a form paints nothing', () => {
		// Nothing says this module is a form: no controls, no form type.
		expect(tokensIn('Me.Hide', {})).toEqual([]);
		// A document module says what its Me is, and it is not a form.
		expect(tokensIn('Me.Hide', { meType: 'Excel.Workbook' })).toEqual([]);
		// An explicit Me type outranks the inference drawn from controls.
		expect(tokensIn('Me.Hide', { implicitMembers: IMPLICIT, meType: 'Excel.Workbook' }))
			.toEqual([]);
	});

	it('answers the whole table from the controls alone', () => {
		// The contract as the report wrote it takes { implicitMembers } and
		// nothing else, and `Me.Hide` is one of its rows. Controls exist only on
		// a form, so their presence is what says `Me` is a form.
		const only = { implicitMembers: IMPLICIT };
		expect(tokensIn('RegionPick.AddItem "North"', only)).toEqual([{ text: 'AddItem', type: 'function' }]);
		expect(tokensIn('NameBox.SetFocus', only)).toEqual([{ text: 'SetFocus', type: 'function' }]);
		expect(tokensIn('Me.Hide', only)).toEqual([{ text: 'Hide', type: 'function' }]);
	});

	it('leaves a With block s leading dot alone', () => {
		// `.Clear` opening a line is a member of the With receiver, not of
		// whatever the line above ended with - and Clear IS a ComboBox method,
		// so nothing but the statement boundary keeps this quiet.
		const source = [
			'Private Sub T()',
			'    With SomeCollection',
			'        Debug.Print RegionPick',
			'        .Clear',
			'    End With',
			'End Sub',
			'',
		].join('\r\n');
		expect(collectImplicitMemberMethodTokens(source, CTX as never)).toEqual([]);
	});

	it('still paints across a line continuation', () => {
		// A continued line is one statement, so the dot is not a leading dot.
		const source = [
			'Private Sub T()',
			'    RegionPick _',
			'        .AddItem "North"',
			'End Sub',
			'',
		].join('\r\n');
		expect(collectImplicitMemberMethodTokens(source, CTX as never)
			.map((t) => source.slice(t.span.start, t.span.end))).toEqual(['AddItem']);
	});

	it('spans cover the member identifier only', () => {
		const source = 'Private Sub T()\r\n    RegionPick.AddItem "North"\r\nEnd Sub\r\n';
		const [token] = collectImplicitMemberMethodTokens(source, CTX as never);
		expect(source.slice(token.span.start, token.span.end)).toBe('AddItem');
		expect(source[token.span.start - 1]).toBe('.');
	});
});

describe('a control carries the Control base surface', () => {
	// SetFocus, Move and ZOrder live on MSForms.Control - the base every placed
	// control extends - not in the per-type dumps, so without the merge the
	// members form code uses most on a control would not resolve.
	const IMPLICIT = [{ name: 'NameBox', type: 'MSForms.TextBox' }];

	function membersOf(expression: string): string[] {
		const source = `Private Sub T()\r\n    ${expression}\r\nEnd Sub\r\n`;
		const at = source.indexOf(expression) + expression.length;
		return resolveMemberCompletions(source, at, { implicitMembers: IMPLICIT } as never)
			.map((m) => m.name);
	}

	it('offers SetFocus and the geometry properties on a control', () => {
		const members = membersOf('NameBox.');
		for (const name of ['SetFocus', 'Move', 'ZOrder', 'Left', 'Top', 'Visible', 'Name']) {
			expect(members, name).toContain(name);
		}
	});

	it('keeps the per-type surface first', () => {
		expect(membersOf('NameBox.')).toContain('Text');
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

describe('a form reached from another module (#22)', () => {
	// #17 gave a form's controls to its own code-behind; from OUTSIDE the form
	// its type carried no members at all - not even Show. The controls belong
	// on the form's member surface in the project index, so every caller kind
	// resolves them the same way.
	const FORM_BODY = [
		'Option Explicit',
		'',
		'Public Cancelled As Boolean',
		'',
		'Private Sub UserForm_Initialize()',
		'End Sub',
		'',
	].join('\r\n');
	const CONTROLS = [
		{ name: 'RegionPick', type: 'MSForms.ComboBox' },
		{ name: 'NameBox', type: 'MSForms.TextBox' },
	];

	function callerCompletions(expression: string, formInput: object): string[] {
		const project = buildVbaProjectIndex([
			{ moduleName: 'EntryForm', moduleKind: 'userform', source: FORM_BODY, ...formInput },
			{ moduleName: 'Launcher', source: 'Attribute VB_Name = "Launcher"\r\nOption Explicit\r\n' },
		]);
		const context = projectEditorSymbolContextForModule(project, 'Launcher');
		const source = `Private Sub Drive()\r\n    Dim f As EntryForm\r\n    ${expression}\r\nEnd Sub\r\n`;
		const at = source.indexOf(expression) + expression.length;
		return resolveMemberCompletions(source, at, {
			projectClassMembers: context.analysisOptions.projectClassMembers,
		} as never).map((m) => m.name);
	}

	it('offers the controls through the default instance', () => {
		const members = callerCompletions('EntryForm.', { implicitMembers: CONTROLS });
		expect(members).toContain('RegionPick');
		expect(members).toContain('NameBox');
	});

	it('offers the form s own code and the form surface it inherits', () => {
		const members = callerCompletions('EntryForm.', { implicitMembers: CONTROLS });
		expect(members).toContain('Cancelled');
		expect(members).toContain('Show');
		expect(members).toContain('Hide');
		expect(members).toContain('Caption');
	});

	it('offers the same surface on a declared variable', () => {
		const members = callerCompletions('f.', { implicitMembers: CONTROLS });
		expect(members).toContain('RegionPick');
		expect(members).toContain('Show');
	});

	it('chains through a control to that control s own members', () => {
		expect(callerCompletions('EntryForm.RegionPick.', { implicitMembers: CONTROLS }))
			.toContain('AddItem');
		expect(callerCompletions('EntryForm.NameBox.', { implicitMembers: CONTROLS }))
			.not.toContain('AddItem');
	});

	it('reads the controls from a .frm header when no host supplies them', () => {
		const headerForm = [
			'VERSION 5.00',
			'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm ',
			'   Begin Forms.ComboBox.1 RegionPick ',
			'   End',
			'End',
			'Attribute VB_Name = "EntryForm"',
			FORM_BODY,
		].join('\r\n');
		const members = callerCompletions('EntryForm.', { source: headerForm });
		expect(members).toContain('RegionPick');
		expect(members).toContain('Show');
	});

	it('a class is untouched: still exhaustive, no form members', () => {
		const project = buildVbaProjectIndex([
			{ moduleName: 'Probe1', moduleKind: 'class', source: 'Attribute VB_Name = "Probe1"\r\nPublic Sub Greet()\r\nEnd Sub\r\n' },
			{ moduleName: 'Launcher', source: 'Attribute VB_Name = "Launcher"\r\nOption Explicit\r\n' },
		]);
		const context = projectEditorSymbolContextForModule(project, 'Launcher');
		const source = 'Private Sub T()\r\n    Dim c As Probe1\r\n    c.\r\nEnd Sub\r\n';
		const members = resolveMemberCompletions(source, source.indexOf('c.') + 2, {
			projectClassMembers: context.analysisOptions.projectClassMembers,
		} as never).map((m) => m.name);
		expect(members).toContain('Greet');
		expect(members).not.toContain('Show');
	});

	it('a declared MSForms type chains on its own', () => {
		// The corollary the control returns rely on: MSForms.<Type> resolves as
		// a declared object type, qualified spelling only.
		const source = 'Private Sub T()\r\n    Dim t As MSForms.TextBox\r\n    t.\r\nEnd Sub\r\n';
		const members = resolveMemberCompletions(source, source.indexOf('t.') + 2, {} as never)
			.map((m) => m.name);
		expect(members).toContain('Text');
		expect(members).toContain('SetFocus');
	});
});

describe('member access through a control member returned object (issue #32)', () => {
	// Views.SelectedItem hovers `As Tab` - the metadata knows the return - but
	// the chain refused the second hop into the returned object's surface, the
	// one receiver shape that dead-ended while host chains resolved.
	const IMPLICIT = [
		{ name: 'Views', type: 'MSForms.TabStrip' },
		{ name: 'ViewNote', type: 'MSForms.Label' },
	];
	const CTX = { implicitMembers: IMPLICIT };

	function membersAt(source: string, marker: string): string[] {
		return resolveMemberCompletions(source, source.indexOf(marker) + marker.length, CTX as never)
			.map((m) => m.name);
	}

	it('completes into the returned Tab surface', () => {
		const source = 'Private Sub T()\r\n    Views.SelectedItem.\r\nEnd Sub\r\n';
		const members = membersAt(source, 'Views.SelectedItem.');
		expect(members).toContain('Caption');
		expect(members).toContain('Index');
		expect(members.length).toBeGreaterThan(5);
	});

	it('hovers the member of the returned object, owner-qualified', () => {
		const source =
			'Private Sub T()\r\n    ViewNote.Caption = Views.SelectedItem.Caption & " view"\r\nEnd Sub\r\n';
		const at = source.indexOf('SelectedItem.Caption') + 'SelectedItem.'.length + 2;
		const info = resolveHover(source, at, CTX as never);
		expect(info?.signature).toBe('Tab.Caption As String');
		expect(info?.details).toContain('Tab property');
	});

	it('chains into a control Font the same way', () => {
		const source = 'Private Sub T()\r\n    ViewNote.Font.\r\nEnd Sub\r\n';
		const members = membersAt(source, 'ViewNote.Font.');
		expect(members).toContain('Bold');
		expect(members).toContain('Size');

		const hoverSource = 'Private Sub T()\r\n    ViewNote.Font.Bold = True\r\nEnd Sub\r\n';
		const info = resolveHover(hoverSource, hoverSource.indexOf('.Bold') + 2, CTX as never);
		expect(info?.signature).toBe('Font.Bold As Boolean');
	});

	it('does not invent a surface past a primitive member', () => {
		// Caption is a String; the library gives it no object surface, and the
		// VBE offers none - the chain must stay silent rather than guess.
		const source = 'Private Sub T()\r\n    ViewNote.Caption.\r\nEnd Sub\r\n';
		expect(membersAt(source, 'ViewNote.Caption.')).toEqual([]);
	});
});
