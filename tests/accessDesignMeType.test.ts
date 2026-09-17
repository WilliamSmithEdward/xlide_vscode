import { existsSync, readFileSync } from 'fs';
import * as path from 'path';
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	workspace: {
		onDidCloseTextDocument: vi.fn(() => ({ dispose: () => undefined })),
	},
}));

import * as vscode from 'vscode';
import {
	collectHostMemberMethodTokens,
	collectImplicitMemberMethodTokens,
	resolveEventHandlerCompletions,
	resolveHover,
	resolveMemberCompletions,
} from '../src/analyzer';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import {
	ACCESS_CONTROL_TYPES,
	accessDesignMembers,
	accessVbaIdentifier,
	type AccessDesign,
} from '../src/vba/access/accessDesign';
import { listModules, readModules } from '../src/vba/projectService';
import type { ProjectEngine } from '../src/projectEngine';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';
import {
	toEventHandlerCompletionContext,
	toMemberCompletionContext,
	VbaEditorProjectContextService,
} from '../src/vbaEditorProjectContext';
import { VbaProjectIndexService } from '../src/vbaProjectIndexService';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';

// `Me` in the code behind an Access form is an Access.Form, the way it is an
// Excel.Worksheet behind a sheet and a Word.Document in ThisDocument. Before
// this, an Access design's module got no host type at all: `Me.` offered
// nothing, a bare `Requery` was called undefined, and the form's controls
// were unknown. Everything below starts from the real database bytes.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const DATABASE = path.join(FIXTURES, 'AccessFormFixture.accdb');
const WORKBOOK = path.join(FIXTURES, 'FormFixture.xlsm');
const FORM = 'Form_Calculator';

/** An idiomatic bound form: controls, sections, form members bare and through Me, and fields. */
const BOUND_FORM_SOURCE = [
	'Option Compare Database',
	'Option Explicit',
	'',
	'Private Sub Form_Load()',
	'    Me.Caption = "Orders"',
	'    Me.Qty = 5',
	'    Qty = 6',
	'    If Me.Express Then Me.Lines.AddItem "x"',
	'    Me.Lines.RowSource = ""',
	'    Me.Detail.BackColor = 1',
	'    Requery',
	'    Me.Requery',
	'    DoCmd.Close',
	'    CustomerID = 4',
	'    Debug.Print Me.CustomerID, Me!OrderID, Nz(Me.Qty, 0)',
	'End Sub',
	'',
].join('\r\n');

function completionsAt(source: string, marker: string, ctx: Parameters<typeof resolveMemberCompletions>[2]): string[] {
	const offset = source.indexOf(marker) + marker.length;
	return resolveMemberCompletions(source.slice(0, offset), offset, ctx).map((item) => item.name);
}

describe('the engine says what an Access design is', () => {
	it('names the class behind a form and lists its sections and controls', () => {
		const form = readModules(DATABASE).find((entry) => entry.name === FORM);
		expect(form?.type).toBe('accessform');
		expect(form?.designerClass).toBe('Access.Form');
		const members = new Map((form?.implicitMembers ?? []).map((member) => [member.name, member.type]));
		expect(members.get('Detail')).toBe('Access.Section');
		expect(members.get('Qty')).toBe('Access.Textbox');
		expect(members.get('Express')).toBe('Access.Checkbox');
		expect(members.get('Lines')).toBe('Access.ListBox');
		expect(members.get('AddLine')).toBe('Access.CommandButton');
		expect(members.get('TotalLabel')).toBe('Access.Label');
		// The listing the tree uses says the same.
		const listed = listModules(DATABASE).find((entry) => entry.name === FORM);
		expect(listed?.designerClass).toBe('Access.Form');
		expect(listed?.implicitMembers).toEqual(form?.implicitMembers);
	});

	it('types every control it can name with a class the Access model carries', () => {
		// The types are a contract with the analyzer's model, which looks a type
		// up by its exact key: `Access.TextBox` would resolve to nothing there.
		const named = (name: string): AccessDesign['objects'][number]['records'] => [
			{ id: 1, code: 20, valueType: 0, width: 0, value: Buffer.from(name, 'utf16le') },
		];
		const design: AccessDesign = {
			header: Buffer.alloc(10),
			trailer: Buffer.alloc(4),
			objects: [
				{ records: named('TheForm') },
				...Object.keys(ACCESS_CONTROL_TYPES).map((code) => ({
					marker: 0xff,
					type: Number(code),
					records: named(`Item${code}`),
				})),
			],
		};
		const model = getAccessObjectModel();
		const missing = accessDesignMembers(design).filter((member) => !(member.type in model.types));
		expect(missing).toEqual([]);
		expect(accessDesignMembers(design)).toHaveLength(Object.keys(ACCESS_CONTROL_TYPES).length);
	});

	it('leaves a UserForm exactly as it was: no designer class, MSForms controls', () => {
		const form = readModules(WORKBOOK).find((entry) => entry.type === 'userform');
		expect(form?.designerClass).toBeUndefined();
		expect(form?.implicitMembers?.length).toBeGreaterThan(0);
		expect(form?.implicitMembers?.every((member) => /^(MSForms|ActiveX)\./.test(member.type))).toBe(true);
	});
});

describe('Me in the code behind an Access form', () => {
	const modules = readModules(DATABASE);
	const project = buildVbaProjectIndex(modules.map((entry) => ({
		moduleName: entry.name,
		type: entry.type,
		source: entry.source ?? '',
		implicitMembers: entry.implicitMembers,
		designerClass: entry.designerClass,
	})));
	const form = modules.find((entry) => entry.name === FORM)!;

	const analyze = (source: string, moduleName: string, extra: Record<string, unknown> = {}): string[] =>
		analyzeVbaModuleSource({
			source,
			moduleName,
			host: 'access',
			...projectAnalysisOptionsForModule(project, moduleName),
			...extra,
		} as Parameters<typeof analyzeVbaModuleSource>[0]).diagnostics.map((d) => `${d.code}: ${d.message}`);

	const formOptions = {
		moduleType: form.type,
		moduleKind: 'userform',
		designerClass: form.designerClass,
	};

	it('finds nothing wrong with the form the fixture ships', () => {
		expect(analyze(form.source ?? '', FORM, formOptions)).toEqual([]);
	});

	it('knows the form s own members unqualified, which used to be called undefined', () => {
		// `Requery` alone is Access.Form.Requery: the module IS the form.
		expect(analyze(BOUND_FORM_SOURCE, FORM, formOptions)).toEqual([]);
		// The same module without the class the engine now supplies.
		expect(analyze(BOUND_FORM_SOURCE, FORM, { ...formOptions, designerClass: undefined }))
			.toContain("unknown-call: Sub or Function not defined: 'Requery'.");
	});

	it('never calls a record-source field undeclared, though the control list is known', () => {
		// A bound form has a member for every field of its record source, and
		// only the running database knows those. `CustomerID` is one.
		const codes = analyze(BOUND_FORM_SOURCE, FORM, formOptions);
		expect(codes.filter((code) => code.startsWith('undeclared-variable'))).toEqual([]);
		expect(codes.filter((code) => code.startsWith('member-not-found'))).toEqual([]);
	});

	it('still reports a call that nothing defines', () => {
		const source = 'Option Explicit\r\n\r\nPrivate Sub Go()\r\n    NoSuchProcedure "x"\r\nEnd Sub\r\n';
		expect(analyze(source, FORM, formOptions)).toContain("unknown-call: Sub or Function not defined: 'NoSuchProcedure'.");
	});

	it('reaches the form from another module as an Access.Form, not a UserForm', () => {
		const surfaces = projectAnalysisOptionsForModule(project, 'Module1').projectClassMembers ?? [];
		const formType = surfaces.find((surface) => surface.name === FORM);
		expect(formType?.designerClass).toBe('Access.Form');
		// Its fields are members no list names, so the list proves nothing absent.
		expect(formType?.exhaustive).toBe(false);

		const source = [
			'Option Explicit',
			'',
			'Public Sub Refill()',
			`    ${FORM}.Requery`,
			`    ${FORM}.RecordSource = "Orders"`,
			`    Debug.Print ${FORM}.CustomerID, ${FORM}.Qty`,
			'End Sub',
			'',
		].join('\r\n');
		expect(analyze(source, 'Module1', { moduleType: 'standard', moduleKind: 'standard' })
			.filter((code) => code.startsWith('member-not-found'))).toEqual([]);

		const names = completionsAt(`${source}\r\n${FORM}.`, `\r\n${FORM}.`, {
			model: getAccessObjectModel(),
			projectClassMembers: surfaces,
		});
		expect(names).toEqual(expect.arrayContaining(['Requery', 'RecordSource', 'Qty', 'AddCurrent']));
		// Show and Hide are a UserForm's. An Access form has neither.
		expect(names).not.toContain('Hide');
		expect(names).not.toContain('StartUpPosition');
	});
});

describe('the editor, from the database on disk to the completion list', () => {
	const uriPath = `/${DATABASE.replace(/\\/g, '/')}`.replace(/^\/\//, '/');
	const form = () => readModules(DATABASE).find((entry) => entry.name === FORM)!;

	beforeEach(() => {
		(vscode.workspace.textDocuments as unknown[]).length = 0;
	});

	function openForm(source: string): vscode.TextDocument {
		const value = `xlide-vba:${uriPath}/${FORM}.bas`;
		const document = {
			uri: { scheme: 'xlide-vba', path: `${uriPath}/${FORM}.bas`, toString: () => value },
			version: 1,
			getText: () => source,
		};
		(vscode.workspace.textDocuments as unknown[]).push(document);
		return document as unknown as vscode.TextDocument;
	}

	function services(): { contexts: VbaEditorProjectContextService; projects: VbaProjectIndexService } {
		// The real engine behind the bridge, so nothing about the entry is assumed.
		const bridge = {
			call: vi.fn(async (method: string, payload: { path: string }) => {
				if (method === 'readModules') { return readModules(payload.path); }
				throw new Error(`Unexpected bridge call ${method}`);
			}),
		} as unknown as ProjectEngine;
		const projects = new VbaProjectIndexService(new VbaSymbolIndex(bridge));
		return { contexts: new VbaEditorProjectContextService(projects), projects };
	}

	it('offers Access.Form s members, the form s own code, and its controls after Me.', async () => {
		const source = `${BOUND_FORM_SOURCE}\r\nPublic Sub Recount()\r\nEnd Sub\r\n\r\nPrivate Sub Probe()\r\n    Me.\r\nEnd Sub\r\n`;
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(source), source, 30_000);
		expect(context?.host).toBe('access');
		expect(context?.meType).toBe('Access.Form');
		expect(context?.meProjectType).toBe(FORM);

		const names = completionsAt(source, 'Probe()\r\n    Me.', toMemberCompletionContext(context!));
		expect(names).toEqual(expect.arrayContaining([
			'RecordSource', 'Requery', 'Filter', 'Controls', // Access.Form
			'Recount', // the module's own code
			'Qty', 'Lines', 'Detail', // the design
		]));
	});

	it('types a control by the design, through Me and on its own', async () => {
		const source = `${BOUND_FORM_SOURCE}\r\nPrivate Sub Probe()\r\n    Me.Lines.\r\n    Qty.\r\n    Me.Detail.\r\nEnd Sub\r\n`;
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(source), source, 30_000);
		const ctx = toMemberCompletionContext(context!);
		expect(completionsAt(source, '    Me.Lines.', ctx)).toEqual(expect.arrayContaining(['AddItem', 'RowSource']));
		expect(completionsAt(source, '    Qty.', ctx)).toEqual(expect.arrayContaining(['SetFocus', 'Value']));
		expect(completionsAt(source, '    Me.Detail.', ctx)).toEqual(expect.arrayContaining(['BackColor', 'Visible']));
	});

	it('describes a form member, a control, and a control s member on hover', async () => {
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(BOUND_FORM_SOURCE), BOUND_FORM_SOURCE, 30_000);
		// The hover provider's own context: the member context plus the module.
		const ctx = { ...toMemberCompletionContext(context!), moduleName: context!.moduleName, moduleKind: context!.moduleKind };
		const hoverAt = (marker: string, word: string) => {
			const offset = BOUND_FORM_SOURCE.indexOf(marker) + marker.indexOf(word) + 1;
			return resolveHover(BOUND_FORM_SOURCE, offset, ctx);
		};
		// Through `Me` the owner shown is the form itself, as it is for a
		// sheet's `Me.Calculate`; what matters is that it resolves at all.
		expect(hoverAt('Me.Requery', 'Requery')?.signature).toContain('Requery');
		expect(hoverAt('Me.Requery', 'Requery')?.details.join(' ')).toContain('method');
		// A control is a member of the form through `Me`, and itself when bare.
		expect(hoverAt('Me.Qty = 5', 'Qty')?.signature).toBe('Form_Calculator.Qty As Textbox');
		expect(hoverAt('    Qty = 6', 'Qty')?.signature).toBe('Qty As Access.Textbox');
		const rowSource = hoverAt('Me.Lines.RowSource', 'RowSource');
		expect(rowSource?.signature).toContain('ListBox.RowSource');
		expect(rowSource?.details.join(' ')).toContain('Access host property');
	});

	it('keeps the controls once the form s code is open and edited', async () => {
		// Folding an open editor into the shared index passed the text alone,
		// which made the index forget what the designer declares: a form s
		// controls left `Me.` the moment its code was open. True of a UserForm
		// too, which is what the second half pins.
		const { contexts, projects } = services();
		const document = openForm(BOUND_FORM_SOURCE) as unknown as { version: number; getText: () => string };
		await contexts.buildEditorProjectContextWithin(document as unknown as vscode.TextDocument, BOUND_FORM_SOURCE, 30_000);
		document.version = 2;
		const edited = `${BOUND_FORM_SOURCE}\r\n' edited\r\n`;
		document.getText = () => edited;
		const live = await projects.contextForProject(DATABASE, 'live');
		expect(live.project.moduleImplicitMembers(FORM).map((member) => member.name)).toContain('Qty');
		expect(live.byModule.get(FORM.toLowerCase())?.designerClass).toBe('Access.Form');
	});

	it('keeps a UserForm s controls and default instance through a live edit as well', async () => {
		const bridge = {
			call: vi.fn(async () => [{
				name: 'EntryForm',
				type: 'userform',
				source: 'Option Explicit\r\n',
				implicitMembers: [{ name: 'NameBox', type: 'MSForms.TextBox' }],
				predeclaredId: true,
			}]),
		} as unknown as ProjectEngine;
		const projects = new VbaProjectIndexService(new VbaSymbolIndex(bridge));
		const book = process.platform === 'win32' ? 'C:/Book.xlsm' : '/work/Book.xlsm';
		const bookUri = book.startsWith('/') ? book : `/${book}`;
		await projects.contextForProject(book, 'live');
		const value = `xlide-vba:${bookUri}/EntryForm.bas`;
		(vscode.workspace.textDocuments as unknown[]).push({
			uri: { scheme: 'xlide-vba', path: `${bookUri}/EntryForm.bas`, toString: () => value },
			version: 1,
			getText: () => "Option Explicit\r\n' edited\r\n",
		});
		const live = await projects.contextForProject(book, 'live');
		expect(live.project.moduleImplicitMembers('EntryForm')).toEqual([{ name: 'NameBox', type: 'MSForms.TextBox' }]);
		expect(live.project.modulePredeclaredId('EntryForm')).toBe(true);
	});

	it('offers the form s, its sections and its controls handlers, and none of a UserForm s', async () => {
		const source = 'Option Explicit\r\n\r\n';
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(source), source, 30_000);
		const stubs = resolveEventHandlerCompletions(source, source.length, toEventHandlerCompletionContext(context!));
		const signatures = stubs.map((stub) => stub.signature);
		// The form s own, under Form_ whatever the form is called.
		expect(signatures).toEqual(expect.arrayContaining(['Form_Load()', 'Form_Open(Cancel As Integer)', 'Form_Current()']));
		// ByVal is part of the declaration: VBA refuses the handler without it.
		expect(signatures).toContain('Form_Unload(Cancel As Integer)');
		expect(signatures).toContain('Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)');
		// A section s, and each control s under the control s name.
		expect(signatures).toEqual(expect.arrayContaining([
			'Detail_Click()',
			'AddLine_Click()',
			'Qty_AfterUpdate()',
			'Qty_KeyDown(KeyCode As Integer, Shift As Integer)',
			'Lines_DblClick(Cancel As Integer)',
			'Express_AfterUpdate()',
			'TotalLabel_Click()',
		]));
		// A label has no AfterUpdate, and nothing here is a UserForm s.
		expect(signatures).not.toContain('TotalLabel_AfterUpdate()');
		expect(stubs.filter((stub) => /^UserForm_/.test(stub.name))).toEqual([]);

		const load = stubs.find((stub) => stub.name === 'Form_Load');
		expect(load?.insertText).toBe('Private Sub Form_Load()\n    $0\nEnd Sub');
		expect(load?.detail).toBe('Form event handler');
		expect(load?.documentation).toContain('Occurs when a form is opened');
		expect(stubs.find((stub) => stub.name === 'Qty_AfterUpdate')?.detail).toBe('Textbox event handler');

		// A real UserForm is still offered its own.
		const userForm = resolveEventHandlerCompletions(source, source.length, { moduleName: 'EntryForm', moduleKind: 'userform' });
		expect(userForm.map((stub) => stub.name)).toContain('UserForm_Initialize');
	});

	it('does not offer a handler the form already has', async () => {
		const source = `${form().source ?? ''}\r\n`;
		// The fixture s form handles these two already.
		expect(source).toMatch(/Sub AddLine_Click\(\)/);
		expect(source).toMatch(/Sub Reset_Click\(\)/);
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(source), source, 30_000);
		const names = resolveEventHandlerCompletions(source, source.length, toEventHandlerCompletionContext(context!))
			.map((stub) => stub.name);
		expect(names).not.toContain('AddLine_Click');
		expect(names).not.toContain('Reset_Click');
		// The same controls other events are still on offer.
		expect(names).toEqual(expect.arrayContaining(['AddLine_DblClick', 'Reset_GotFocus', 'Form_Load']));
	});

	it('narrows to what is being typed, and offers nothing inside a procedure', async () => {
		const typed = 'Option Explicit\r\n\r\nPrivate Sub Qty_Af';
		const { contexts } = services();
		const context = await contexts.buildEditorProjectContextWithin(openForm(typed), typed, 30_000);
		const ctx = toEventHandlerCompletionContext(context!);
		const tail = resolveEventHandlerCompletions(typed, typed.length, ctx);
		expect(tail.map((stub) => stub.name)).toEqual(['Qty_AfterUpdate']);
		expect(tail[0].insertText).toBe('Qty_AfterUpdate()\n    $0\nEnd Sub');

		const inside = 'Option Explicit\r\n\r\nPrivate Sub Go()\r\n    \r\nEnd Sub\r\n';
		const offset = inside.indexOf('    \r\n') + 4;
		expect(resolveEventHandlerCompletions(inside, offset, ctx)).toEqual([]);
	});
});

describe('colouring in an Access form', () => {
	const controls = [
		{ name: 'Lines', type: 'Access.ListBox' },
		{ name: 'Qty', type: 'Access.Textbox' },
	];
	const source = [
		'Private Sub Go()',
		'    Me.Requery',
		'    Me.Lines.AddItem "x"',
		'    Qty.SetFocus',
		'    Me.Show',
		'End Sub',
		'',
	].join('\r\n');

	it('paints Access.Form s and the controls methods, resolved against the Access model', () => {
		const painted = collectHostMemberMethodTokens(source, {
			model: getAccessObjectModel(),
			implicitMembers: controls,
			meType: 'Access.Form',
			meProjectType: FORM,
			projectTypes: [{ name: FORM, kind: 'userform' }],
		}).filter((token) => token.tokenType === 'function').map((token) => token.name);
		expect(painted).toEqual(expect.arrayContaining(['Requery', 'AddItem', 'SetFocus']));
		expect(painted).not.toContain('Show');
	});

	it('leaves Me alone in the forms collector, which would paint a UserForm s Show', () => {
		// Controls alone used to mean "this is a UserForm", so `Me.Show` painted
		// as a method in a form that has no Show.
		expect(collectImplicitMemberMethodTokens(source, { implicitMembers: controls, meType: 'Access.Form' })).toEqual([]);
		expect(collectImplicitMemberMethodTokens(source, { implicitMembers: [{ name: 'Pick', type: 'MSForms.ListBox' }] })
			.map((token) => token.name)).toContain('Show');
	});
});

// Access itself is the authority on a handler. scripts/measure-access-event-
// handlers.py had Access 16.0 write one, through Module.CreateEventProc, for
// every event of a form, a report, each kind of section and each kind of
// control; what it wrote is tests/fixtures/access/eventHandlerOracle.json. A
// stub XLIDE offers must be that line exactly: VBA compiles a handler against
// the event, so a missing ByVal or a wrong type is a compile error, not a style.
describe('every handler stub is the line Access writes', () => {
	interface DesignOracle {
		self: Record<string, string>;
		sections: Record<string, Record<string, string>>;
		controls: Record<string, Record<string, string>>;
	}
	const oracle = JSON.parse(
		readFileSync(path.join(__dirname, 'fixtures', 'access', 'eventHandlerOracle.json'), 'utf8'),
	) as Record<'form' | 'report', DesignOracle>;
	const model = getAccessObjectModel();
	const codeOf = new Map(Object.entries(ACCESS_CONTROL_TYPES).map(([code, name]) => [name, Number(code)]));
	// The type code a section has in a design, by where it sits (AcSection 0 to 6).
	const SECTION_CODES = [152, 153, 154, 155, 156, 157, 158];

	/** The engine s member for one object of a design, named as Access named it. */
	function member(kind: 'form' | 'report', code: number, name: string) {
		const design: AccessDesign = {
			header: Buffer.alloc(10),
			trailer: Buffer.alloc(4),
			objects: [
				{ records: [{ id: 1, code: 20, valueType: 0, width: 0, value: Buffer.from('Design', 'utf16le') }] },
				{ marker: 0xff, type: code, records: [{ id: 1, code: 20, valueType: 0, width: 0, value: Buffer.from(name, 'utf16le') }] },
			],
		};
		return accessDesignMembers(design, kind)[0];
	}

	/** The declaration line of every stub offered, by handler name. */
	function offered(kind: 'form' | 'report', implicitMembers: ReturnType<typeof member>[]): Map<string, string> {
		const stubs = resolveEventHandlerCompletions('', 0, {
			moduleKind: 'userform',
			host: 'access',
			model,
			meType: kind === 'form' ? 'Access.Form' : 'Access.Report',
			implicitMembers,
		});
		return new Map(stubs.map((stub) => [stub.name, stub.insertText.split('\n')[0]]));
	}

	/** `Attachment22` out of `Private Sub Attachment22_AfterUpdate()`. */
	function objectNameIn(line: string, event: string): string {
		return line.slice('Private Sub '.length, line.indexOf(`_${event}(`));
	}

	for (const kind of ['form', 'report'] as const) {
		it(`writes a ${kind} s own handlers as Access does`, () => {
			const stubs = offered(kind, []);
			const wrote = Object.values(oracle[kind].self);
			expect(wrote.length).toBe(kind === 'form' ? 50 : 25);
			expect([...stubs.values()].sort()).toEqual([...wrote].sort());
		});

		it(`writes each ${kind} section s handlers as Access does`, () => {
			const sections = Object.entries(oracle[kind].sections);
			expect(sections.length).toBe(kind === 'form' ? 5 : 7);
			for (const [key, events] of sections) {
				const [index, name] = key.split(':');
				const stubs = offered(kind, [member(kind, SECTION_CODES[Number(index)], name)]);
				const mine = [...stubs.entries()].filter(([stub]) => stub.startsWith(`${name}_`)).map(([, line]) => line);
				expect(mine.sort(), key).toEqual(Object.values(events).sort());
			}
		});

		it(`writes each ${kind} control s handlers as Access does`, () => {
			const controls = Object.entries(oracle[kind].controls);
			expect(controls.length).toBe(kind === 'form' ? 26 : 24);
			let compared = 0;
			for (const [key, events] of controls) {
				// `CheckBox in OptionGroup` is a CheckBox to the design, and Access
				// gave it the same events as one standing alone.
				const typeName = key.replace(/ in OptionGroup$/, '');
				const code = codeOf.get(typeName);
				expect(code, key).toBeDefined();
				const lines = Object.entries(events);
				if (lines.length === 0) {
					// A Line and a PageBreak raise nothing, and nothing is offered.
					const silent = offered(kind, [member(kind, code!, 'Silent1')]);
					expect([...silent.keys()].filter((stub) => stub.startsWith('Silent1_')), key).toEqual([]);
					continue;
				}
				const name = objectNameIn(lines[0][1], lines[0][0]);
				const stubs = offered(kind, [member(kind, code!, name)]);
				const mine = [...stubs.entries()].filter(([stub]) => stub.startsWith(`${name}_`)).map(([, line]) => line);
				expect(mine.sort(), key).toEqual(Object.values(events).sort());
				compared += lines.length;
			}
			expect(compared).toBeGreaterThan(200);
		});
	}

	it('gives a report s sections the events a form s do not have', () => {
		const names = (kind: 'form' | 'report', code: number, name: string): string[] =>
			[...offered(kind, [member(kind, code, name)]).keys()].filter((stub) => stub.startsWith(`${name}_`));
		expect(names('report', 152, 'Detail')).toEqual(expect.arrayContaining(['Detail_Format', 'Detail_Print', 'Detail_Retreat']));
		// A page header is printed once a page, so it has nothing to retreat from.
		expect(names('report', 155, 'PageHeaderSection')).toContain('PageHeaderSection_Format');
		expect(names('report', 155, 'PageHeaderSection')).not.toContain('PageHeaderSection_Retreat');
		// A form s sections, its page header included, have none of the three.
		for (const code of [152, 155]) {
			expect(names('form', code, 'Part').filter((stub) => /_(Format|Print|Retreat)$/.test(stub))).toEqual([]);
		}
		// It is still an Access.Section to everything that reads its members.
		expect(member('report', 152, 'Detail')).toEqual({ name: 'Detail', type: 'Access.Section', eventClass: 'Access._SectionInReport' });
		expect(member('report', 156, 'PageFooterSection').eventClass).toBe('Access._PageHdrFtrInReport');
		expect(member('form', 155, 'PageHeaderSection')).toEqual({ name: 'PageHeaderSection', type: 'Access.Section' });
	});

	// The file the model s events are generated from is part of the reference
	// corpus, which is not committed; so, as tests/vbaHostConstantRoundTrip
	// does, this runs where the corpus is - the machines that can regenerate -
	// and the oracle above is what holds the committed model everywhere else.
	const EVENT_SOURCES = path.join(__dirname, '..', 'reference', 'access', 'eventSources.json');

	it.runIf(existsSync(EVENT_SOURCES))('carries in the model exactly the events the type library lists', () => {
		// The model is generated; this holds a stale one to the file it came from.
		const sources = JSON.parse(readFileSync(EVENT_SOURCES, 'utf8')) as {
			classes: Record<string, string>;
			interfaces: Record<string, { name: string; params: { name: string; type: string; byVal?: boolean }[] }[]>;
		};
		const typeKeys = new Map(Object.keys(model.types).map((key) => [key.toLowerCase(), key]));
		let held = 0;
		for (const [className, interfaceName] of Object.entries(sources.classes)) {
			const key = typeKeys.get(`access.${className.toLowerCase()}`);
			if (!key) {
				continue;
			}
			const expected = sources.interfaces[interfaceName].map((event) => `${event.name}(${event.params
				.map((param) => `${param.byVal ? 'ByVal ' : ''}${param.name} As ${param.type}`).join(', ')})`);
			const carried = (model.types[key].members ?? [])
				.filter((entry) => entry.kind === 'event').map((entry) => entry.signature);
			expect(carried.sort(), key).toEqual(expected.sort());
			held += 1;
		}
		// TextBox, CheckBox and ComboBox are spelled Textbox, Checkbox and
		// Combobox in the model; matching by case alone finds all three.
		expect(held).toBe(44);
		for (const key of ['Access.Textbox', 'Access.Checkbox', 'Access.Combobox']) {
			expect((model.types[key].members ?? []).some((entry) => entry.kind === 'event'), key).toBe(true);
		}
	});

	it('offers nothing it cannot vouch for', () => {
		const stubsFor = (ctx: Parameters<typeof resolveEventHandlerCompletions>[2]): string[] =>
			resolveEventHandlerCompletions('', 0, ctx).map((stub) => stub.name);
		const base = { moduleKind: 'userform' as const, host: 'access', meType: 'Access.Form' };
		// No model, no stubs - and never the UserForm table in its place.
		expect(stubsFor(base)).toEqual([]);
		// A control of a class the model does not carry.
		const unknown = stubsFor({ ...base, model, implicitMembers: [{ name: 'Gauge', type: 'Access.NoSuchControl' }] });
		expect(unknown.filter((stub) => stub.startsWith('Gauge_'))).toEqual([]);
		// An event with no parameter list: its ByVal cannot be known, so it is left out.
		const bare = {
			...model,
			types: {
				...model.types,
				'Access.Bare': { displayName: 'Bare', members: [
					{ name: 'Known', kind: 'event' as const, signature: 'Known(ByVal Flag As Boolean)' },
					{ name: 'Unknown', kind: 'event' as const },
				] },
			},
		};
		expect(stubsFor({ ...base, model: bare, implicitMembers: [{ name: 'Dial', type: 'Access.Bare' }] })
			.filter((stub) => stub.startsWith('Dial_'))).toEqual(['Dial_Known']);
	});
});

describe('the name VBA knows a control by', () => {
	it('converts a name the way Access recorded it', () => {
		// Each pair is what Access 16.0 stored in a form s TypeInfo stream beside
		// the name shown, and what Module.CreateEventProc would accept.
		for (const [shown, identifier] of [
			['Plain', 'Plain'],
			['Order Date', 'Order_Date'],
			['Qty-1', 'Qty_1'],
			['Total$', 'Total_'],
			['Tax (VAT)', 'Tax__VAT_'],
			['Dbl  space', 'Dbl__space'],
			['trail ', 'trail_'],
			['a.b', 'a_b'],
			['2ndBox', 'Ctl2ndBox'],
			['9', 'Ctl9'],
			['_under', 'Ctl_under'],
			[' lead', 'Ctl_lead'],
			['Under_score', 'Under_score'],
			// A reserved word is left alone, and so is anything outside ASCII.
			['Me', 'Me'],
			['café', 'café'],
			['x²', 'x²'],
			['€uro', '€uro'],
			['a–b', 'a–b'],
			['½x', '½x'],
			['naïve one', 'naïve_one'],
		]) {
			expect(accessVbaIdentifier(shown), JSON.stringify(shown)).toBe(identifier);
		}
	});

	it('names a member, and so its handlers, by that identifier', () => {
		const named = (name: string): AccessDesign['objects'][number]['records'] => [
			{ id: 1, code: 20, valueType: 0, width: 0, value: Buffer.from(name, 'utf16le') },
		];
		const design: AccessDesign = {
			header: Buffer.alloc(10),
			trailer: Buffer.alloc(4),
			objects: [
				{ records: named('Orders') },
				{ marker: 0xff, type: 109, records: named('Order Date') },
				{ marker: 0xff, type: 104, records: named('2ndBox') },
			],
		};
		const members = accessDesignMembers(design);
		// `Me.Order_Date` compiles; `Me.Order Date` is not VBA.
		expect(members).toEqual([
			{ name: 'Order_Date', type: 'Access.Textbox' },
			{ name: 'Ctl2ndBox', type: 'Access.CommandButton' },
		]);
		const stubs = resolveEventHandlerCompletions('', 0, {
			moduleKind: 'userform', host: 'access', model: getAccessObjectModel(), meType: 'Access.Form', implicitMembers: members,
		}).map((stub) => stub.signature);
		expect(stubs).toEqual(expect.arrayContaining(['Order_Date_AfterUpdate()', 'Ctl2ndBox_Click()']));
	});
});
