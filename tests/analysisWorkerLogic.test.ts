import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../src/vbaProjectAnalysis';

const MOD_A = [
	'Option Explicit',
	'Sub Alpha()',
	'    undeclaredThing = 1',
	'End Sub',
].join('\n');

const MOD_B = [
	'Option Explicit',
	'Public Sub Helper(ByVal n As Long)',
	'End Sub',
].join('\n');

describe('AnalysisWorkerState', () => {
	it('analyzes with seeded project context and matches the in-host pass', () => {
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed',
			workbookKey: 'wb1',
			generation: 1,
			modules: [
				{ moduleName: 'ModA', source: MOD_A, type: 'standard' },
				{ moduleName: 'ModB', source: MOD_B, type: 'standard' },
			],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: MOD_A, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(response.diagnostics.some((d) => d.message.includes('undeclaredThing'))).toBe(true);
		// Second analyze with a body-only change engages incremental.
		const edited = MOD_A.replace('    undeclaredThing = 1', '    undeclaredThing = 1\n    alsoUndeclared = 2');
		const second = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: edited, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(second?.kind).toBe('result');
		if (second?.kind !== 'result') { return; }
		expect(second.incrementalMode).toBe('incremental');
		expect(second.diagnostics.some((d) => d.message.includes('alsoUndeclared'))).toBe(true);
	});

	it('requests a reseed for an unknown or stale workbook generation', () => {
		const state = new AnalysisWorkerState();
		const unknown = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'nope', generation: 1,
			source: MOD_A, moduleName: 'ModA',
		});
		expect(unknown?.kind).toBe('needSeed');
		state.handle({ kind: 'seed', workbookKey: 'wb1', generation: 1, modules: [{ moduleName: 'ModA', source: MOD_A }] });
		const stale = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 2,
			source: MOD_A, moduleName: 'ModA',
		});
		expect(stale?.kind).toBe('needSeed');
	});

	it('analyzes standalone (no workbook) and matches direct analysis', () => {
		const state = new AnalysisWorkerState();
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1',
			source: MOD_A, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		const direct = analyzeVbaModuleSource({ source: MOD_A, moduleName: 'ModA', moduleType: 'standard' });
		expect(response.diagnostics.map((d) => `${d.code}:${d.span.start}`).sort())
			.toEqual(direct.diagnostics.map((d) => `${d.code}:${d.span.start}`).sort());
	});

	it('forget clears per-document incremental state', () => {
		const state = new AnalysisWorkerState();
		state.handle({ kind: 'seed', workbookKey: 'wb1', generation: 1, modules: [{ moduleName: 'ModA', source: MOD_A }] });
		state.handle({ kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'wb1', generation: 1, source: MOD_A, moduleName: 'ModA' });
		state.handle({ kind: 'forget', docKey: 'doc1' });
		const after = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: MOD_A.replace('= 1', '= 2'), moduleName: 'ModA',
		});
		expect(after?.kind).toBe('result');
		if (after?.kind !== 'result') { return; }
		expect(after.incrementalMode).toBe('full');
	});
});

describe('AnalysisWorkerState host-supplied implicit members', () => {
	// A form's controls are declared by the designer, and a host that seeds
	// CodeModule text (xlide_vbide reads the VBE, not a .frm file) sends source
	// with no designer header for the worker to read them out of. Without a way
	// across the boundary every control reference reads as undeclared.
	const FORM_SOURCE = [
		'Option Explicit',
		'Private Sub UserForm_Initialize()',
		'    RegionPick.AddItem "West"',
		'End Sub',
	].join('\n');
	const CONTROLS = [{ name: 'RegionPick', type: 'MSForms.ComboBox' }];
	const undeclaredNames = (diagnostics: readonly { code: string; message: string }[]): string[] =>
		diagnostics.filter((d) => d.code === 'undeclared-variable').map((d) => d.message);

	it('reports a control as undeclared when nothing supplies it', () => {
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-form', generation: 1,
			modules: [{ moduleName: 'FrmPicker', source: FORM_SOURCE, type: 'userform' }],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-form', workbookKey: 'wb-form', generation: 1,
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(undeclaredNames(response.diagnostics).some((m) => m.includes('RegionPick'))).toBe(true);
	});

	it('takes implicit members from the seed and matches the in-host pass', () => {
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-form', generation: 1,
			modules: [{
				moduleName: 'FrmPicker', source: FORM_SOURCE, type: 'userform',
				implicitMembers: CONTROLS,
			}],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-form', workbookKey: 'wb-form', generation: 1,
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(undeclaredNames(response.diagnostics)).toEqual([]);

		const project = buildVbaProjectIndex([{ moduleName: 'FrmPicker', source: FORM_SOURCE, type: 'userform' }]);
		const inHost = analyzeVbaModuleSource({
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
			...projectAnalysisOptionsForModule(project, 'FrmPicker', projectProcedureSignatures(project)),
			implicitMembers: CONTROLS,
		});
		expect(response.diagnostics).toEqual(inHost.diagnostics);
	});

	it('lets the analyze request override what the seed said', () => {
		// The designer is edited far more often than a module is re-seeded, so
		// the request is what a host with live designer state sends.
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-form', generation: 1,
			modules: [{
				moduleName: 'FrmPicker', source: FORM_SOURCE, type: 'userform',
				implicitMembers: [{ name: 'StaleName', type: 'MSForms.ComboBox' }],
			}],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-form', workbookKey: 'wb-form', generation: 1,
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
			implicitMembers: CONTROLS,
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(undeclaredNames(response.diagnostics)).toEqual([]);
	});

	it('re-analyzes when the request changes the controls', () => {
		// A renamed control changes diagnostics without changing a line of
		// code, so the control list has to reach the incremental fingerprint -
		// otherwise the first answer sticks for the rest of the session.
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-form', generation: 1,
			modules: [{ moduleName: 'FrmPicker', source: FORM_SOURCE, type: 'userform' }],
		});
		const first = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-form', workbookKey: 'wb-form', generation: 1,
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
			implicitMembers: CONTROLS,
		});
		expect(first?.kind === 'result' ? undeclaredNames(first.diagnostics) : undefined).toEqual([]);
		const renamed = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc-form', workbookKey: 'wb-form', generation: 1,
			source: FORM_SOURCE, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
			implicitMembers: [{ name: 'RegionPicker', type: 'MSForms.ComboBox' }],
		});
		expect(renamed?.kind).toBe('result');
		if (renamed?.kind !== 'result') { return; }
		expect(undeclaredNames(renamed.diagnostics).some((m) => m.includes('RegionPick'))).toBe(true);
	});

	it('keeps the designer-header parse when the host supplies nothing', () => {
		// A standalone .frm still carries its control tree in its own text, and
		// the new field must not displace that.
		const withHeader = [
			'VERSION 5.00',
			'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPicker ',
			'   Caption         =   "Pick"',
			'   Begin Forms.ComboBox.1 RegionPick ',
			'      Height          =   24',
			'   End',
			'End',
			"Attribute VB_Name = \"FrmPicker\"",
			...FORM_SOURCE.split('\n'),
		].join('\n');
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-frm', generation: 1,
			modules: [{ moduleName: 'FrmPicker', source: withHeader, type: 'userform' }],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-frm', workbookKey: 'wb-frm', generation: 1,
			source: withHeader, moduleName: 'FrmPicker', moduleType: 'userform', moduleKind: 'userform',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(undeclaredNames(response.diagnostics)).toEqual([]);
	});
});

describe('AnalysisWorkerState suppressed diagnostics', () => {
	it('returns suppressed findings alongside active ones', () => {
		// Analyze Workbook reports suppressed problems in their own panel
		// section, so the worker result must carry them across the boundary
		// rather than dropping them like the live-diagnostics path can.
		const source = [
			'Option Explicit',
			'Sub Alpha()',
			"    ' @xlide-analysis-disable-next-line undeclared-variable",
			'    silenced = 1',
			'    loudAndClear = 2',
			'End Sub',
		].join('\n');
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', workbookKey: 'wb-supp', generation: 1,
			modules: [{ moduleName: 'ModS', source, type: 'standard' }],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc-supp', workbookKey: 'wb-supp', generation: 1,
			source, moduleName: 'ModS', moduleType: 'standard',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(response.diagnostics.some((d) => d.message.includes('loudAndClear'))).toBe(true);
		expect(response.diagnostics.some((d) => d.message.includes('silenced'))).toBe(false);
		expect(response.suppressedDiagnostics.some((d) => d.message.includes('silenced'))).toBe(true);

		// Byte-for-byte the same partition the in-host pass produces.
		const project = buildVbaProjectIndex([{ moduleName: 'ModS', source, type: 'standard' }]);
		const inHost = analyzeVbaModuleSource({
			source, moduleName: 'ModS', moduleType: 'standard',
			...projectAnalysisOptionsForModule(project, 'ModS', projectProcedureSignatures(project)),
		});
		expect(response.diagnostics).toEqual(inHost.diagnostics);
		expect(response.suppressedDiagnostics).toEqual(inHost.suppressedDiagnostics);
	});
});
