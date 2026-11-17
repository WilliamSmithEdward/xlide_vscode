// Diagnostics tests: a statement that calls one of the module's controls, as
// `CommandButton1`, `Call CommandButton1` or `CommandButton1 1` (issue #1555).
// Measured in Excel, Word, PowerPoint and Access 16.0 (build 20527) through
// the VBE oracle, with ActiveX enabled: a control on a sheet, document or
// slide gives "Expected procedure, not variable"; a UserForm's control, and
// an Access form's or report's control or section, "Invalid use of property".

import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { analyzeModule } from '../../../src/analyzer';
import { readModulesFromBuffer } from '../../../src/vba/projectService';
import { moduleKindFromType } from '../../../src/vbaProjectAnalysis';

type Options = Parameters<typeof analyzeModule>[1];

function reported(body: string, options: Options): string[] {
	const src = `Option Explicit\nPublic Sub Probe()\n    ${body}\nEnd Sub\n`;
	// The call rule runs only with the project's known names, as in a project.
	const diags = analyzeModule(src, { knownIdentifiers: new Set(), knownProcedures: new Set(), ...options });
	return diags.filter((d) => d.severity === 'error').map((d) => `${d.code} ${src.slice(d.span.start, d.span.end)}: ${d.message}`);
}

const button = [{ name: 'CommandButton1', type: 'Object' }];
const sheet: Options = { moduleName: 'Sheet1', moduleKind: 'document', documentType: 'worksheet', designerClass: 'Excel.Worksheet', host: 'excel', implicitMembers: button };
const doc: Options = { moduleName: 'ThisDocument', moduleKind: 'document', designerClass: 'Word.Document', host: 'word', implicitMembers: button };
const slide: Options = { moduleName: 'Slide1', moduleKind: 'document', designerClass: 'PowerPoint.Slide', host: 'powerpoint', implicitMembers: button };
const form: Options = { moduleName: 'Form1', moduleKind: 'userform', host: 'excel', implicitMembers: [{ name: 'CommandButton1', type: 'MSForms.CommandButton' }] };
const accessForm: Options = {
	moduleName: 'Form_Frm1', moduleKind: 'userform', designerClass: 'Access.Form', host: 'access',
	implicitMembers: [{ name: 'Detail', type: 'Access.Section' }, { name: 'CommandButton1', type: 'Access.CommandButton' }],
};

const calls = ['CommandButton1', 'Call CommandButton1', 'CommandButton1 1', 'Call CommandButton1(1)'];
const variable = "non-callable-call CommandButton1: Cannot call 'CommandButton1' because it is a control, not a Sub or Function. This is a VBE compile error: Expected procedure, not variable.";
const property = "invalid-property-use CommandButton1: 'CommandButton1' is a control, and a statement cannot call one. This is a VBE compile error: Invalid use of property.";

describe('a statement that calls a control (issue #1555)', () => {
	it.each([['sheet', sheet], ['document', doc], ['slide', slide]] as const)('is "Expected procedure, not variable" on a %s', (_, options) => {
		for (const body of calls) {
			expect(reported(body, options), body).toEqual([variable]);
		}
	});

	it.each([['a UserForm', form], ['an Access form', accessForm]] as const)('is "Invalid use of property" on %s', (_, options) => {
		for (const body of calls) {
			expect(reported(body, options), body).toEqual([property]);
		}
	});

	it("names an Access section as a section", () => {
		expect(reported('Detail', accessForm)).toEqual([
			"invalid-property-use Detail: 'Detail' is a section, and a statement cannot call one. This is a VBE compile error: Invalid use of property.",
		]);
		expect(reported('Detail', { ...accessForm, moduleName: 'Report_Rpt1', designerClass: 'Access.Report' })).toEqual([
			"invalid-property-use Detail: 'Detail' is a section, and a statement cannot call one. This is a VBE compile error: Invalid use of property.",
		]);
	});

	it.each([['a sheet', sheet], ['a document', doc], ['a slide', slide], ['a UserForm', form], ['an Access form', accessForm]] as const)('leaves a call of a member of the control alone on %s', (_, options) => {
		expect(reported('CommandButton1.SetFocus', options)).toEqual([]);
		expect(reported('Call CommandButton1.SetFocus', options)).toEqual([]);
	});

	it('still reports a name that is no control as not defined', () => {
		expect(reported('Nope', sheet)).toEqual(["unknown-call Nope: Sub or Function not defined: 'Nope'."]);
		expect(reported('Nope', form)).toEqual(["unknown-call Nope: Sub or Function not defined: 'Nope'."]);
	});

	it('leaves the name to what declares it, when the procedure or the module does', () => {
		expect(reported('Dim CommandButton1 As Long\n    CommandButton1', sheet).map((line) => line.split(':')[0])).toEqual(['non-callable-call CommandButton1']);
		const src = 'Option Explicit\nPublic Sub Probe()\n    CommandButton1\nEnd Sub\nPrivate Sub CommandButton1()\nEnd Sub\n';
		const diags = analyzeModule(src, { knownIdentifiers: new Set(), knownProcedures: new Set(['probe', 'commandbutton1']), ...sheet });
		expect(diags.filter((d) => d.code === 'non-callable-call' || d.code === 'unknown-call')).toEqual([]);
		for (const options of [sheet, form]) {
			const withEnum = 'Option Explicit\nPrivate Enum Kinds\n    CommandButton1 = 1\nEnd Enum\nPublic Sub Probe()\n    Call CommandButton1\nEnd Sub\n';
			const found = analyzeModule(withEnum, { knownIdentifiers: new Set(), knownProcedures: new Set(['probe']), ...options }).filter((d) => d.severity === 'error');
			expect(found.map((d) => d.message)).toEqual(["Cannot call 'CommandButton1' because it resolves to an enum member, not a Sub or Function."]);
		}
	});

	it('keeps the control over a procedure of that name elsewhere in the project', () => {
		expect(reported('CommandButton1', { ...sheet, knownProcedures: new Set(['commandbutton1']) })).toEqual([variable]);
		expect(reported('Call CommandButton1', { ...form, knownProcedures: new Set(['commandbutton1']) })).toEqual([property]);
	});

	it("keeps the control over VBA's and the host's own names", () => {
		for (const name of ['Beep', 'Calculate']) {
			expect(reported(name, { ...form, implicitMembers: [{ name, type: 'MSForms.CommandButton' }] })).toEqual([
				`invalid-property-use ${name}: '${name}' is a control, and a statement cannot call one. This is a VBE compile error: Invalid use of property.`,
			]);
		}
	});

	it('gives an Access form its message whichever module kind it arrives as', () => {
		expect(reported('CommandButton1', { ...accessForm, moduleKind: 'document' })).toEqual([property]);
	});

	it("keeps VB6's forms as they were, since they were not measured", () => {
		expect(reported('CommandButton1', { ...form, host: 'vb6', designerClass: 'VB.Form' })).toEqual(["unknown-call CommandButton1: Sub or Function not defined: 'CommandButton1'."]);
	});

	it.each([
		['WordControlsFixture.docm', 'word', 'ThisDocument', 'non-callable-call'],
		['PowerPointControlsFixture.pptm', 'powerpoint', 'Slide1', 'non-callable-call'],
		['FormFixture.xlsm', 'excel', 'FrmPicker', 'invalid-property-use'],
		['AccessControlNamesFixture.accdb', 'access', 'Form_Names', 'invalid-property-use'],
	])('reaches the extension through the analysis worker, for %s', (file, host, moduleName, code) => {
		const entries = readModulesFromBuffer(fs.readFileSync(path.join(__dirname, '../../fixtures/binaries', file)), true);
		const mod = entries.find((entry) => entry.name === moduleName)!;
		const control = mod.implicitMembers![0].name;
		const source = `Option Explicit\r\nPublic Sub Probe()\r\n    Call ${control}\r\n    ${control}.SetFocus\r\nEnd Sub\r\n`;
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed', projectKey: 'p', generation: 1,
			modules: entries.map((entry) => ({ moduleName: entry.name, source: entry === mod ? source : entry.source ?? '', type: entry.type, documentType: entry.documentType, designerClass: entry.designerClass, implicitMembers: entry.implicitMembers })),
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'm', projectKey: 'p', generation: 1, host,
			source, moduleName: mod.name, moduleType: mod.type, moduleKind: moduleKindFromType(mod.type), documentType: mod.documentType, designerClass: mod.designerClass,
		});
		const codes = response?.kind === 'result' ? response.diagnostics.filter((d) => d.severity === 'error').map((d) => `${d.code} ${source.slice(d.span.start, d.span.end)}`) : ['no result'];
		expect(codes).toEqual([`${code} ${control}`]);
	});
});
