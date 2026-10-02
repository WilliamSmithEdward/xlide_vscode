// Diagnostics tests: a VB6 form's own VB.Form members reached through the
// form's name, its default instance, or a variable typed as the form (issue
// #358). No VB6 compiler is at hand; the evidence is VB.Form in the vb6 model
// and upstream's Diabetes-prediction project, whose MCD.frm draws with
// Form1.Cls, Form1.CurrentX and Form1.Print.

import { describe, it, expect } from 'vitest';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule, type ProjectTestModule } from './helpers';
import { readVb6Modules } from '../../src/vba/vb6/vb6Project';
import * as path from 'path';

const FORM = 'Option Explicit\nPrivate Sub Form_Load()\nEnd Sub\n';

function findings(currentModule: string, body: string): string[] {
	const modules: ProjectTestModule[] = [
		{ moduleName: 'Form1', source: FORM, moduleKind: 'userform', implicitMembers: [{ name: 'Text1', type: 'VB.TextBox' }], predeclaredId: true, designerClass: 'VB.Form' },
		{ moduleName: 'Module1', source: 'Option Explicit\n' },
	];
	const src = `Option Explicit\nPublic Sub Draw()\n${body}\nEnd Sub\n`;
	const extra = currentModule === 'Form1' ? { host: 'vb6', moduleKind: 'userform', designerClass: 'VB.Form', implicitMembers: [{ name: 'Text1', type: 'VB.TextBox' }] } : { host: 'vb6' };
	const diagnostics = analyzeProjectModule(src, modules, currentModule, extra as never);
	return ['member-not-found', 'argument-count'].flatMap((code) => byCode(diagnostics, code).map((d) => `${code}: ${d.message}`));
}

describe("a VB6 form's own members", () => {
	const lines = [
		'    Form1.Cls',
		'    Form1.CurrentX = 10',
		'    Form1.CurrentY = 20',
		'    Form1.Print "S" & 1',
		'    Form1.Print "a"; 2',
		'    Form1.Line (1, 2)-(3, 4)',
		'    Debug.Print Form1.ScaleWidth',
		'    Form1.Caption = "x"',
		'    Form1.Text1.Text = "y"',
	];

	it('are reached through the form name, from the form and from a module', () => {
		for (const module of ['Form1', 'Module1']) {
			expect(findings(module, lines.join('\n')), module).toEqual([]);
		}
	});

	it('are reached through a variable typed as the form', () => {
		const body = ['    Dim f As Form1', '    Set f = New Form1', '    f.Cls', '    f.CurrentX = 1', '    f.Print "x"'].join('\n');
		expect(findings('Module1', body)).toEqual([]);
	});

	it('take Print with an argument list through Me', () => {
		expect(findings('Form1', ['    Me.Cls', '    Me.Print "me"', '    Me.Print "a"; 1; "b"', '    Me.Line (1, 2)-(3, 4)'].join('\n'))).toEqual([]);
	});

	it('still reports a member VB.Form does not have', () => {
		expect(findings('Module1', '    Form1.NoSuchMember')).toEqual([expect.stringContaining('member-not-found')]);
	});

	it('draw nothing in upstream MCD.frm, analyzed under its own name', () => {
		const modules = readVb6Modules(path.join(__dirname, '..', 'fixtures', 'vb6', 'Diabetes-prediction-1.0', 'MCD_prj.vbp'))
			.filter((m) => m.source !== undefined)
			.map((m) => ({
				moduleName: m.name,
				source: m.source!,
				moduleKind: m.type === 'userform' ? 'userform' as const : m.type === 'class' ? 'class' as const : 'standard' as const,
				implicitMembers: m.implicitMembers,
				designerClass: m.designerClass,
				predeclaredId: m.predeclaredId,
			}));
		const form = modules.find((m) => m.moduleName === 'Form1')!;
		const diagnostics = analyzeProjectModule(form.source, modules, 'Form1', {
			host: 'vb6', moduleKind: 'userform', designerClass: form.designerClass, implicitMembers: form.implicitMembers,
		} as never);
		expect(['member-not-found', 'argument-count'].flatMap((code) => byCode(diagnostics, code).map((d) => `${code} L${d.line}: ${d.message}`))).toEqual([]);
	});
});
