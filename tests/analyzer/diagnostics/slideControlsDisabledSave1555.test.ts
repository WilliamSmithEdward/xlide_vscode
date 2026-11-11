// Issue #1555: ActiveX controls on the slides of a presentation saved after a
// VBA edit while ActiveX was disabled. That save drops each slide module's
// `Attribute VB_Control` lines, but the controls stay on the slides, and
// XLIDE reads them there. The fixture is a PowerPoint 16.0 file (build 20527)
// with four slides, moved, deleted and added so that their IDs are out of
// order, and one control on each: ThirdBtn on the slide whose module is
// Slide3, and so on. Reopened with ActiveX enabled, the VBE compiled each
// control bare, through Me and through the slide's name, and refused
// `Me.Nope`, a misspelt control and another slide's control, in this
// presentation and in one made the same way in Office 2016 (build 12527),
// each saved as .pptm, .potm and .ppsm.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { readModulesFromBuffer } from '../../../src/vba/projectService';
import { moduleKindFromType } from '../../../src/vbaProjectAnalysis';
import { ZipArchive } from '../../../src/vba/zip';

const FILE = path.join(__dirname, '..', '..', 'fixtures', 'binaries', 'PowerPointControlsDisabledSaveFixture.pptm');
const CONTROLS: Record<string, string> = { Slide1: 'FirstBtn', Slide3: 'ThirdBtn', Slide4: 'LastBtn', Slide5: 'NewBtn' };

function entries(data = fs.readFileSync(FILE)) {
	return readModulesFromBuffer(data, true);
}

/** The errors the analysis worker reports for one probe line, in the extension's way. */
function errors(moduleName: string, statement: string): string[] {
	const all = entries();
	const mod = all.find((entry) => entry.name === moduleName)!;
	const source = `Option Explicit\r\nPublic Sub Probe()\r\n    ${statement}\r\nEnd Sub\r\n`;
	const state = new AnalysisWorkerState();
	state.handle({
		kind: 'seed', projectKey: 'p', generation: 1,
		modules: all.map((entry) => ({ moduleName: entry.name, source: entry === mod ? source : entry.source ?? '', type: entry.type, documentType: entry.documentType, designerClass: entry.designerClass, implicitMembers: entry.implicitMembers })),
	});
	const response = state.handle({
		kind: 'analyze', requestId: 1, docKey: 'm', projectKey: 'p', generation: 1, host: 'powerpoint',
		source, moduleName: mod.name, moduleType: mod.type, moduleKind: moduleKindFromType(mod.type), documentType: mod.documentType, designerClass: mod.designerClass,
	});
	return response?.kind === 'result' ? response.diagnostics.filter((d) => d.severity === 'error').map((d) => d.code) : ['no result'];
}

describe('the controls on slides saved with ActiveX disabled (issue #1555)', () => {
	it('has lost the header lines', () => {
		for (const name of Object.keys(CONTROLS)) {
			const source = entries().find((entry) => entry.name === name)?.source;
			expect(source, name).toBeDefined();
			expect(source, name).not.toMatch(/VB_Control/i);
		}
	});

	it('gives each slide module the controls of the slide whose ID its name carries', () => {
		const byName = new Map(entries().map((entry) => [entry.name, entry]));
		for (const [name, control] of Object.entries(CONTROLS)) {
			expect(byName.get(name)?.designerClass, name).toBe('PowerPoint.Slide');
			expect(byName.get(name)?.implicitMembers, name).toEqual([{ name: control, type: 'Object' }]);
		}
		expect(byName.has('Module1')).toBe(true);
		expect(byName.get('Module1')?.implicitMembers).toBeUndefined();
	});

	it('accepts what the VBE compiles', () => {
		for (const [name, control] of Object.entries(CONTROLS)) {
			expect(errors(name, `Debug.Print TypeName(${control})`), name).toEqual([]);
			expect(errors(name, `Debug.Print TypeName(Me.${control})`), name).toEqual([]);
			expect(errors('Module1', `Debug.Print TypeName(${name}.${control})`), name).toEqual([]);
		}
	});

	it('reports what the VBE refuses', () => {
		expect(errors('Slide1', 'Debug.Print TypeName(Me.Nope)')).toEqual(['member-not-found']);
		expect(errors('Slide1', 'Debug.Print TypeName(CtlMisspelt)')).toEqual(['undeclared-variable']);
		expect(errors('Slide1', 'Debug.Print TypeName(Me.ThirdBtn)')).toEqual(['member-not-found']);
		expect(errors('Slide3', 'Debug.Print TypeName(Me.FirstBtn)')).toEqual(['member-not-found']);
	});

	it('leaves every slide module as it was when one of them names no slide of the presentation', () => {
		// Slide3's slide gets another ID, so no slide has the ID its name gives.
		const zip = ZipArchive.read(fs.readFileSync(FILE));
		const presentation = zip.read('ppt/presentation.xml').toString('utf8');
		expect(presentation).toContain('<p:sldId id="258"');
		zip.write('ppt/presentation.xml', Buffer.from(presentation.replace('<p:sldId id="258"', '<p:sldId id="270"'), 'utf8'));
		for (const entry of entries(zip.toBytes()).filter((candidate) => candidate.type === 'document')) {
			expect(entry.designerClass, entry.name).toBeUndefined();
			expect(entry.implicitMembers, entry.name).toBeUndefined();
		}
	});
});
