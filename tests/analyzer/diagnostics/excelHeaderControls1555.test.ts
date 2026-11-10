// Issue #1555: an ActiveX control on a sheet of an .xls, .xla, .xlt or .xlsb
// workbook. Excel lists a sheet module's controls in its attribute header,
// one `Attribute VB_Control` line each, as Word and PowerPoint do, and XLIDE
// reads them there. It reads an .xlsm's controls from the sheet's XML part,
// which a legacy workbook doesn't have and an .xlsb keeps in binary. The
// fixtures are Excel 16.0's own files (build 20527), with eleven controls on
// Sheet1 and one on Data. With ActiveX enabled, the VBE compiled every
// control bare, through Me and through the sheet's name, and refused
// `Me.Nope`, a misspelt control and another sheet's control, in these files
// and in .xla and .xlt files built the same way, in builds 20527 and 12527.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { readModulesFromBuffer } from '../../../src/vba/projectService';
import { moduleKindFromType } from '../../../src/vbaProjectAnalysis';

const BINARIES = path.join(__dirname, '..', '..', 'fixtures', 'binaries');
const FILES = ['ExcelControlsFixture.xls', 'ExcelControlsFixture.xlsb'];
const SHEET1 = ['CtlCommandButton', 'CtlTextBox', 'CtlCheckBox', 'CtlOptionButton', 'CtlToggleButton', 'CtlComboBox', 'CtlListBox', 'CtlLabel', 'CtlImage', 'CtlScrollBar', 'CtlSpinButton'];

function entries(file: string) {
	return readModulesFromBuffer(fs.readFileSync(path.join(BINARIES, file)), true);
}

/** The errors the analysis worker reports for one probe line, in the extension's way. */
function errors(file: string, moduleName: string, statement: string): string[] {
	const all = entries(file);
	const mod = all.find((entry) => entry.name === moduleName)!;
	const source = `Option Explicit\r\nPublic Sub Probe()\r\n    ${statement}\r\nEnd Sub\r\n`;
	const state = new AnalysisWorkerState();
	state.handle({
		kind: 'seed', projectKey: 'p', generation: 1,
		modules: all.map((entry) => ({ moduleName: entry.name, source: entry === mod ? source : entry.source ?? '', type: entry.type, documentType: entry.documentType, designerClass: entry.designerClass, implicitMembers: entry.implicitMembers })),
	});
	const response = state.handle({
		kind: 'analyze', requestId: 1, docKey: 'm', projectKey: 'p', generation: 1, host: 'excel',
		source, moduleName: mod.name, moduleType: mod.type, moduleKind: moduleKindFromType(mod.type), documentType: mod.documentType, designerClass: mod.designerClass,
	});
	return response?.kind === 'result' ? response.diagnostics.filter((d) => d.severity === 'error').map((d) => d.code) : ['no result'];
}

describe.each(FILES)('the controls on a sheet of %s (issue #1555)', (file) => {
	it('gives each sheet module the controls its header lists, as a Worksheet', () => {
		const byName = new Map(entries(file).map((entry) => [entry.name, entry]));
		expect(byName.get('Sheet1')?.designerClass).toBe('Excel.Worksheet');
		// The order is not compared: an .xlsb's binary sheet part, which XLIDE
		// may read instead of the header, lists them newest first.
		const sorted = (members: readonly { name: string }[] | undefined) => [...(members ?? [])].sort((a, b) => a.name.localeCompare(b.name));
		expect(sorted(byName.get('Sheet1')?.implicitMembers)).toEqual(sorted(SHEET1.map((name) => ({ name, type: 'Object' }))));
		expect(byName.get('Data')?.designerClass).toBe('Excel.Worksheet');
		expect(byName.get('Data')?.implicitMembers).toEqual([{ name: 'DataBtn', type: 'Object' }]);
		for (const name of ['ThisWorkbook', 'Chart3', 'Module1']) {
			expect(byName.has(name), name).toBe(true);
			expect(byName.get(name)?.designerClass, name).toBeUndefined();
			expect(byName.get(name)?.implicitMembers, name).toBeUndefined();
		}
	});

	it('accepts what the VBE compiles', () => {
		for (const control of SHEET1) {
			expect(errors(file, 'Sheet1', `Debug.Print TypeName(${control})`), control).toEqual([]);
			expect(errors(file, 'Sheet1', `Debug.Print TypeName(Me.${control})`), control).toEqual([]);
			expect(errors(file, 'Module1', `Debug.Print TypeName(Sheet1.${control})`), control).toEqual([]);
		}
		expect(errors(file, 'Data', 'Debug.Print TypeName(DataBtn)')).toEqual([]);
		expect(errors(file, 'Module1', 'Debug.Print TypeName(Data.DataBtn)')).toEqual([]);
		// The sheet's own members still resolve bare.
		expect(errors(file, 'Sheet1', 'Debug.Print UsedRange.Address')).toEqual([]);
	});

	it('reports what the VBE refuses', () => {
		expect(errors(file, 'Sheet1', 'Debug.Print TypeName(Me.Nope)')).toEqual(['member-not-found']);
		expect(errors(file, 'Sheet1', 'Debug.Print TypeName(CtlMisspelt)')).toEqual(['undeclared-variable']);
		expect(errors(file, 'Sheet1', 'Debug.Print TypeName(Me.DataBtn)')).toEqual(['member-not-found']);
		expect(errors(file, 'Data', 'Debug.Print TypeName(Me.CtlCommandButton)')).toEqual(['member-not-found']);
	});
});
