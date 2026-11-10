// Issue #1555: an ActiveX control on a sheet of an .xlsb workbook saved with
// ActiveX disabled. That save drops the module's `Attribute VB_Control`
// lines, but the sheet's binary part still lists each control, in a
// BrtActiveX record (MS-XLSB 2.4.4), and XLIDE reads them there, as it reads
// an .xlsm's from the sheet's XML part. The fixture is an Excel 16.0 file
// (build 20527) with eleven controls on Sheet1 and one on Data, saved after
// a VBA edit with ActiveX disabled. Reopened with ActiveX enabled, the VBE
// compiled every control bare, through Me and through the sheet's name,
// and refused `Me.Nope`, a misspelt control and another sheet's control,
// in this file and in a workbook built the same way in Office 2016 (build
// 12527).

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { readModulesFromBuffer } from '../../../src/vba/projectService';
import { moduleKindFromType } from '../../../src/vbaProjectAnalysis';

const FILE = path.join(__dirname, '..', '..', 'fixtures', 'binaries', 'ExcelControlsDisabledSaveFixture.xlsb');
const SHEET1 = ['CtlCommandButton', 'CtlTextBox', 'CtlCheckBox', 'CtlOptionButton', 'CtlToggleButton', 'CtlComboBox', 'CtlListBox', 'CtlLabel', 'CtlImage', 'CtlScrollBar', 'CtlSpinButton'];

function entries() {
	return readModulesFromBuffer(fs.readFileSync(FILE), true);
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
		kind: 'analyze', requestId: 1, docKey: 'm', projectKey: 'p', generation: 1, host: 'excel',
		source, moduleName: mod.name, moduleType: mod.type, moduleKind: moduleKindFromType(mod.type), documentType: mod.documentType, designerClass: mod.designerClass,
	});
	return response?.kind === 'result' ? response.diagnostics.filter((d) => d.severity === 'error').map((d) => d.code) : ['no result'];
}

describe('the controls on a sheet of an .xlsb saved with ActiveX disabled (issue #1555)', () => {
	it('has lost the header lines, so the controls come from the sheet parts', () => {
		const sheet1 = entries().find((entry) => entry.name === 'Sheet1');
		expect(sheet1?.source).toBeDefined();
		expect(sheet1?.source).not.toMatch(/VB_Control/i);
	});

	it('gives each sheet module the controls its sheet part lists, as a Worksheet', () => {
		const byName = new Map(entries().map((entry) => [entry.name, entry]));
		expect(byName.get('Sheet1')?.designerClass).toBe('Excel.Worksheet');
		expect(byName.get('Sheet1')?.implicitMembers).toEqual(SHEET1.map((name) => ({ name, type: 'Object' })));
		expect(byName.get('Data')?.designerClass).toBe('Excel.Worksheet');
		expect(byName.get('Data')?.implicitMembers).toEqual([{ name: 'DataBtn', type: 'Object' }]);
		for (const name of ['ThisWorkbook', 'Chart3', 'Module1']) {
			expect(byName.has(name), name).toBe(true);
			expect(byName.get(name)?.designerClass, name).toBeUndefined();
			expect(byName.get(name)?.implicitMembers, name).toBeUndefined();
		}
	});

	it('makes a worksheet with no control a Worksheet too, as in an .xlsm', () => {
		const sheets = readModulesFromBuffer(fs.readFileSync(path.join(path.dirname(FILE), 'SheetsFixture.xlsb')), true)
			.filter((entry) => entry.documentType === 'worksheet');
		expect(sheets.length).toBeGreaterThan(0);
		for (const sheet of sheets) {
			expect(sheet.designerClass, sheet.name).toBe('Excel.Worksheet');
			expect(sheet.implicitMembers, sheet.name).toEqual([]);
		}
	});

	it('accepts what the VBE compiles', () => {
		for (const control of SHEET1) {
			expect(errors('Sheet1', `Debug.Print TypeName(${control})`), control).toEqual([]);
			expect(errors('Sheet1', `Debug.Print TypeName(Me.${control})`), control).toEqual([]);
			expect(errors('Module1', `Debug.Print TypeName(Sheet1.${control})`), control).toEqual([]);
		}
		expect(errors('Data', 'Debug.Print TypeName(DataBtn)')).toEqual([]);
		expect(errors('Module1', 'Debug.Print TypeName(Data.DataBtn)')).toEqual([]);
	});

	it('reports what the VBE refuses', () => {
		expect(errors('Sheet1', 'Debug.Print TypeName(Me.Nope)')).toEqual(['member-not-found']);
		expect(errors('Sheet1', 'Debug.Print TypeName(CtlMisspelt)')).toEqual(['undeclared-variable']);
		expect(errors('Sheet1', 'Debug.Print TypeName(Me.DataBtn)')).toEqual(['member-not-found']);
		expect(errors('Data', 'Debug.Print TypeName(Me.CtlCommandButton)')).toEqual(['member-not-found']);
	});
});
