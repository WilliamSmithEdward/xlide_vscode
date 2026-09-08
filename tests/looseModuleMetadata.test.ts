import { describe, expect, it } from 'vitest';
import { looseModuleMetadata } from '../src/looseModuleMetadata';
import { moduleKindFromType } from '../src/vbaProjectAnalysis';
import { analyzeModule } from '../src/analyzer';
import { blankDesignerHeader } from '../src/vba/moduleSource';

function clsHeader(name: string, predeclared: boolean, exposed: boolean): string {
	return [
		'VERSION 1.0 CLASS',
		'BEGIN',
		"  MultiUse = -1  'True",
		'END',
		`Attribute VB_Name = "${name}"`,
		'Attribute VB_GlobalNameSpace = False',
		'Attribute VB_Creatable = False',
		`Attribute VB_PredeclaredId = ${predeclared ? 'True' : 'False'}`,
		`Attribute VB_Exposed = ${exposed ? 'True' : 'False'}`,
	].join('\r\n');
}

const SHEET1 = [
	clsHeader('Sheet1', true, true),
	'Option Explicit',
	'',
	'Private Sub Worksheet_Change(ByVal Target As Range)',
	'    Debug.Print Me.Name',
	'End Sub',
	'',
	'Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)',
	'End Sub',
	'',
].join('\r\n');

const TICKET = [
	clsHeader('Ticket', false, false),
	'Option Explicit',
	'',
	'Public Sub Describe()',
	'    Debug.Print TypeName(Me)',
	'End Sub',
	'',
].join('\r\n');

describe('looseModuleMetadata', () => {
	it('reads a class from a .cls that claims nothing else', () => {
		expect(looseModuleMetadata('Ticket.cls', TICKET)).toEqual({
			moduleType: 'class',
			documentType: undefined,
		});
	});

	it('reads a worksheet document module from an exported Sheet1.cls', () => {
		expect(looseModuleMetadata('Sheet1.cls', SHEET1)).toEqual({
			moduleType: 'document',
			documentType: 'worksheet',
		});
	});

	it('reads ThisWorkbook as the workbook document module', () => {
		const source = clsHeader('ThisWorkbook', true, true);
		expect(looseModuleMetadata('ThisWorkbook.cls', source)).toEqual({
			moduleType: 'document',
			documentType: 'workbook',
		});
	});

	it('keeps a .bas standard however it is named', () => {
		// Classifying by content would make this a document module on its name.
		expect(looseModuleMetadata('Sheet1.bas', 'Option Explicit\r\n')).toEqual({
			moduleType: 'standard',
			documentType: undefined,
		});
	});

	it('maps the designer extensions to their own kinds', () => {
		expect(looseModuleMetadata('frmMain.frm', '').moduleType).toBe('userform');
		expect(looseModuleMetadata('ctlGrid.ctl', '').moduleType).toBe('usercontrol');
		expect(looseModuleMetadata('pagGeneral.pag', '').moduleType).toBe('propertypage');
		expect(looseModuleMetadata('dsrReport.dsr', '').moduleType).toBe('designer');
	});

	it('gives every object module a kind where Me is valid', () => {
		for (const fileName of ['Ticket.cls', 'Sheet1.cls', 'frmMain.frm', 'ctlGrid.ctl']) {
			const kind = moduleKindFromType(looseModuleMetadata(fileName, TICKET).moduleType);
			expect(kind, fileName).not.toBe('standard');
		}
	});
});

describe('a loose file analyzed with the kind it states (#73)', () => {
	const analyzeLoose = (fileName: string, source: string) => {
		const meta = looseModuleMetadata(fileName, source);
		return analyzeModule(blankDesignerHeader(source), {
			moduleName: fileName.replace(/\.[^.]+$/, ''),
			moduleType: meta.moduleType,
			moduleKind: moduleKindFromType(meta.moduleType),
			documentType: meta.documentType,
		} as never).map((d) => d.code);
	};

	it('does not report Me in a loose class', () => {
		expect(analyzeLoose('Ticket.cls', TICKET)).toEqual([]);
	});

	it('does not report Me or the worksheet handlers in a loose Sheet1.cls', () => {
		expect(analyzeLoose('Sheet1.cls', SHEET1)).toEqual([]);
	});

	it('still reports Me in a loose standard module', () => {
		const bas = 'Option Explicit\r\n\r\nPublic Sub T()\r\n    Debug.Print Me.Name\r\nEnd Sub\r\n';
		expect(analyzeLoose('Module1.bas', bas)).toContain('me-outside-object-module');
	});

	it('still reports a worksheet handler that sits in a loose standard module', () => {
		const bas = [
			'Option Explicit',
			'',
			'Private Sub Worksheet_Change(ByVal Target As Range)',
			'End Sub',
			'',
		].join('\r\n');
		expect(analyzeLoose('Module1.bas', bas)).toContain('event-handler-module-scope');
	});

	it('still reports a worksheet handler that sits in a loose ordinary class', () => {
		const source = [
			clsHeader('Ticket', false, false),
			'Option Explicit',
			'',
			'Private Sub Worksheet_Change(ByVal Target As Range)',
			'End Sub',
			'',
		].join('\r\n');
		expect(analyzeLoose('Ticket.cls', source)).toContain('event-handler-module-scope');
	});
});
