// Diagnostics tests: Access form, report and control event handlers (issue
// #227). The verdicts are the issue's, measured in Access 16.0 through
// pyVBAharness on databases built with pyOpenVBA, and agree with the event
// signatures of MSACC.OLB that the Access model carries.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';
import { byCode } from '../helpers/diagnostics';

const CONTROLS = [
	{ name: 'Qty', type: 'Access.TextBox' },
	{ name: 'Pick', type: 'Access.ComboBox' },
	{ name: 'Go', type: 'Access.CommandButton' },
];

function hits(designer: 'Access.Form' | 'Access.Report', declaration: string, body = '    Debug.Print 1\n'): string[] {
	const src = `Option Compare Database\nOption Explicit\n${declaration.replace(/^(Private |Public )?(Sub|Function) (.*)$/, (_, visibility, kind, rest) => `${visibility ?? 'Private '}${kind} ${rest}`)}\n${body}End ${declaration.includes('Function ') ? 'Function' : 'Sub'}\n`;
	return byCode(analyzeModule(src, {
		moduleName: designer === 'Access.Report' ? 'Report_R1' : 'Form_Orders',
		moduleKind: 'userform',
		designerClass: designer,
		implicitMembers: designer === 'Access.Form' ? CONTROLS : [],
		host: 'access',
		hostModel: getAccessObjectModel(),
	}), 'event-handler-signature').map((d) => d.message);
}

describe('Access control handlers take Access events, not MSForms ones (issue #227)', () => {
	it.each([
		'Sub Qty_BeforeUpdate(Cancel As Integer)',
		'Sub Qty_Exit(Cancel As Integer)',
		'Sub Pick_BeforeUpdate(Cancel As Integer)',
		'Sub Go_Exit(Cancel As Integer)',
		'Sub Qty_Click()',
		'Sub Qty_AfterUpdate()',
		'Sub Qty_KeyDown(KeyCode As Integer, Shift As Integer)',
	])('takes %s', (declaration) => {
		expect(hits('Access.Form', declaration)).toEqual([]);
	});

	it.each(['Sub Qty_AfterUpdate(ByVal x As Long)', 'Sub Qty_BeforeUpdate(ByVal Cancel As Integer)'])('flags %s', (declaration) => {
		const messages = hits('Access.Form', declaration);
		expect(messages).toHaveLength(1);
		expect(messages[0]).not.toContain('ReturnBoolean');
	});
});

describe("an Access form's and report's own events (issue #227)", () => {
	it.each([
		'Sub Form_Load(ByVal x As Long)',
		'Sub Form_BeforeUpdate(ByVal Cancel As Integer)',
		'Sub Form_BeforeUpdate(Cancel As Boolean)',
		'Sub Form_BeforeUpdate()',
		'Sub Form_Error(DataErr As Integer)',
		'Sub Form_KeyDown(KeyCode As Long, Shift As Integer)',
		'Function Form_Load()',
	])('flags %s', (declaration) => {
		expect(hits('Access.Form', declaration)).toHaveLength(1);
	});

	it.each([
		'Sub Form_Load()',
		'Sub Form_BeforeUpdate(Cancel As Integer)',
		'Sub Form_Open(Cancel As Integer)',
		'Sub Form_Unload(Cancel As Integer)',
		'Sub Form_Error(DataErr As Integer, Response As Integer)',
		'Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)',
		'Sub Form_Current()',
		'Sub Form_Dirty(Cancel As Integer)',
		'Sub Form_Delete(Cancel As Integer)',
		'Public Sub Form_Load()',
	])('takes %s', (declaration) => {
		expect(hits('Access.Form', declaration)).toEqual([]);
	});

	it.each(['Sub Report_NoData()', 'Sub Report_Open(ByVal Cancel As Integer)'])('flags the report handler %s', (declaration) => {
		expect(hits('Access.Report', declaration)).toHaveLength(1);
	});

	it.each(['Sub Report_NoData(Cancel As Integer)', 'Sub Report_Page()'])('takes the report handler %s', (declaration) => {
		expect(hits('Access.Report', declaration)).toEqual([]);
	});
});
