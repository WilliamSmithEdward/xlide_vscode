// Diagnostics tests: an event handler declared unlike its event (issue #195).
// Each refused declaration was measured through pyVBAharness on 2026-09-29 in
// Excel 16.0 (build 20326), in a class module with a WithEvents variable; each
// quiet one compiles there. A sheet module's handlers use the same event
// interface as a Worksheet WithEvents variable.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { HOST_EVENT_SIGNATURES } from '../../src/analyzer/host/eventSignaturesData';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';
import type { VbaDiagnostic } from '../../src/analyzer';

const CODE = 'event-handler-signature';

function inClass(declaration: string, header: string, body = '    Fired = Fired + 1'): string {
	const end = /\bFunction\b/.test(header) ? 'End Function' : 'End Sub';
	return `Option Explicit\n${declaration}\nPublic Fired As Long\n${header}\n${body}\n${end}\n`;
}

describe('event-handler-signature (issue #195)', () => {
	it('refuses a handler whose passing, types, count or kind differ from the event', () => {
		const app = 'Private WithEvents app As Application';
		const cases: Array<[string, string, string]> = [
			[app, 'Private Sub app_SheetChange(Sh As Object, ByVal Target As Range)', "'Sh' must be ByVal"],
			[app, 'Private Sub app_SheetChange(ByVal Sh As Variant, ByVal Target As Range)', "'Sh' is Variant, and the event's is Object"],
			[app, 'Private Sub app_SheetChange(ByVal Sh, ByVal Target As Range)', "is Variant"],
			[app, 'Private Sub app_SheetChange(ByVal Sh As Worksheet, ByVal Target As Range)', 'is Worksheet'],
			[app, 'Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Object)', "'Target' is Object, and the event's is Range"],
			[app, 'Private Sub app_SheetChange(ByVal Sh As Object)', 'the event passes 2 parameters, this Sub takes 1'],
			[app, 'Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range, Optional x As Long)', 'takes 3'],
			[app, 'Private Sub app_SheetChange(ParamArray p() As Variant)', 'takes 1'],
			[app, 'Private Function app_SheetChange(ByVal Sh As Object, ByVal Target As Range) As Long', 'an event handler is a Sub'],
			[app, 'Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, ByVal Cancel As Boolean)', "'Cancel' must be ByRef"],
			['Private WithEvents ws As Worksheet', 'Private Sub ws_Change(Target As Range)', "'Target' must be ByVal"],
			['Private WithEvents app As Excel.Application', 'Private Sub app_NewWorkbook(Wb As Workbook)', "'Wb' must be ByVal"],
		];
		for (const [declaration, header, message] of cases) {
			const src = inClass(declaration, header);
			const hits = byCode(analyzeModule(src, { moduleKind: 'class' }), CODE);
			expect(hits, header).toHaveLength(1);
			expect(hits[0].message, header).toContain(message);
			expect(hits[0].message).toContain('Procedure declaration does not match description of event');
		}
	});

	it('accepts renamed parameters, a library prefix, Public, ByRef written out, an enum As Long and an empty body', () => {
		const quiet: Array<[string, string, string?]> = [
			['Private WithEvents app As Application', 'Public Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range)'],
			['Private WithEvents app As Application', 'Private Sub app_SheetChange(ByVal s As Object, ByVal r As Range)'],
			['Private WithEvents app As Application', 'Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)'],
			['Private WithEvents app As Application', 'Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, ByRef Cancel As Boolean)'],
			['Private WithEvents wb As Workbook', 'Private Sub wb_AfterXmlExport(ByVal Map As XmlMap, ByVal Url As String, ByVal Result As Long)'],
			['Private WithEvents app As Application', 'Private Sub app_SheetChange(Sh As Object, ByVal Target As Range)', ''],
			['Private WithEvents app As Application', 'Private Sub app_NotAnEvent(Sh As Object)'],
		];
		for (const [declaration, header, body] of quiet) {
			const src = inClass(declaration, header, body);
			expect(byCode(analyzeModule(src, { moduleKind: 'class' }), CODE), header).toHaveLength(0);
		}
	});

	it('checks a sheet module\'s own handlers and a UserForm control\'s', () => {
		const sheet = 'Option Explicit\nPrivate Sub Worksheet_Change(Target As Range)\n    Debug.Print 1\nEnd Sub\n';
		expectDiagnostic(sheet, analyzeModule(sheet, { moduleKind: 'document', documentType: 'worksheet' }), CODE, { span: 'Worksheet_Change', message: "'Target' must be ByVal" });
		// Exit is the control extender's event, not the TextBox's own.
		const form = 'Option Explicit\nPrivate Sub CommandButton1_Click(ByVal x As Long)\n    Debug.Print 1\nEnd Sub\n'
			+ 'Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)\n    Debug.Print 1\nEnd Sub\n'
			+ 'Private Sub TextBox2_Exit(Cancel As MSForms.ReturnBoolean)\n    Debug.Print 1\nEnd Sub\n';
		const hits = byCode(analyzeModule(form, {
			moduleKind: 'userform',
			implicitMembers: [
				{ name: 'CommandButton1', type: 'MSForms.CommandButton' },
				{ name: 'TextBox1', type: 'MSForms.TextBox' },
				{ name: 'TextBox2', type: 'MSForms.TextBox' },
			],
		}), CODE);
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining("'CommandButton1_Click' does not match the event CommandButton.Click()"),
			expect.stringContaining("'TextBox2_Exit' does not match the event Control.Exit(ByVal Cancel As ReturnBoolean)"),
		]);
	});

	it('holds the events as the type libraries declare them', () => {
		expect(HOST_EVENT_SIGNATURES['Excel.Application'].SheetChange).toBe('ByVal Sh As Object, ByVal Target As Range');
		expect(HOST_EVENT_SIGNATURES['Excel.Workbook'].BeforeClose).toBe('Cancel As Boolean');
		expect(HOST_EVENT_SIGNATURES['MSForms.Control'].Exit).toBe('ByVal Cancel As ReturnBoolean');
	});
});

describe('handlers for an event the project declares (issue #220)', () => {
	const class2 = 'Public Enum Kind\n    kA = 1\nEnd Enum\nPublic Event Changed(ByVal v As Long, s As String)\nPublic Event Done(x)\nPublic Event Got(a() As Long)\nPublic Event Picked(ByVal k As Kind)\nPublic Event Sent(ByVal c As Collection)\n';
	const handler = (event: string, params: string, body = '    Debug.Print 1\n'): VbaDiagnostic[] => {
		const class1 = `Private WithEvents src As Class2\nPrivate Sub src_${event}(${params})\n${body}End Sub\n`;
		return byCode(analyzeProjectModule(class1, [
			{ moduleName: 'Class1', source: class1, type: 'class' },
			{ moduleName: 'Class2', source: class2, type: 'class' },
		], 'Class1'), 'event-handler-signature');
	};

	it.each([
		['Changed', 'v As Long, s As String', 'must be ByVal'],
		['Changed', 'ByVal v As Long, ByVal s As String', 'must be ByRef'],
		['Changed', 'ByVal v As Integer, s As String', 'Integer'],
		['Changed', 'ByVal v As Variant, s As Variant', 'Variant'],
		['Changed', 'ByVal v As Long', 'passes 2'],
		['Changed', 'ByVal v As Long, s As String, ByVal x As Long', 'passes 2'],
		['Changed', '', 'passes 2'],
		['Changed', 'ByVal v As Long, Optional s As String', 'Optional'],
		['Done', 'x As Long', 'Long'],
		['Got', 'a As Long', 'not an array'],
		['Sent', 'ByVal c As Object', 'Object'],
	])('flags src_%s(%s)', (event, params, text) => {
		const hits = handler(event, params);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain(text);
		expect(hits[0].message).toContain('Class2.');
	});

	it.each([
		['Changed', 'ByVal v As Long, s As String'],
		['Changed', 'ByVal n As Long, t As String'],
		['Changed', 'ByVal v As Long, ByRef s As String'],
		['Changed', 'ByVal v&, s$'],
		['Done', 'x'],
		['Done', 'x As Variant'],
		['Got', 'a() As Long'],
		['Picked', 'ByVal k As Long'],
		['Picked', 'ByVal k As Kind'],
		['Sent', 'ByVal c As Collection'],
	])('stays quiet on src_%s(%s)', (event, params) => {
		expect(handler(event, params)).toHaveLength(0);
	});

	it('stays quiet on an empty handler that does not match, as for a library event', () => {
		expect(handler('Changed', 'v As Long, s As String', '')).toHaveLength(0);
		expect(handler('Changed', '', '')).toHaveLength(0);
	});
});
