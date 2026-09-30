// Diagnostics tests: a module's own object in its code (issue #228). Each
// case was measured in 64-bit Excel and Word 16.0 (build 20326, 2026-09-30),
// with the document's code written through VBIDE and a full compile.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';
import { byCode } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

type Options = Parameters<typeof analyzeModule>[1];

function undeclared(body: string, options: Options): string[] {
	const src = `Option Explicit\nPublic Function Probe() As Variant\n    ${body}\nEnd Function\n`;
	// The rule runs only with the project's known names, as in a project.
	const diags = analyzeModule(src, { knownIdentifiers: new Set(), knownProcedures: new Set(), ...options });
	return [...byCode(diags, 'undeclared-variable'), ...byCode(diags, 'unknown-call')].map((d) => src.slice(d.span.start, d.span.end));
}

const excel = getExcelObjectModel();
const word = getWordObjectModel();
const sheet: Options = { moduleName: 'Sheet1', moduleKind: 'document', documentType: 'worksheet', hostModel: excel };
const sheetUnknown: Options = { moduleName: 'Sheet1', moduleKind: 'document', hostModel: excel };
const book: Options = { moduleName: 'ThisWorkbook', moduleKind: 'document', hostModel: excel };
const doc: Options = { moduleName: 'ThisDocument', moduleKind: 'document', hostModel: word, host: 'word' };
const form: Options = { moduleName: 'F1', moduleKind: 'userform', implicitMembers: [] };

describe("a document's own members named bare (issue #228)", () => {
	it.each([
		[sheet, 'Probe = UsedRange.Address'],
		[sheet, 'Probe = Index'],
		[sheet, 'Probe = CodeName'],
		[sheet, 'Probe = Shapes.Count'],
		[sheetUnknown, 'Probe = UsedRange.Address'],
		[book, 'Probe = FullName'],
		[book, 'Probe = Saved'],
		[book, 'Probe = CodeName'],
		[doc, 'Probe = Content.Text'],
		[doc, 'Probe = Paragraphs.Count'],
		[doc, 'Probe = FullName'],
		[doc, 'Probe = Tables.Count'],
	])('takes them (%#)', (options, body) => {
		expect(undeclared(body, options)).toEqual([]);
	});

	it.each([
		[sheet, 'Probe = Nope', 'Nope'],
		[sheet, 'Probe = FullName', 'FullName'],
		[book, 'Probe = UsedRange.Address', 'UsedRange'],
		[doc, 'Probe = Nope', 'Nope'],
		[doc, 'Probe = UsedRange', 'UsedRange'],
	])('still reports a name the object lacks (%#)', (options, body, name) => {
		expect(undeclared(body, options)).toEqual([name]);
	});
});

describe("a UserForm's own members named bare (issue #228)", () => {
	it.each([
		'Probe = Controls.Count',
		'Probe = ActiveControl Is Nothing',
		'Tag = "t": Probe = Tag',
		'Probe = BackColor',
		'Probe = StartUpPosition',
		'Probe = ScrollBars',
		'Probe = InsideWidth > 0',
		'Repaint',
		'Probe = Caption',
		'Hide',
	])('takes %s', (body) => {
		expect(undeclared(body, form)).toEqual([]);
	});

	it.each([['Probe = Nope', 'Nope'], ['Probe = UsedRange.Address', 'UsedRange']])('still reports %s', (body, name) => {
		expect(undeclared(body, form)).toEqual([name]);
	});
});

describe("a document's own handlers are checked when empty (issue #228)", () => {
	const empty = (signature: string, options: Options) => byCode(analyzeModule(`Option Explicit\nPrivate Sub ${signature}\nEnd Sub\n`, options), 'event-handler-signature');

	it.each([
		['Worksheet_Activate(ByVal x As Long)', sheet],
		['Worksheet_Change()', sheet],
		['Workbook_Open(ByVal x As Long)', book],
		['Workbook_BeforeClose(ByVal Cancel As Boolean)', book],
		['Document_Open(ByVal x As Long)', doc],
		['Document_Close(ByVal x As Long)', doc],
	])('flags an empty %s', (signature, options) => {
		expect(empty(signature, options)).toHaveLength(1);
	});

	it.each([
		['Worksheet_Change(ByVal Target As Range)', sheet],
		['Document_Open()', doc],
	])('takes an empty %s', (signature, options) => {
		expect(empty(signature, options)).toHaveLength(0);
	});

	it('leaves an empty WithEvents handler alone', () => {
		const src = 'Option Explicit\nPrivate WithEvents app As Application\nPrivate Sub app_SheetActivate()\nEnd Sub\n';
		expect(byCode(analyzeModule(src, { moduleName: 'Class1', moduleKind: 'class', hostModel: excel }), 'event-handler-signature')).toHaveLength(0);
	});
});

describe("Word's ThisDocument as an assignment target (issue #228)", () => {
	const errors = (line: string) => {
		const src = `Option Explicit\nFunction Main() As Variant\n    ${line}\nEnd Function\n`;
		return analyzeProjectModule(src, [
			{ moduleName: 'Module1', source: src },
			{ moduleName: 'ThisDocument', source: 'Option Explicit\n', type: 'document' },
		], 'Module1', { hostModel: word, host: 'word' }).filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
	};

	it.each(['Set ThisDocument = Nothing', 'ThisDocument = 5'])('flags %s at compile time', (line) => {
		const hits = errors(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^set-requires-object: .*Invalid use of property/);
	});
});
