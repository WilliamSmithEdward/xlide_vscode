import { describe, it, expect } from 'vitest';
import {
	buildModuleSymbols,
	DocRegistry,
	extractLeadingDoc,
	parseDocBody,
	parseMetadataFile,
	renderDocMarkdown,
	resolveHover,
	resolveSignatureHelp,
} from '../src/analyzer';

describe('parseDocBody', () => {
	it('parses XML tags into a doc model', () => {
		const doc = parseDocBody(
			[
				'<summary>Computes the tax owed.</summary>',
				'<param name="Amount">Pre-tax amount.</param>',
				'<param name="Rate">Tax rate as a fraction.</param>',
				'<returns>The tax owed.</returns>',
				'<remarks>Rounds to the cent.</remarks>',
			].join('\n'),
			'inline',
		);
		expect(doc.summary).toBe('Computes the tax owed.');
		expect(doc.params).toEqual([
			{ name: 'Amount', text: 'Pre-tax amount.' },
			{ name: 'Rate', text: 'Tax rate as a fraction.' },
		]);
		expect(doc.returns).toBe('The tax owed.');
		expect(doc.remarks).toBe('Rounds to the cent.');
		expect(doc.source).toBe('inline');
	});

	it('parses optional type, unit, and value hints', () => {
		const doc = parseDocBody(
			[
				'<summary>Computes the invoice total.</summary>',
				'<param name="Subtotal" type="Currency" unit="money">Pre-tax amount.</param>',
				'<param name="TaxRate" type="Double" unit="decimal" value="0 to 1">Tax rate.</param>',
				'<returns type="Currency" unit="money">Invoice total.</returns>',
			].join('\n'),
			'inline',
		);
		expect(doc.params[0]).toEqual({
			name: 'Subtotal',
			text: 'Pre-tax amount.',
			type: 'Currency',
			unit: 'money',
		});
		expect(doc.params[1]).toEqual({
			name: 'TaxRate',
			text: 'Tax rate.',
			type: 'Double',
			unit: 'decimal',
			value: '0 to 1',
		});
		expect(doc.returns).toBe('Invoice total.');
		expect(doc.returnsType).toBe('Currency');
		expect(doc.returnsUnit).toBe('money');
	});

	it('treats untagged text as the summary', () => {
		const doc = parseDocBody('Just a plain note.', 'inline');
		expect(doc.summary).toBe('Just a plain note.');
		expect(doc.params).toHaveLength(0);
	});

	it('collapses whitespace and decodes entities', () => {
		const doc = parseDocBody(
			'<summary>A &lt;b&gt; tag\n   wraps   text</summary>',
			'inline',
		);
		expect(doc.summary).toBe('A <b> tag wraps text');
	});

	it('preserves example layout as a code block source', () => {
		const doc = parseDocBody(
			'<example>Dim x As Long\nx = ComputeTax(100)</example>',
			'inline',
		);
		expect(doc.example).toBe('Dim x As Long\nx = ComputeTax(100)');
	});
});

describe('extractLeadingDoc', () => {
	it('captures a contiguous triple-quote block above a declaration', () => {
		const src = [
			"''' <summary>Greets the user.</summary>",
			"''' <param name=\"Name\">Who to greet.</param>",
			'Public Sub Greet(ByVal Name As String)',
			'End Sub',
		].join('\n');
		const declStart = src.indexOf('Public Sub');
		const doc = extractLeadingDoc(src, declStart);
		expect(doc?.summary).toBe('Greets the user.');
		expect(doc?.params).toEqual([{ name: 'Name', text: 'Who to greet.' }]);
	});

	it('stops at a non-doc line and ignores ordinary comments', () => {
		const src = [
			"' an ordinary comment",
			"''' <summary>Only this line counts.</summary>",
			'Sub A()',
			'End Sub',
		].join('\n');
		const doc = extractLeadingDoc(src, src.indexOf('Sub A'));
		expect(doc?.summary).toBe('Only this line counts.');
	});

	it('returns undefined when there is no doc comment', () => {
		const src = 'Sub A()\nEnd Sub\n';
		expect(extractLeadingDoc(src, 0)).toBeUndefined();
	});
});

describe('buildModuleSymbols inline docs', () => {
	it('attaches docs to procedures and module variables', () => {
		const src = [
			"''' <summary>The running total.</summary>",
			'Public Total As Long',
			'',
			"''' <summary>Adds to the total.</summary>",
			'Public Sub Add(ByVal N As Long)',
			'    Total = Total + N',
			'End Sub',
		].join('\n');
		const mod = buildModuleSymbols('Module1', 'standard', src);
		const total = mod.root.children?.find((c) => c.name === 'Total');
		const add = mod.root.children?.find((c) => c.name === 'Add');
		expect(total?.doc?.summary).toBe('The running total.');
		expect(add?.doc?.summary).toBe('Adds to the total.');
	});
});

describe('parseMetadataFile', () => {
	it('parses member entries with qualified and bare names', () => {
		const xml = [
			'<xlideDoc>',
			'  <member name="Module1.ComputeTax">',
			'    <summary>Returns the tax owed.</summary>',
			'    <param name="Amount">The pre-tax amount.</param>',
			'  </member>',
			'  <member name="MsgBox">',
			'    <summary>Team note about MsgBox.</summary>',
			'  </member>',
			'</xlideDoc>',
		].join('\n');
		const entries = parseMetadataFile(xml);
		expect(entries).toHaveLength(2);
		expect(entries[0].name).toBe('Module1.ComputeTax');
		expect(entries[0].doc.summary).toBe('Returns the tax owed.');
		expect(entries[1].name).toBe('MsgBox');
	});
});

describe('DocRegistry', () => {
	it('resolves qualified, bare, and member-of-qualified lookups', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				[
					'<member name="Module1.ComputeTax"><summary>Q</summary></member>',
					'<member name="MsgBox"><summary>B</summary></member>',
				].join('\n'),
			),
		);
		expect(reg.lookup('ComputeTax', 'Module1')?.summary).toBe('Q');
		expect(reg.lookup('ComputeTax')?.summary).toBe('Q');
		expect(reg.lookup('MsgBox')?.summary).toBe('B');
		expect(reg.lookup('Unknown')).toBeUndefined();
	});
});

describe('hover documentation', () => {
	it('shows the inline doc summary for a user procedure', () => {
		const src = [
			"''' <summary>Adds two numbers.</summary>",
			'Public Function Add(A As Long, B As Long) As Long',
			'    Add = A + B',
			'End Function',
			'',
			'Sub Caller()',
			'    Dim r As Long',
			'    r = Add(1, 2)',
			'End Sub',
		].join('\n');
		const offset = src.lastIndexOf('Add(') + 1;
		const info = resolveHover(src, offset, { moduleName: 'Module1' });
		expect(info?.documentation).toContain('Adds two numbers.');
	});

	it('lets external metadata document a runtime function', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				'<member name="MsgBox"><summary>Team note about MsgBox.</summary></member>',
			),
		);
		const src = 'Sub A()\n    MsgBox "hi"\nEnd Sub\n';
		const offset = src.indexOf('MsgBox') + 1;
		const info = resolveHover(src, offset, { moduleName: 'Module1', docRegistry: reg });
		expect(info?.documentation).toContain('Team note about MsgBox.');
	});

	it('lets external metadata override generated host hover docs', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				'<member name="Application.Calculate"><summary>Team calc note.</summary></member>',
			),
		);
		const src = 'Sub A()\n    Application.Calculate\nEnd Sub\n';
		const offset = src.indexOf('Calculate') + 1;
		const info = resolveHover(src, offset, { docRegistry: reg });
		expect(info?.documentation).toContain('Team calc note.');
		expect(info?.documentation).not.toContain('Calculates all open workbooks');
	});

	it('prefers the inline doc over an external entry', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				'<member name="Module1.Add"><summary>External wins?</summary></member>',
			),
		);
		const src = [
			"''' <summary>Inline wins.</summary>",
			'Public Sub Add()',
			'End Sub',
		].join('\n');
		const offset = src.indexOf('Public Sub Add') + 'Public Sub '.length + 1;
		const info = resolveHover(src, offset, { moduleName: 'Module1', docRegistry: reg });
		expect(info?.documentation).toContain('Inline wins.');
		expect(info?.documentation).not.toContain('External wins');
	});

	it('shows inline docs for project class property members', () => {
		const src = 'Sub Caller()\n    Dim p As Person\n    p.Age\nEnd Sub\n';
		const offset = src.indexOf('Age') + 1;
		const info = resolveHover(src, offset, {
			moduleName: 'Module1',
			projectClassMembers: [
				{
					name: 'Person',
					kind: 'class',
					moduleName: 'Person',
					members: [
						{
							name: 'Age',
							kind: 'property',
							returns: 'Integer',
							moduleName: 'Person',
							doc: {
								summary: 'Age in whole years.',
								params: [],
								source: 'inline',
							},
						},
					],
				},
			],
		});
		expect(info?.signature).toBe('Person.Age As Integer');
		expect(info?.documentation).toContain('Age in whole years.');
	});
});

describe('signature help documentation', () => {
	it('attaches summary and per-parameter docs from an inline comment', () => {
		const src = [
			"''' <summary>Adds two numbers.</summary>",
			"''' <param name=\"A\">First addend.</param>",
			"''' <param name=\"B\">Second addend.</param>",
			'Public Function Add(A As Long, B As Long) As Long',
			'End Function',
			'',
			'Sub Caller()',
			'    Dim r As Long',
			'    r = Add(',
			'End Sub',
		].join('\n');
		const offset = src.lastIndexOf('Add(') + 'Add('.length;
		const info = resolveSignatureHelp(src, offset, { moduleSource: src });
		expect(info?.documentation).toContain('Adds two numbers.');
		expect(info?.parameters[0].documentation).toBe('First addend.');
		expect(info?.parameters[1].documentation).toBe('Second addend.');
	});

	it('lets external metadata override generated host signature docs', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				[
					'<member name="Workbooks.Open">',
					'  <summary>Team open note.</summary>',
					'  <param name="Filename">Team path note.</param>',
					'</member>',
				].join('\n'),
			),
		);
		const src = 'Sub A()\n    Workbooks.Open(\nEnd Sub\n';
		const offset = src.indexOf('Open(') + 'Open('.length;
		const info = resolveSignatureHelp(src, offset, { docRegistry: reg });
		expect(info?.label.startsWith('Open(Filename As String')).toBe(true);
		expect(info?.documentation).toContain('Team open note.');
		expect(info?.parameters[0].documentation).toBe('Team path note.');
	});

	it('synthesizes a signature from external metadata when none is known', () => {
		const reg = new DocRegistry();
		reg.add(
			parseMetadataFile(
				[
					'<member name="DoThing">',
					'  <signature>DoThing(Path As String) As Boolean</signature>',
					'  <summary>Does the thing.</summary>',
					'  <param name="Path">Where to do it.</param>',
					'</member>',
				].join('\n'),
			),
		);
		const src = 'Sub A()\n    DoThing(\nEnd Sub\n';
		const offset = src.indexOf('DoThing(') + 'DoThing('.length;
		const info = resolveSignatureHelp(src, offset, { moduleSource: src, docRegistry: reg });
		expect(info?.label).toBe('DoThing(Path As String) As Boolean');
		expect(info?.documentation).toContain('Does the thing.');
		expect(info?.parameters[0].documentation).toBe('Where to do it.');
	});
});

describe('renderDocMarkdown', () => {
	it('renders summary, parameters, returns, and remarks', () => {
		const doc = parseDocBody(
			[
				'<summary>S</summary>',
				'<param name="X">px</param>',
				'<returns>R</returns>',
				'<remarks>M</remarks>',
			].join('\n'),
			'inline',
		);
		const md = renderDocMarkdown(doc);
		expect(md).toContain('S');
		expect(md).toContain('`X`: px');
		expect(md).toContain('**Returns:** R');
		expect(md).toContain('**Remarks:** M');
	});

	it('renders optional type and unit hints', () => {
		const doc = parseDocBody(
			[
				'<summary>S</summary>',
				'<param name="Amount" type="Currency" unit="money">px</param>',
				'<returns type="Currency" unit="money">R</returns>',
			].join('\n'),
			'inline',
		);
		const md = renderDocMarkdown(doc);
		expect(md).toContain('`Amount` (As Currency, unit: money): px');
		expect(md).toContain('**Returns (As Currency, unit: money):** R');
	});
});
