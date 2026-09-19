// Diagnostics tests: a procedure's XML doc comment against its declaration.
//
// Once a Sub, Function, Property, Declare or Event carries a `'''` block
// written in XML, the block has to describe the surface a caller sees - every
// parameter, and the value of anything that returns one - and nothing else.
// The quick fixes are applied to the text here, because where a new tag lands
// is the whole of their job.

import { describe, expect, it } from 'vitest';
import {
	analyzeModule,
	extractLeadingDoc,
	filterDiagnosticsWithSuppressions,
	resolveDiagnosticCodeActions,
	type VbaDiagnostic,
} from '../../src/analyzer';

const DOC_CODES = [
	'doc-param-missing',
	'doc-param-unknown',
	'doc-returns-missing',
	'doc-returns-unexpected',
	'doc-tag-duplicate',
	'doc-tag-unclosed',
];

function docDiagnostics(src: string): VbaDiagnostic[] {
	return analyzeModule(src).filter((d) => DOC_CODES.includes(d.code));
}

/** `code @ covered text: message`, one per doc finding, in order. */
function findings(src: string): string[] {
	return docDiagnostics(src).map((d) => `${d.code} @ ${src.slice(d.span.start, d.span.end)}: ${d.message}`);
}

function fixTitles(src: string, diagnostic: VbaDiagnostic): string[] {
	return resolveDiagnosticCodeActions(src, { code: diagnostic.code, span: diagnostic.span, data: diagnostic.data })
		.map((action) => action.title);
}

/** The module after the quick fix titled `title` on the finding with `code`. */
function applyFix(src: string, code: string, title: string, index = 0): string {
	const diagnostic = docDiagnostics(src).filter((d) => d.code === code)[index];
	expect(diagnostic, code).toBeDefined();
	const action = resolveDiagnosticCodeActions(src, { code, span: diagnostic.span, data: diagnostic.data })
		.find((candidate) => candidate.title === title);
	expect(action, title).toBeDefined();
	let out = src;
	for (const edit of [...action!.edits].sort((a, b) => b.span.start - a.span.start)) {
		out = out.slice(0, edit.span.start) + edit.newText + out.slice(edit.span.end);
	}
	return out;
}

const lines = (...text: string[]): string => `${text.join('\n')}\n`;

describe('analyzeModule - doc comments describe the whole surface', () => {
	it('accepts a doc comment that describes every parameter and the return value', () => {
		const src = lines(
			"''' <summary>Calculates the invoice total after tax.</summary>",
			"''' <param name=\"Subtotal\" type=\"Currency\" unit=\"money\">The pre-tax invoice amount.</param>",
			"''' <param name=\"TaxRate\" type=\"Double\" unit=\"decimal\">The tax rate as a decimal value.</param>",
			"''' <returns type=\"Currency\" unit=\"money\">The subtotal plus calculated tax.</returns>",
			'Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency',
			'    InvoiceTotal = Subtotal + (Subtotal * TaxRate)',
			'End Function',
		);
		expect(findings(src)).toEqual([]);
	});

	it('reports each parameter the doc comment leaves out, at the parameter, as a warning', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"item\">The item.</param>",
			'Public Sub AddItem(ByVal item As String, ByVal quantity As Long, Optional ByVal note As String)',
			'End Sub',
		);
		expect(findings(src)).toEqual([
			"doc-param-missing @ quantity: The doc comment does not describe parameter 'quantity'.",
			"doc-param-missing @ note: The doc comment does not describe parameter 'note'.",
		]);
		expect(docDiagnostics(src).map((d) => d.severity)).toEqual(['warning', 'warning']);
	});

	it('matches parameter names without regard to case, as the call tip does', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"ITEM\">The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		expect(findings(src)).toEqual([]);
	});

	it('reports a <param> naming a parameter the procedure does not have, at the name', () => {
		const src = lines(
			"''' <summary>Calculates the total.</summary>",
			"''' <param name=\"Rate\">The tax rate.</param>",
			"''' <returns>The total.</returns>",
			'Public Function Total(ByVal TaxRate As Double) As Double',
			'End Function',
		);
		expect(findings(src)).toEqual([
			"doc-param-unknown @ Rate: 'Total' has no parameter named 'Rate'.",
			"doc-param-missing @ TaxRate: The doc comment does not describe parameter 'TaxRate'.",
		]);
	});

	it('reports a <param> with no name', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param>The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		expect(findings(src)).toEqual([
			'doc-param-unknown @ <param>: This <param> has no name="..." to say which parameter it describes.',
			"doc-param-missing @ item: The doc comment does not describe parameter 'item'.",
		]);
	});

	it('reports a parameter described twice, and a single tag given twice', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <summary>Adds something.</summary>",
			"''' <param name=\"item\">The item.</param>",
			"''' <param name=\"Item\">The item again.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		expect(findings(src)).toEqual([
			'doc-tag-duplicate @ <summary>: The doc comment already has a <summary>, and only the first is shown.',
			"doc-tag-duplicate @ Item: The doc comment already describes parameter 'Item'.",
		]);
	});

	it('reports a Function whose doc comment says nothing of what it returns, at its name', () => {
		const src = lines(
			"''' <summary>Counts the rows.</summary>",
			'Public Function RowCount() As Long',
			'End Function',
		);
		expect(findings(src)).toEqual([
			"doc-returns-missing @ RowCount: The doc comment does not describe what 'RowCount' returns.",
		]);
	});

	it('reports a <returns> on anything that returns nothing', () => {
		const src = lines(
			"''' <summary>Clears the cache.</summary>",
			"''' <returns>Nothing.</returns>",
			'Public Sub ClearCache()',
			'End Sub',
			"''' <summary>The name.</summary>",
			"''' <returns>The name.</returns>",
			'Public Property Let Name(ByVal value As String)',
			'End Property',
		);
		expect(findings(src)).toEqual([
			"doc-returns-unexpected @ <returns>: Sub 'ClearCache' returns no value, but its doc comment describes one.",
			"doc-returns-unexpected @ <returns>: Property Let 'Name' returns no value, but its doc comment describes one.",
		]);
	});

	it('reports an empty <param> or <returns>, and takes a type hint as a description', () => {
		const src = lines(
			"''' <summary>Scales a value.</summary>",
			"''' <param name=\"value\"></param>",
			"''' <param name=\"factor\"/>",
			"''' <param name=\"places\" type=\"Long\" unit=\"digits\"/>",
			"''' <returns>  </returns>",
			'Public Function Scaled(ByVal value As Double, ByVal factor As Double, ByVal places As Long) As Double',
			'End Function',
		);
		expect(findings(src)).toEqual([
			"doc-param-missing @ <param name=\"value\">: The <param> for 'value' is empty.",
			"doc-param-missing @ <param name=\"factor\"/>: The <param> for 'factor' is empty.",
			'doc-returns-missing @ <returns>: The <returns> is empty.',
		]);
	});

	it('reports a tag that is never closed, and counts it as there', () => {
		const src = lines(
			"''' <summary>Adds an item.",
			"''' <param name=\"item\">The item.",
			"''' <param name=\"quantity\">How many.</param>",
			'Public Sub AddItem(ByVal item As String, ByVal quantity As Long)',
			'End Sub',
		);
		expect(findings(src)).toEqual([
			'doc-tag-unclosed @ <summary>: This <summary> is not closed; end it with </summary>.',
			'doc-tag-unclosed @ <param name="item">: This <param> is not closed; end it with </param>.',
		]);
	});

	it('reads a property as its value: no <returns> for a Get, no <param> for the value a Let or Set receives', () => {
		const src = lines(
			"''' <summary>Age in whole years.</summary>",
			'Public Property Get Age() As Integer',
			'End Property',
			"''' <summary>Age in whole years.</summary>",
			'Public Property Let Age(ByVal value As Integer)',
			'End Property',
			"''' <summary>The owner.</summary>",
			'Public Property Set Owner(ByVal value As Object)',
			'End Property',
		);
		expect(findings(src)).toEqual([]);
	});

	it("still asks for a property's index parameters", () => {
		const src = lines(
			"''' <summary>The item at a position.</summary>",
			'Public Property Get Item(ByVal index As Long) As Variant',
			'End Property',
			"''' <summary>The item at a position.</summary>",
			'Public Property Let Item(ByVal index As Long, ByVal value As Variant)',
			'End Property',
		);
		expect(findings(src)).toEqual([
			"doc-param-missing @ index: The doc comment does not describe parameter 'index'.",
			"doc-param-missing @ index: The doc comment does not describe parameter 'index'.",
		]);
	});

	it('checks a Declare and an Event the same way', () => {
		const src = lines(
			"''' <summary>Milliseconds since Windows started.</summary>",
			'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
			"''' <summary>Pauses.</summary>",
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)',
			"''' <summary>Raised after a change.</summary>",
			"''' <returns>Nothing.</returns>",
			'Public Event Changed(ByVal oldValue As Long)',
		);
		expect(findings(src)).toEqual([
			"doc-returns-missing @ GetTickCount: The doc comment does not describe what 'GetTickCount' returns.",
			"doc-param-missing @ dwMilliseconds: The doc comment does not describe parameter 'dwMilliseconds'.",
			"doc-returns-unexpected @ <returns>: Event 'Changed' returns no value, but its doc comment describes one.",
			"doc-param-missing @ oldValue: The doc comment does not describe parameter 'oldValue'.",
		]);
	});

	it('leaves a plain-text note, a banner of apostrophes and an undocumented procedure alone', () => {
		const src = lines(
			"''' Adds an item to the cart.",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
			"''''''''''''''''''''''''''''''''",
			"' Procedure: Total",
			"''''''''''''''''''''''''''''''''",
			'Public Function Total(ByVal a As Long) As Long',
			'End Function',
			'Public Function Plain(ByVal a As Long) As Long',
			'End Function',
		);
		expect(findings(src)).toEqual([]);
	});

	it('reads the block through the directives stacked between it and the declaration', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"' @xlide-test",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		expect(findings(src)).toEqual([
			"doc-param-missing @ item: The doc comment does not describe parameter 'item'.",
		]);
	});

	it('skips a procedure in an inactive #If branch', () => {
		const src = lines(
			'#If False Then',
			"''' <summary>Adds an item.</summary>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
			'#End If',
		);
		expect(findings(src)).toEqual([]);
	});
});

describe('doc comment quick fixes', () => {
	it('adds a missing <param> in signature order, written like its neighbour', () => {
		const src = lines(
			"    ''' <summary>Adds an item.</summary>",
			"    ''' <param name=\"item\">The item.</param>",
			"    ''' <param name=\"note\">A note.</param>",
			'    Public Sub AddItem(ByVal item As String, ByVal quantity As Long, ByVal note As String)',
			'    End Sub',
		);
		const [missing] = docDiagnostics(src);
		expect(fixTitles(src, missing)).toContain("Add a <param> for 'quantity'");
		expect(applyFix(src, 'doc-param-missing', "Add a <param> for 'quantity'")).toBe(lines(
			"    ''' <summary>Adds an item.</summary>",
			"    ''' <param name=\"item\">The item.</param>",
			"    ''' <param name=\"quantity\"></param>",
			"    ''' <param name=\"note\">A note.</param>",
			'    Public Sub AddItem(ByVal item As String, ByVal quantity As Long, ByVal note As String)',
			'    End Sub',
		));
	});

	it('keeps the indentation inside the comment as well', () => {
		const src = lines(
			"''' <summary>",
			"'''   Adds an item.",
			"''' </summary>",
			"'''   <param name=\"item\">The item.</param>",
			'Public Sub AddItem(ByVal item As String, ByVal quantity As Long)',
			'End Sub',
		);
		expect(applyFix(src, 'doc-param-missing', "Add a <param> for 'quantity'").split('\n').slice(3, 5)).toEqual([
			"'''   <param name=\"item\">The item.</param>",
			"'''   <param name=\"quantity\"></param>",
		]);
	});

	it('adds every missing <param> at once, after the summary when there are none', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <returns>True when added.</returns>",
			'Public Function AddItem(ByVal item As String, ByVal quantity As Long) As Boolean',
			'End Function',
		);
		const fixed = applyFix(src, 'doc-param-missing', 'Add the 2 missing <param> tags');
		expect(fixed).toBe(lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"item\"></param>",
			"''' <param name=\"quantity\"></param>",
			"''' <returns>True when added.</returns>",
			'Public Function AddItem(ByVal item As String, ByVal quantity As Long) As Boolean',
			'End Function',
		));
		// Added, the tags are there; empty, they still ask to be written.
		expect(findings(fixed)).toEqual([
			"doc-param-missing @ <param name=\"item\">: The <param> for 'item' is empty.",
			"doc-param-missing @ <param name=\"quantity\">: The <param> for 'quantity' is empty.",
		]);
	});

	it('puts a <param> before the <returns> when there is no summary', () => {
		const src = lines(
			"''' <returns>The total.</returns>",
			'Public Function Total(ByVal a As Long) As Long',
			'End Function',
		);
		expect(applyFix(src, 'doc-param-missing', "Add a <param> for 'a'")).toBe(lines(
			"''' <param name=\"a\"></param>",
			"''' <returns>The total.</returns>",
			'Public Function Total(ByVal a As Long) As Long',
			'End Function',
		));
	});

	it('renames a stale <param> to the parameter nothing describes, as the preferred fix', () => {
		const src = lines(
			"''' <summary>Calculates the total.</summary>",
			"''' <param name=\"Rate\">The tax rate.</param>",
			"''' <returns>The total.</returns>",
			'Public Function Total(ByVal TaxRate As Double) As Double',
			'End Function',
		);
		const unknown = docDiagnostics(src).find((d) => d.code === 'doc-param-unknown')!;
		const actions = resolveDiagnosticCodeActions(src, { code: unknown.code, span: unknown.span, data: unknown.data });
		expect(actions.map((action) => [action.title, action.isPreferred])).toEqual([
			["Rename the <param> to 'TaxRate'", true],
			["Remove the <param> for 'Rate'", false],
		]);
		const renamed = applyFix(src, 'doc-param-unknown', "Rename the <param> to 'TaxRate'");
		expect(renamed).toContain("''' <param name=\"TaxRate\">The tax rate.</param>\n");
		expect(findings(renamed)).toEqual([]);
		expect(applyFix(src, 'doc-param-unknown', "Remove the <param> for 'Rate'")).toBe(lines(
			"''' <summary>Calculates the total.</summary>",
			"''' <returns>The total.</returns>",
			'Public Function Total(ByVal TaxRate As Double) As Double',
			'End Function',
		));
	});

	it('names a nameless <param>', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param>The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		const fixed = applyFix(src, 'doc-param-unknown', "Name the <param> 'item'");
		expect(fixed).toContain("''' <param name=\"item\">The item.</param>\n");
		expect(findings(fixed)).toEqual([]);
	});

	it('removes a repeated tag, and only the tag when it shares its line', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary> <summary>Adds something.</summary>",
			"''' <param name=\"item\">The item.</param>",
			"''' <param name=\"item\">The item",
			"''' again.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		const once = applyFix(src, 'doc-tag-duplicate', 'Remove the repeated <summary>');
		expect(once.split('\n')[0]).toBe("''' <summary>Adds an item.</summary> ");
		expect(applyFix(src, 'doc-tag-duplicate', 'Remove the repeated <param>', 1)).toBe(lines(
			"''' <summary>Adds an item.</summary> <summary>Adds something.</summary>",
			"''' <param name=\"item\">The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		));
	});

	it('adds a <returns> after the last <param>, and removes one a Sub cannot have', () => {
		const fn = lines(
			"''' <summary>Counts the rows.</summary>",
			"''' <param name=\"sheet\">The sheet.</param>",
			'Public Function RowCount(ByVal sheet As Object) As Long',
			'End Function',
		);
		expect(applyFix(fn, 'doc-returns-missing', 'Add a <returns>')).toBe(lines(
			"''' <summary>Counts the rows.</summary>",
			"''' <param name=\"sheet\">The sheet.</param>",
			"''' <returns></returns>",
			'Public Function RowCount(ByVal sheet As Object) As Long',
			'End Function',
		));
		const sub = lines(
			"''' <summary>Clears the cache.</summary>",
			"''' <returns>Nothing.</returns>",
			'Public Sub ClearCache()',
			'End Sub',
		);
		expect(applyFix(sub, 'doc-returns-unexpected', 'Remove the <returns>')).toBe(lines(
			"''' <summary>Clears the cache.</summary>",
			'Public Sub ClearCache()',
			'End Sub',
		));
	});

	it('adds a line above the directives guarding the next one, so each keeps its line', () => {
		const src = lines(
			"''' <summary>Adds.</summary>",
			"' @xlide-analysis-disable-next-line doc-param-missing",
			"''' <param name=\"b\"></param>",
			'Public Sub F(ByVal a As Long, ByVal b As Long)',
			'End Sub',
		);
		// The first finding is b's empty tag, which the directive suppresses.
		const fixed = applyFix(src, 'doc-param-missing', "Add a <param> for 'a'", 1);
		expect(fixed).toBe(lines(
			"''' <summary>Adds.</summary>",
			"''' <param name=\"a\"></param>",
			"' @xlide-analysis-disable-next-line doc-param-missing",
			"''' <param name=\"b\"></param>",
			'Public Sub F(ByVal a As Long, ByVal b As Long)',
			'End Sub',
		));
		const live = filterDiagnosticsWithSuppressions(fixed, docDiagnostics(fixed)).diagnostics;
		expect(live.map((d) => fixed.slice(d.span.start, d.span.end))).toEqual(['<param name="a">']);
	});

	it('removes a disable-next-line with the tag it guarded, and leaves other directives', () => {
		const src = lines(
			"''' <summary>Adds.</summary>",
			"' @xlide-test",
			"' @xlide-analysis-disable-next-line all",
			"''' <param name=\"y\">Gone.</param>",
			'Public Sub F(ByVal x As Long)',
			'End Sub',
		);
		expect(applyFix(src, 'doc-param-unknown', "Remove the <param> for 'y'")).toBe(lines(
			"''' <summary>Adds.</summary>",
			"' @xlide-test",
			'Public Sub F(ByVal x As Long)',
			'End Sub',
		));
	});

	it('writes the module line breaks', () => {
		const src = "''' <summary>Adds an item.</summary>\r\nPublic Sub AddItem(ByVal item As String)\r\nEnd Sub\r\n";
		expect(applyFix(src, 'doc-param-missing', "Add a <param> for 'item'")).toBe(
			"''' <summary>Adds an item.</summary>\r\n''' <param name=\"item\"></param>\r\nPublic Sub AddItem(ByVal item As String)\r\nEnd Sub\r\n",
		);
	});

	it('offers nothing but the suppression for an empty or unclosed tag', () => {
		const src = lines(
			"''' <summary>Adds an item.",
			"''' <param name=\"item\"></param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		for (const diagnostic of docDiagnostics(src)) {
			expect(fixTitles(src, diagnostic), diagnostic.message).toEqual([]);
		}
	});
});

describe('suppressing a doc comment finding', () => {
	const suppressed = (src: string): string[] =>
		filterDiagnosticsWithSuppressions(src, docDiagnostics(src)).diagnostics.map((d) => d.code);

	it('lets disable-next-member cover the doc comment, above it or below it', () => {
		const body = [
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"thing\">The item.</param>",
		];
		const decl = ['Public Sub AddItem(ByVal item As String)', 'End Sub'];
		const directive = "' @xlide-analysis-disable-next-member doc-param-unknown,doc-param-missing";
		expect(suppressed(lines(...body, ...decl))).toEqual(['doc-param-unknown', 'doc-param-missing']);
		expect(suppressed(lines(directive, ...body, ...decl))).toEqual([]);
		expect(suppressed(lines(...body, directive, ...decl))).toEqual([]);
	});

	it('suppresses a finding inside the block from the line above it, and the block still documents the Sub', () => {
		const src = lines(
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"thing\">The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'End Sub',
		);
		const unknown = docDiagnostics(src).find((d) => d.code === 'doc-param-unknown')!;
		const action = resolveDiagnosticCodeActions(src, {
			code: unknown.code,
			span: unknown.span,
			data: unknown.data,
			includeSuppressionAction: true,
		}).find((candidate) => candidate.title.startsWith('Suppress'))!;
		const edit = action.edits[0];
		const out = src.slice(0, edit.span.start) + edit.newText + src.slice(edit.span.end);
		expect(suppressed(out)).toEqual(['doc-param-missing']);
		expect(extractLeadingDoc(out, out.indexOf('Public Sub'))?.params.map((p) => p.name)).toEqual(['thing']);
	});
});
