import { describe, it, expect } from 'vitest';
import {
	callableCompletionShouldInsertParens,
	getHostGlobals,
	resolveIdentifierCompletions,
	IdentifierCompletionContext,
} from '../src/analyzer';

/** Offset just after `marker` in `src`. */
function at(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function names(
	src: string,
	marker: string,
	ctx: IdentifierCompletionContext = {},
): string[] {
	return resolveIdentifierCompletions(src, at(src, marker), ctx).map((c) => c.name);
}

describe('host globals (canonical casing)', () => {
	it('exposes canonical-cased global identifiers', () => {
		const got = getHostGlobals().map((g) => g.name);
		expect(got).toContain('ThisWorkbook');
		expect(got).toContain('ActiveSheet');
		expect(got).toContain('Application');
		expect(got).toContain('ActiveWorkbook');
		expect(got).toContain('Selection');
		expect(got).toContain('ActiveCell');
	});
});

describe('identifier completion - host globals', () => {
	it('offers globals at an empty statement position', () => {
		const src = 'Sub Test()\n    \nEnd Sub\n';
		const got = names(src, '    ');
		expect(got).toContain('ThisWorkbook');
		expect(got).toContain('ActiveSheet');
		expect(got).toContain('Application');
	});

	it('filters globals by the typed prefix', () => {
		const src = 'Sub Test()\n    Th\nEnd Sub\n';
		const got = names(src, '    Th');
		expect(got).toContain('ThisWorkbook');
		expect(got).not.toContain('Application');
	});

	it('is case-insensitive on the typed prefix', () => {
		const src = 'Sub Test()\n    activ\nEnd Sub\n';
		const got = names(src, '    activ');
		expect(got).toContain('ActiveSheet');
		expect(got).toContain('ActiveWorkbook');
		expect(got).toContain('ActiveCell');
	});

	it('offers globals on the right-hand side of an assignment', () => {
		const src = 'Sub Test()\n    Set x = Th\nEnd Sub\n';
		const got = names(src, '= Th');
		expect(got).toContain('ThisWorkbook');
	});
});

describe('identifier completion - code names', () => {
	it('offers worksheet/document code names', () => {
		const src = 'Sub Test()\n    She\nEnd Sub\n';
		const ctx = { codeNames: ['Sheet1', 'Sheet3', 'MySheet'] };
		const got = names(src, '    She', ctx);
		expect(got).toContain('Sheet1');
		expect(got).toContain('Sheet3');
		expect(got).not.toContain('MySheet');
	});

	it('does not duplicate ThisWorkbook between globals and code names', () => {
		const src = 'Sub Test()\n    This\nEnd Sub\n';
		const ctx = { codeNames: ['ThisWorkbook'] };
		const got = names(src, '    This', ctx);
		expect(got.filter((n) => n.toLowerCase() === 'thisworkbook')).toHaveLength(1);
	});
});

describe('identifier completion - in-scope declarations', () => {
	it('offers locals, parameters and module variables', () => {
		const src =
			'Public gCount As Long\n' +
			'Sub Test(ByVal arg As String)\n' +
			'    Dim ws As Worksheet\n' +
			'    w\n' +
			'End Sub\n';
		const got = names(src, '    w');
		expect(got).toContain('ws');
		const all = names(src, '    w'.replace('w', ''));
		expect(all).toContain('gCount');
		expect(all).toContain('arg');
		expect(all).toContain('ws');
	});

	it('offers other procedures in the module', () => {
		const src =
			'Sub Helper()\nEnd Sub\n' +
			'Sub Test()\n    Hel\nEnd Sub\n';
		const got = names(src, '    Hel');
		expect(got).toContain('Helper');
	});

	it('offers same-module Declare callables with full declaration details', () => {
		const src =
			'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)\n' +
			'Sub Test()\n    Sle\nEnd Sub\n';
		const got = resolveIdentifierCompletions(src, at(src, '    Sle'));
		const item = got.find((c) => c.name === 'Sleep');

		expect(item?.kind).toBe('procedure');
		expect(item?.detail).toBe(
			'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)',
		);
	});

	it('offers exported procedures from other standard modules', () => {
		const src = 'Sub Test()\n    my\nEnd Sub\n';
		const got = resolveIdentifierCompletions(src, at(src, '    my'), {
			moduleName: 'Module2',
			projectProcedures: [
				{
					name: 'mySub',
					moduleName: 'TestModule123',
					kind: 'sub',
					params: [],
				},
				{
					name: 'myFunction',
					moduleName: 'TestModule123',
					kind: 'function',
					params: [],
					returnType: 'String',
					doc: {
						summary: 'Returns a test value.',
						params: [],
						source: 'inline',
					},
				},
			],
		});

		expect(got.find((c) => c.name === 'mySub')?.detail).toBe(
			'Sub mySub() in TestModule123',
		);
		expect(got.find((c) => c.name === 'myFunction')?.detail).toBe(
			'Function myFunction() As String in TestModule123',
		);
		expect(got.find((c) => c.name === 'myFunction')?.documentation).toContain(
			'Returns a test value.',
		);
	});

	it('includes inline documentation for documented procedures', () => {
		const src =
			"''' <summary>Calculates the invoice total after tax.</summary>\n" +
			"''' <param name=\"Subtotal\">The pre-tax invoice amount.</param>\n" +
			"''' <returns>The subtotal plus calculated tax.</returns>\n" +
			'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency\n' +
			'End Function\n' +
			'Sub Test()\n    invoi\nEnd Sub\n';
		const got = resolveIdentifierCompletions(src, at(src, '    invoi'));
		const item = got.find((c) => c.name === 'InvoiceTotal');
		expect(item?.documentation).toContain('Calculates the invoice total after tax.');
		expect(item?.documentation).toContain('`Subtotal`: The pre-tax invoice amount.');
		expect(item?.documentation).toContain('**Returns:** The subtotal plus calculated tax.');
	});

	it('treats the enclosing Function name as its return variable', () => {
		const src =
			'Function myFunction() As String\n' +
			'    my\n' +
			'End Function\n';
		const got = resolveIdentifierCompletions(src, at(src, '    my'));
		const item = got.find((c) => c.name === 'myFunction');

		expect(item?.kind).toBe('variable');
		expect(item?.detail).toBe('Function return As String');
	});

	it('still treats other Function names as callable procedures', () => {
		const src =
			'Function Helper() As String\n' +
			'End Function\n' +
			'Function myFunction() As String\n' +
			'    Hel\n' +
			'End Function\n';
		const got = resolveIdentifierCompletions(src, at(src, '    Hel'));
		const item = got.find((c) => c.name === 'Helper');

		expect(item?.kind).toBe('procedure');
	});

	it('offers enum members by bare name', () => {
		const src =
			'Public Enum Color\n    Red\n    Green\nEnd Enum\n' +
			'Sub Test()\n    Gr\nEnd Sub\n';
		const got = names(src, '    Gr');
		expect(got).toContain('Green');
	});

	it('does not leak locals from another procedure', () => {
		const src =
			'Sub A()\n    Dim onlyInA As Long\nEnd Sub\n' +
			'Sub B()\n    o\nEnd Sub\n';
		const got = names(src, 'Sub B()\n    o');
		expect(got).not.toContain('onlyInA');
	});
});

describe('callable completion insertion contexts', () => {
	it('does not add parens for standalone VBA call statements', () => {
		const src = 'Sub T()\n    mySu\nEnd Sub\n';
		expect(callableCompletionShouldInsertParens(src, at(src, '    mySu'))).toBe(false);
	});

	it('does not carry expression context from the previous statement', () => {
		const src =
			'Sub mySub()\nEnd Sub\n' +
			'Function T() As String\n' +
			'    T = "hello world!"\n' +
			'    mySub\n' +
			'End Function\n';
		expect(callableCompletionShouldInsertParens(src, at(src, '    mySub'))).toBe(false);
	});

	it('adds parens for explicit Call statements', () => {
		const src = 'Sub T()\n    Call mySu\nEnd Sub\n';
		expect(callableCompletionShouldInsertParens(src, at(src, 'Call mySu'))).toBe(true);
	});

	it('adds parens for explicit Call member statements', () => {
		const src = 'Sub T()\n    Call Application.Calc\nEnd Sub\n';
		expect(callableCompletionShouldInsertParens(src, at(src, 'Application.Calc'))).toBe(true);
	});

	it('adds parens for expression contexts', () => {
		const src = 'Sub T()\n    value = Lef\nEnd Sub\n';
		expect(callableCompletionShouldInsertParens(src, at(src, '= Lef'))).toBe(true);
	});

	it('does not add parens for standalone member call statements', () => {
		const src = 'Sub T()\n    Application.Calc\nEnd Sub\n';
		expect(callableCompletionShouldInsertParens(src, at(src, 'Application.Calc'))).toBe(
			false,
		);
	});
});

describe('identifier completion - suppressed positions', () => {
	it('returns nothing after a member-access dot', () => {
		const src = 'Sub Test()\n    ws.Ra\nEnd Sub\n';
		expect(names(src, 'ws.Ra')).toEqual([]);
	});

	it('returns nothing in a type position (after As)', () => {
		const src = 'Sub Test()\n    Dim ws As Wor\nEnd Sub\n';
		expect(names(src, 'As Wor')).toEqual([]);
	});

	it('returns nothing in a declaration-name position (after Dim)', () => {
		const src = 'Sub Test()\n    Dim ne\nEnd Sub\n';
		expect(names(src, 'Dim ne')).toEqual([]);
	});

	it('returns nothing after Public/Private', () => {
		const src = 'Sub Test()\nEnd Sub\nPublic gv';
		expect(names(src, 'Public gv')).toEqual([]);
	});

	it('can be told to omit host globals', () => {
		const src = 'Sub Test()\n    Th\nEnd Sub\n';
		const got = names(src, '    Th', { includeGlobals: false });
		expect(got).not.toContain('ThisWorkbook');
	});
});
