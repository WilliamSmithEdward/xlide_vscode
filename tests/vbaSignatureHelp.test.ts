import { describe, it, expect } from 'vitest';
import { ProjectIndex, resolveSignatureHelp, SignatureHelpContext } from '../src/analyzer';

/** Resolves signature help with the caret at the `|` marker in `src`. */
function help(src: string, ctx: SignatureHelpContext = {}) {
	const offset = src.indexOf('|');
	if (offset < 0) {
		throw new Error('caret marker | not found');
	}
	const source = src.slice(0, offset) + src.slice(offset + 1);
	return resolveSignatureHelp(source, offset, ctx);
}

describe('signature help - host members', () => {
	it('offers Workbooks.Open with the active parameter at the start', () => {
		const info = help('Sub T()\nWorkbooks.Open(|\nEnd Sub');
		expect(info).toBeDefined();
		expect(info?.label.startsWith('Open(Filename As String')).toBe(true);
		expect(info?.label.endsWith(') As Workbook')).toBe(true);
		expect(info?.parameters[0].label).toBe('Filename As String');
		expect(info?.activeParameter).toBe(0);
	});

	it('advances the active parameter across commas', () => {
		const info = help('Sub T()\nWorkbooks.Open("a.xlsx", True, |\nEnd Sub');
		expect(info?.activeParameter).toBe(2);
		expect(info?.parameters[2].label).toBe('[ReadOnly As Variant]');
	});

	it('clamps the active parameter to the last for excess commas', () => {
		const many = '1,'.repeat(40);
		const info = help(`Sub T()\nWorkbooks.Open(${many}|\nEnd Sub`);
		expect(info).toBeDefined();
		const last = (info?.parameters.length ?? 0) - 1;
		expect(info?.activeParameter).toBe(last);
	});

	it('supports the parenless call statement form', () => {
		const info = help('Sub T()\nWorkbooks.Open |\nEnd Sub');
		expect(info?.label.startsWith('Open(Filename As String')).toBe(true);
		expect(info?.activeParameter).toBe(0);
	});

	it('tracks parameters in a parenless call statement', () => {
		const info = help('Sub T()\nWorkbooks.Open "a.xlsx", |\nEnd Sub');
		expect(info?.activeParameter).toBe(1);
	});

	it('resolves a member on a chained receiver', () => {
		const src = 'Sub T()\nActiveSheet.Range("A1").Offset(|\nEnd Sub';
		const info = help(src);
		expect(info?.label.startsWith('Offset(')).toBe(true);
		expect(info?.parameters[0].label).toBe('[RowOffset]');
	});

	it('includes generated reference docs for host member call tips', () => {
		const info = help('Sub T()\nWorkbooks.Open(|\nEnd Sub');
		expect(info?.documentation).toContain('Opens a workbook');
		expect(info?.parameters[0].documentation).toContain('file name');
	});

	it('offers generated no-argument host method signatures', () => {
		const info = help('Sub T()\nApplication.Calculate(|\nEnd Sub');
		expect(info?.label).toBe('Calculate()');
		expect(info?.documentation).toContain('Calculates all open workbooks');
		expect(info?.parameters).toEqual([]);
	});

	it('returns undefined for a member without a signature', () => {
		// Workbook.Name is a property with no transcribed signature.
		const info = help('Sub T()\nThisWorkbook.Name(|\nEnd Sub');
		expect(info).toBeUndefined();
	});

	it('does not offer call tips for host events', () => {
		const info = help('Sub T()\nThisWorkbook.AfterSave(|\nEnd Sub');
		expect(info).toBeUndefined();
	});
});

describe('signature help - runtime built-ins', () => {
	it('offers MsgBox', () => {
		const info = help('Sub T()\nx = MsgBox(|\nEnd Sub');
		expect(info?.label.startsWith('MsgBox(Prompt')).toBe(true);
		expect(info?.parameters[0].label).toBe('Prompt');
		expect(info?.activeParameter).toBe(0);
	});

	it('advances MsgBox to the Buttons parameter', () => {
		const info = help('Sub T()\nx = MsgBox("hi", |\nEnd Sub');
		expect(info?.activeParameter).toBe(1);
		expect(info?.parameters[1].label).toBe('[Buttons As VbMsgBoxStyle = vbOKOnly]');
	});

	it('offers Left with two parameters', () => {
		const info = help('Sub T()\ns = Left(|\nEnd Sub');
		expect(info?.label).toBe('Left(String, Length) As String');
		expect(info?.parameters.map((p) => p.label)).toEqual(['String', 'Length']);
	});

	it('does not offer DoEvents when it is used as an invalid explicit Call target', () => {
		expect(help('Sub T()\nCall DoEvents(|\nEnd Sub')).toBeUndefined();
		expect(help('Sub T()\nvalue = DoEvents(|\nEnd Sub')?.label).toBe(
			'DoEvents() As Integer',
		);
	});
});

describe('signature help - user procedures', () => {
	const src = [
		'Sub Greet(ByVal Name As String, Optional Loud As Boolean = False)',
		'End Sub',
		'Sub Caller()',
		'Greet(|',
		'End Sub',
	].join('\n');

	it('builds a signature from the parsed procedure', () => {
		const info = help(src);
		expect(info?.label).toBe(
			'Greet(Name As String, [Loud As Boolean = False])',
		);
	});

	it('exposes both parameters with optional bracketed', () => {
		const info = help(src);
		expect(info?.parameters.map((p) => p.label)).toEqual([
			'Name As String',
			'[Loud As Boolean = False]',
		]);
	});

	it('tracks the active parameter for a user procedure', () => {
		const after = src.replace('Greet(|', 'Greet "x", |');
		const info = help(after);
		expect(info?.activeParameter).toBe(1);
	});

	it('renders a Function return type', () => {
		const fn = [
			'Function Add(A As Long, B As Long) As Long',
			'End Function',
			'Sub C()',
			'x = Add(|',
			'End Sub',
		].join('\n');
		const info = help(fn);
		expect(info?.label).toBe('Add(A As Long, B As Long) As Long');
	});

	it('uses exported project procedures from other standard modules', () => {
		const info = help('Sub Caller()\n    InvoiceTotal(|\nEnd Sub\n', {
			projectProcedures: [
				{
					name: 'InvoiceTotal',
					moduleName: 'Helpers',
					kind: 'function',
					returnType: 'Currency',
					params: [
						{
							name: 'Subtotal',
							type: 'Currency',
							optional: false,
							paramArray: false,
							isArray: false,
						},
						{
							name: 'TaxRate',
							type: 'Double',
							optional: true,
							paramArray: false,
							isArray: false,
							defaultRaw: '0.08',
						},
					],
					doc: {
						summary: 'Calculates the invoice total.',
						params: [
							{ name: 'Subtotal', text: 'Pre-tax amount.' },
						],
						source: 'inline',
					},
				},
			],
		});
		expect(info?.label).toBe(
			'InvoiceTotal(Subtotal As Currency, [TaxRate As Double = 0.08]) As Currency',
		);
		expect(info?.parameters.map((p) => p.label)).toEqual([
			'Subtotal As Currency',
			'[TaxRate As Double = 0.08]',
		]);
		expect(info?.documentation).toContain('Calculates the invoice total.');
		expect(info?.parameters[0].documentation).toBe('Pre-tax amount.');
	});
});

describe('signature help - project class members', () => {
	const index = new ProjectIndex();
	index.setModule({
		moduleName: 'Person',
		moduleKind: 'class',
		source: [
			"''' <summary>Saves the person.</summary>",
			"''' <param name=\"Caption\">Caption text.</param>",
			'Public Sub Save(ByVal Caption As String, Optional Loud As Boolean)',
			'End Sub',
			'Public Function Manager(ByVal Depth As Long) As Person',
			'End Function',
		].join('\n'),
	});
	const projectClassMembers = index.projectClassMembers();

	it('uses source-backed project class method signatures', () => {
		const info = help('Sub T()\nDim p As Person\np.Save(|\nEnd Sub', {
			projectClassMembers,
		});
		expect(info?.label).toBe('Save(Caption As String, [Loud As Boolean])');
		expect(info?.parameters.map((p) => p.label)).toEqual([
			'Caption As String',
			'[Loud As Boolean]',
		]);
	});

	it('uses source-backed current class method signatures through Me', () => {
		const info = help('Sub T()\nMe.Save(|\nEnd Sub', {
			meProjectType: 'Person',
			projectClassMembers,
		});
		expect(info?.label).toBe('Save(Caption As String, [Loud As Boolean])');
	});

	it('tracks active parameters and return types for project class functions', () => {
		const info = help('Sub T()\nDim p As Person\nSet p = p.Manager(1, |\nEnd Sub', {
			projectClassMembers,
		});
		expect(info?.label).toBe('Manager(Depth As Long) As Person');
		expect(info?.activeParameter).toBe(0);
	});

	it('carries inline XML docs into project class member signature help', () => {
		const info = help('Sub T()\nDim p As Person\np.Save(|\nEnd Sub', {
			projectClassMembers,
		});
		expect(info?.documentation).toContain('Saves the person.');
		expect(info?.parameters[0].documentation).toBe('Caption text.');
	});
});

describe('signature help - negative cases', () => {
	it('returns undefined for a grouping paren', () => {
		const info = help('Sub T()\nx = (1 + |\nEnd Sub');
		expect(info).toBeUndefined();
	});

	it('returns undefined for an unknown callee', () => {
		const info = help('Sub T()\nNoSuchThing(|\nEnd Sub');
		expect(info).toBeUndefined();
	});

	it('returns undefined for an assignment, not a call', () => {
		const info = help('Sub T()\nx = |\nEnd Sub');
		expect(info).toBeUndefined();
	});

	it('does not fire for a file Open statement', () => {
		const info = help('Sub T()\nOpen "f" For Input As #1\nx = |\nEnd Sub');
		expect(info).toBeUndefined();
	});
});

describe('signature help - default parameter (ByVal omitted)', () => {
	it('marks a ByRef-less plain parameter without prefix', () => {
		const src = [
			'Sub P(A As Long)',
			'End Sub',
			'Sub C()',
			'P(|',
			'End Sub',
		].join('\n');
		const info = help(src);
		expect(info?.parameters[0].label).toBe('A As Long');
	});
});
