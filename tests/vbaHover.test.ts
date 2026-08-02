import { describe, it, expect } from 'vitest';
import { resolveHover, HoverContext, ProjectIndex } from '../src/analyzer';

/** Offset of the first character of `marker` in `src`. */
function at(src: string, marker: string, within = 0): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + within;
}

function hover(src: string, marker: string, ctx: HoverContext = {}, within = 1) {
	return resolveHover(src, at(src, marker, within), ctx);
}

describe('hover - user symbols', () => {
	it('describes a Function with its signature, module and visibility', () => {
		const src =
			'Public Function GetCustomer(id As Long) As Customer\n' +
			'End Function\n';
		const info = resolveHover(src, src.indexOf('GetCustomer') + 2, {
			moduleName: 'CustomerApi',
		});
		expect(info).toBeDefined();
		expect(info?.signature).toBe(
			'Function GetCustomer(id As Long) As Customer',
		);
		expect(info?.details).toContain('Declared in Module: CustomerApi');
		expect(info?.details).toContain('Visibility: Public');
	});

	it('describes a Sub with no parameters', () => {
		const src = 'Private Sub DoWork()\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('DoWork') + 1, {
			moduleName: 'Module1',
		});
		expect(info?.signature).toBe('Sub DoWork()');
		expect(info?.details).toContain('Visibility: Private');
	});

	it('describes exported project procedures from other standard modules', () => {
		const src = 'Sub Caller()\n    total = InvoiceTotal(100)\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('InvoiceTotal') + 2, {
			moduleName: 'Caller',
			projectProcedures: [
				{
					name: 'InvoiceTotal',
					moduleName: 'Helpers',
					kind: 'function',
					params: [
						{
							name: 'Subtotal',
							type: 'Currency',
							optional: false,
							paramArray: false,
							isArray: false,
						},
					],
					returnType: 'Currency',
					visibility: 'Public',
					doc: {
						summary: 'Calculates the invoice total.',
						params: [],
						source: 'inline',
					},
				},
			],
		});
		expect(info?.signature).toBe('Function InvoiceTotal(Subtotal As Currency) As Currency');
		expect(info?.details).toContain('Declared in Module: Helpers');
		expect(info?.details).toContain('Visibility: Public');
		expect(info?.documentation).toContain('Calculates the invoice total.');
	});

	it('describes external Declare callables from the current module', () => {
		const src =
			'Private Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As Long\n' +
			'Sub Caller()\n    value = GetTickCount()\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('GetTickCount()') + 2, {
			moduleName: 'NativeApi',
		});

		expect(info?.signature).toBe(
			'Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As Long',
		);
		expect(info?.details).toContain('External declaration');
		expect(info?.details).toContain('Lib: kernel32');
		expect(info?.details).toContain('Alias: GetTickCount64');
	});

	it('describes exported project Declare callables from other standard modules', () => {
		const src = 'Sub Caller()\n    Sleep 100\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Sleep') + 2, {
			moduleName: 'Caller',
			projectProcedures: [
				{
					name: 'Sleep',
					moduleName: 'NativeApi',
					kind: 'sub',
					params: [
						{
							name: 'Milliseconds',
							type: 'LongPtr',
							optional: false,
							paramArray: false,
							isArray: false,
							byVal: true,
						},
					],
					external: true,
					ptrSafe: true,
					libName: 'kernel32',
					visibility: 'Public',
				},
			],
		});

		expect(info?.signature).toBe(
			'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)',
		);
		expect(info?.details).toContain('External declaration');
		expect(info?.details).toContain('Declared in Module: NativeApi');
	});

	it('describes a local variable from a usage inside the procedure', () => {
		const src =
			'Sub Test()\n' +
			'    Dim total As Double\n' +
			'    total = 1\n' +
			'End Sub\n';
		// Hover on the usage of `total` on the assignment line.
		const usage = src.lastIndexOf('total');
		const info = resolveHover(src, usage + 1, { moduleName: 'Module1' });
		expect(info?.signature).toBe('total As Double');
		expect(info?.details.some((d) => d.includes('Test'))).toBe(true);
	});

	it('preserves fixed-length String suffixes on variable hovers', () => {
		const src =
			'Sub Test()\n' +
			'    Dim buffer As String * 20\n' +
			'    buffer = "abc"\n' +
			'End Sub\n';
		const usage = src.lastIndexOf('buffer');
		const info = resolveHover(src, usage + 1, { moduleName: 'Module1' });
		expect(info?.signature).toBe('buffer As String * 20');
	});

	it('describes a parameter', () => {
		const src = 'Sub Greet(name As String)\n    Debug.Print name\nEnd Sub\n';
		const usage = src.lastIndexOf('name');
		const info = resolveHover(src, usage + 1, { moduleName: 'Module1' });
		expect(info?.signature).toBe('name As String');
		expect(info?.details.some((d) => d.includes('Greet'))).toBe(true);
	});

	it('describes a module-level constant with Const prefix', () => {
		const src = 'Public Const MAX As Long = 10\nSub T()\n    Debug.Print MAX\nEnd Sub\n';
		const usage = src.lastIndexOf('MAX');
		const info = resolveHover(src, usage + 1, { moduleName: 'Consts' });
		expect(info?.signature).toBe('Const MAX As Long = 10');
		expect(info?.details).toContain('Declared in Module: Consts');
	});

	it('describes an Enum and its members', () => {
		const src =
			'Public Enum Color\n' +
			'    Red\n' +
			'    Green\n' +
			'End Enum\n' +
			'Sub T()\n' +
			'    Dim c As Color\n' +
			'    c = Green\n' +
			'End Sub\n';
		const enumHit = resolveHover(src, src.indexOf('Enum Color') + 6, {
			moduleName: 'Palette',
		});
		expect(enumHit?.signature).toBe('Enum Color');

		const memberHit = resolveHover(src, src.lastIndexOf('Green') + 1, {
			moduleName: 'Palette',
		});
		expect(memberHit?.signature).toBe('Green');
		expect(memberHit?.details.some((d) => d.includes('Color'))).toBe(true);
	});
});

describe('hover - standard module-qualified symbols', () => {
	it('describes the module qualifier and qualified members', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Finance',
			moduleKind: 'standard',
			source: [
				"''' <summary>Finance helper module.</summary>",
				"''' <remarks>Shared workbook calculations.</remarks>",
				'',
				"''' <summary>Calculates the invoice total.</summary>",
				'Public Function InvoiceTotal(ByVal Subtotal As Currency) As Currency',
				'End Function',
				'Public Const DefaultTaxRate As Double = 0.08',
				'Public Enum SharedMode',
				'    SharedOnly',
				'End Enum',
			].join('\n'),
		});
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });
		const ctx = { projectClassMembers: index.projectMemberSurfaces('Caller') };
		const src =
			'Sub T()\n' +
			'    total = Finance.InvoiceTotal(100)\n' +
			'    rate = Finance.DefaultTaxRate\n' +
			'    mode = Finance.SharedOnly\n' +
			'End Sub\n';

		const moduleHover = resolveHover(src, src.indexOf('Finance.InvoiceTotal') + 2, ctx);
		expect(moduleHover?.signature).toBe('Module Finance');
		expect(moduleHover?.details).toContain('Standard module');
		expect(moduleHover?.documentation).toContain('Finance helper module.');
		expect(moduleHover?.documentation).toContain('Shared workbook calculations.');

		const functionHover = resolveHover(src, src.indexOf('InvoiceTotal') + 2, ctx);
		expect(functionHover?.signature).toBe(
			'Finance.InvoiceTotal(Subtotal As Currency) As Currency',
		);
		expect(functionHover?.details).toContain('Finance method');
		expect(functionHover?.documentation).toContain('Calculates the invoice total.');

		const constantHover = resolveHover(src, src.indexOf('DefaultTaxRate') + 2, ctx);
		expect(constantHover?.signature).toBe('Finance.DefaultTaxRate As Double');
		expect(constantHover?.details).toContain('Finance property');

		const enumMemberHover = resolveHover(src, src.indexOf('SharedOnly') + 2, ctx);
		expect(enumMemberHover?.signature).toBe('Finance.SharedOnly As SharedMode');
		expect(enumMemberHover?.details).toContain('Finance property');
	});

	it('describes a class used as a bare receiver, and its UserForm sibling', () => {
		// A class with a predeclared instance is addressed by its own name, the way a
		// factory-style module is: `Inventory.DataView(rows)`. Hovering the class name in that
		// expression-root position describes the class, exactly as hovering it in a type
		// position would.
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Inventory',
			moduleKind: 'class',
			source: [
				"''' <summary>Sheet-backed inventory window.</summary>",
				'',
				'Public Function DataView(ByVal rows As Long) As Inventory',
				'End Function',
			].join('\n'),
		});
		index.setModule({ moduleName: 'EntryForm', moduleKind: 'userform', source: '' });
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });
		const ctx = { projectClassMembers: index.projectMemberSurfaces('Caller') };
		const src =
			'Sub T()\n' +
			'    Set view = Inventory.DataView(10)\n' +
			'    EntryForm.Show\n' +
			'End Sub\n';

		const classHover = resolveHover(src, src.indexOf('Inventory.DataView') + 2, ctx);
		expect(classHover?.signature).toBe('Class Inventory');
		expect(classHover?.details).toContain('Class module');
		expect(classHover?.documentation).toContain('Sheet-backed inventory window.');

		const formHover = resolveHover(src, src.indexOf('EntryForm.Show') + 2, ctx);
		expect(formHover?.signature).toBe('UserForm EntryForm');
		expect(formHover?.details).toContain('UserForm');
	});
});

describe('hover - host symbols', () => {
	it('describes a host global', () => {
		const src = 'Sub T()\n    ThisWorkbook.Save\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('ThisWorkbook') + 2, {});
		expect(info?.signature).toBe('ThisWorkbook As Workbook');
		expect(info?.details).toContain('Excel host global');
	});

	it('describes a host member after a dot', () => {
		const src = 'Sub T()\n    ThisWorkbook.Worksheets\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Worksheets') + 2, {});
		expect(info?.signature.startsWith('Workbook.Worksheets')).toBe(true);
		expect(info?.details[0]).toMatch(/Excel host (property|method)/);
	});

	it('describes a leading-dot host member inside With', () => {
		const src = 'Sub T()\n    With Range("A1")\n        .Value\n    End With\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Value') + 2, {});
		expect(info?.signature.startsWith('Range.Value')).toBe(true);
		expect(info?.details[0]).toMatch(/Excel host (property|method)/);
	});

	it('describes a source-backed current class member after Me', () => {
		const src = 'Sub T()\n    Me.Save\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Save') + 2, {
			meProjectType: 'Person',
			projectClassMembers: [
				{
					name: 'Person',
					kind: 'class',
					moduleName: 'Person',
					members: [{ name: 'Save', kind: 'method', moduleName: 'Person' }],
				},
			],
		});
		expect(info?.signature).toBe('Person.Save()');
		expect(info?.details).toContain('Person method');
	});

	it('describes a source-backed UDT field after a dot', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: 'Public Type TPoint\n    X As Long\n    Label As String * 12\nEnd Type\n',
		});
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });
		const src = 'Sub T()\n    Dim p As TPoint\n    p.X\n    p.Label\nEnd Sub\n';
		const x = resolveHover(src, src.indexOf('X') + 1, {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(x?.signature).toBe('TPoint.X As Long');
		expect(x?.details).toContain('TPoint property');

		const label = resolveHover(src, src.indexOf('Label') + 1, {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(label?.signature).toBe('TPoint.Label As String * 12');
		expect(label?.details).toContain('TPoint property');
	});

	it('includes generated reference docs on promoted host member hovers', () => {
		const src = 'Sub T()\n    Application.Calculate\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Calculate') + 2, {});
		expect(info?.signature).toBe('Application.Calculate()');
		expect(info?.documentation).toContain('Calculates all open workbooks');
	});

	it('describes a worksheet code name', () => {
		const src = 'Sub T()\n    Sheet1.Range("A1")\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Sheet1') + 2, {
			codeNames: { sheet1: 'Excel.Worksheet' },
		});
		expect(info?.signature).toBe('Sheet1 As Worksheet');
		expect(info?.details).toContain('Worksheet code name');
	});

	it('does not guess unknown members', () => {
		const src = 'Sub T()\n    ThisWorkbook.NotARealMember\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('NotARealMember') + 2, {});
		expect(info).toBeUndefined();
	});
});

describe('hover - built-in VBA runtime', () => {
	it('describes MsgBox with its signature', () => {
		const src = 'Sub T()\n    MsgBox "hi"\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('MsgBox') + 2, {
			moduleName: 'Module1',
		});
		expect(info?.signature.startsWith('MsgBox(Prompt')).toBe(true);
		expect(info?.signature).toContain('As VbMsgBoxResult');
		expect(info?.details).toContain('VBA runtime function');
	});

	it('describes the intrinsic Err object and its members', () => {
		const src = 'Sub T()\n    Err.Raise vbObjectError + 1, "M", "boom"\nEnd Sub\n';
		const err = resolveHover(src, src.indexOf('Err') + 1, {
			moduleName: 'M',
		});
		expect(err?.signature).toBe('Err As ErrObject');
		expect(err?.details).toContain('VBA runtime object');

		const raise = resolveHover(src, src.indexOf('Raise') + 1, {
			moduleName: 'M',
		});
		expect(raise?.signature).toContain('Err.Raise(Number As Long');
		expect(raise?.details).toContain('VBA runtime method');
	});

	it('describes common string and conversion functions', () => {
		const src = 'Sub T()\n    Dim s As String\n    s = Left(CStr(1), 2)\nEnd Sub\n';
		const left = resolveHover(src, src.indexOf('Left(') + 1, {
			moduleName: 'M',
		});
		expect(left?.signature).toBe('Left(String, Length) As String');
		const cstr = resolveHover(src, src.indexOf('CStr(') + 1, { moduleName: 'M' });
		expect(cstr?.signature).toBe('CStr(Expression) As String');
	});

	it('describes built-in VBA and Excel constants', () => {
		const src =
			'Sub T()\n' +
			'    MsgBox "hi", vbOKOnly\n' +
			'    ActiveSheet.Range("A1").End(xlUp).Select\n' +
			'    ActiveSheet.Shapes(1).Line.DashStyle = msoLineDash\n' +
			'End Sub\n';
		const vb = resolveHover(src, src.indexOf('vbOKOnly') + 2, {
			moduleName: 'M',
		});
		expect(vb?.signature).toBe('Const vbOKOnly As VbMsgBoxStyle = 0');
		expect(vb?.details).toContain('VBA runtime constant');

		const xl = resolveHover(src, src.indexOf('xlUp') + 2, {
			moduleName: 'M',
		});
		expect(xl?.signature).toBe('Const xlUp As XlDirection = -4162');
		expect(xl?.details).toContain('Excel/Office constant');

		const mso = resolveHover(src, src.indexOf('msoLineDash') + 2, {
			moduleName: 'M',
		});
		expect(mso?.signature).toBe('Const msoLineDash As MsoLineDashStyle = 4');
		expect(mso?.details).toContain('Excel/Office constant');
	});

	it('lets a user symbol shadow a built-in of the same name', () => {
		// A user Function named Format must win over the runtime Format().
		const src =
			'Public Function Format() As String\n' +
			'    Format = "x"\n' +
			'End Function\n';
		const info = resolveHover(src, src.indexOf('Function Format') + 10, {
			moduleName: 'M',
		});
		expect(info?.signature).toBe('Function Format() As String');
		expect(info?.details).toContain('Declared in Module: M');
	});

	it('describes an intrinsic type in an As clause as a type, not a runtime function', () => {
		const src = 'Sub T()\n    Dim s As String\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('As String') + 4, {
			moduleName: 'M',
		});
		expect(info?.signature).toBe('String');
		expect(info?.details).toContain('VBA primitive type');
	});

	it('does not describe an intrinsic type name outside a type position', () => {
		const src = 'Sub T()\n    Debug.Print String\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('String') + 2, {
			moduleName: 'M',
		});
		expect(info).toBeUndefined();
	});
});

describe('hover - type names', () => {
	it('describes project and host types in declaration type positions', () => {
		const src =
			'Sub T()\n' +
			'    Dim p As Person\n' +
			'    Dim ws As Worksheet\n' +
			'End Sub\n';
		const person = resolveHover(src, src.indexOf('Person') + 2, {
			moduleName: 'M',
			projectTypes: [{ name: 'Person', kind: 'class' }],
		});
		expect(person?.signature).toBe('Class Person');
		expect(person?.details).toContain('Class');

		const worksheet = resolveHover(src, src.indexOf('Worksheet') + 2, {
			moduleName: 'M',
		});
		expect(worksheet?.signature).toBe('Worksheet');
		expect(worksheet?.details).toContain('Excel host type');
	});

	it('shows documentation for documented project type names', () => {
		const src = 'Sub T()\n    Set p = New Person\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Person') + 2, {
			moduleName: 'M',
			projectTypes: [
				{
					name: 'Person',
					kind: 'class',
					doc: {
						summary: 'Represents a person.',
						params: [],
						source: 'inline',
					},
				},
			],
		});
		expect(info?.signature).toBe('Class Person');
		expect(info?.documentation).toContain('Represents a person.');
	});

	it('describes qualified project type names in declaration positions', () => {
		const src = 'Sub T()\n    Dim p As Geometry.TPoint\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('TPoint') + 2, {
			moduleName: 'M',
			projectTypes: [
				{
					name: 'TPoint',
					kind: 'userType',
					moduleName: 'Geometry',
					doc: {
						summary: 'A shared coordinate payload.',
						params: [],
						source: 'inline',
					},
				},
				{ name: 'TPoint', kind: 'userType', moduleName: 'OtherGeometry' },
			],
		});

		expect(info?.signature).toBe('Type TPoint');
		expect(info?.details).toContain('User type');
		expect(info?.documentation).toContain('A shared coordinate payload.');
	});

	it('keeps colliding project type hovers generic', () => {
		const src = 'Sub T()\n    Dim state As Status\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Status') + 2, {
			moduleName: 'M',
			projectTypes: [
				{ name: 'Status', kind: 'class' },
				{ name: 'Status', kind: 'enum' },
			],
		});
		expect(info?.signature).toBe('Status');
		expect(info?.details).toContain('Ambiguous project type');
	});
});

describe('hover - edge cases', () => {
	it('returns undefined off any identifier', () => {
		const src = 'Sub T()\n    Dim x As Long\nEnd Sub\n';
		// Offset inside leading whitespace.
		expect(resolveHover(src, src.indexOf('Dim') - 1, {})).toBeUndefined();
	});

	it('does not merge a dangling dot on a previous line', () => {
		// A receiver on the prior line must not bind to an identifier below it.
		const src = 'Sub T()\n    wb.\n    Sheet1\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Sheet1') + 1, {
			codeNames: { sheet1: 'Excel.Worksheet' },
		});
		expect(info?.signature).toBe('Sheet1 As Worksheet');
	});
});
