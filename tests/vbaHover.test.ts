import { describe, it, expect } from 'vitest';
import { resolveHover, HoverContext } from '../src/analyzer';

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

	it('describes common string and conversion functions', () => {
		const src = 'Sub T()\n    Dim s As String\n    s = Left(CStr(1), 2)\nEnd Sub\n';
		const left = resolveHover(src, src.indexOf('Left(') + 1, {
			moduleName: 'M',
		});
		expect(left?.signature).toBe('Left(String, Length) As String');
		const cstr = resolveHover(src, src.indexOf('CStr(') + 1, { moduleName: 'M' });
		expect(cstr?.signature).toBe('CStr(Expression) As String');
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
			projectTypes: [{ name: 'Person', kind: 'class', moduleName: 'Person' }],
		});
		expect(person?.signature).toBe('Class Person');
		expect(person?.details).toContain('Class');

		const worksheet = resolveHover(src, src.indexOf('Worksheet') + 2, {
			moduleName: 'M',
		});
		expect(worksheet?.signature).toBe('Worksheet');
		expect(worksheet?.details).toContain('Excel host type');
	});

	it('keeps colliding project type hovers generic', () => {
		const src = 'Sub T()\n    Dim state As Status\nEnd Sub\n';
		const info = resolveHover(src, src.indexOf('Status') + 2, {
			moduleName: 'M',
			projectTypes: [
				{ name: 'Status', kind: 'class', moduleName: 'StatusClass' },
				{ name: 'Status', kind: 'enum', moduleName: 'SharedTypes' },
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
