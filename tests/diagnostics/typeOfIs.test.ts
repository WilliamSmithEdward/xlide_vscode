// Diagnostics tests: typeof-is-always-false rule (expression-binder Slice 4).
//
// `TypeOf x Is Y` is flagged only when x's declared object type can never be a Y.
// Verified against MS-VBAL 5.6 (TypeOf...Is) plus the shared object tables.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, spanText } from '../helpers/diagnostics';

const CODE = 'typeof-is-always-false';

function hits(src: string) {
	return byCode(analyzeModule(src), CODE);
}

function wrap(body: string): string {
	return `Sub T()\n${body}\nEnd Sub\n`;
}

describe('analyzeModule - typeof-is-always-false', () => {
	it('flags an If condition where the operand can never be the target type', () => {
		const src = wrap('    Dim wb As Workbook\n    If TypeOf wb Is Worksheet Then\n        Debug.Print 1\n    End If');
		const found = hits(src);
		expect(found).toHaveLength(1);
		expect(found[0].severity).toBe('warning');
		expect(spanText(src, found[0])).toBe('TypeOf wb Is Worksheet');
	});

	it('flags a TypeOf used outside an If condition (assignment RHS)', () => {
		const src = wrap('    Dim x As Range\n    Dim b As Boolean\n    b = (TypeOf x Is Worksheet)');
		expect(hits(src)).toHaveLength(1);
	});

	it('stays quiet when operand and target are the same type', () => {
		expect(hits(wrap('    Dim ws As Worksheet\n    If TypeOf ws Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for an Object or Variant operand (could hold anything)', () => {
		expect(hits(wrap('    Dim o As Object\n    If TypeOf o Is Worksheet Then\n    End If'))).toHaveLength(0);
		expect(hits(wrap('    Dim v As Variant\n    If TypeOf v Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for `Is Object` (always True, not False)', () => {
		expect(hits(wrap('    Dim wb As Workbook\n    If TypeOf wb Is Object Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet when the operand type is unknown', () => {
		expect(hits(wrap('    Dim x As FooUnknownType\n    If TypeOf x Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a scalar operand', () => {
		expect(hits(wrap('    Dim n As Long\n    If TypeOf n Is Worksheet Then\n    End If'))).toHaveLength(0);
	});

	it('stays quiet for a non-identifier operand (not yet modeled)', () => {
		// Member-access operand: the operand type is not resolved in this slice.
		expect(hits(wrap('    Dim wb As Workbook\n    If TypeOf wb.ActiveSheet Is Workbook Then\n    End If'))).toHaveLength(0);
	});
});
