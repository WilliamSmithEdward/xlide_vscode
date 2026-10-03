// Diagnostics tests: a Boolean local given to a host property with limits
// (issue #630). Excel takes True where it refuses -1, and refuses False as
// it does 0. Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TARGETS = ['Font.Underline', 'Font.Size', 'HorizontalAlignment', 'Borders(xlEdgeBottom).LineStyle', 'ColumnWidth'];

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Boolean local as a host property value (issue #630)', () => {
	it('takes a Boolean holding True', () => {
		for (const target of TARGETS) {
			expect(errors(`Dim b As Boolean\n    b = True\n    ActiveSheet.Range("A1").${target} = b`), target).toEqual([]);
		}
	});

	it('still reports -1 in a Long, and a Boolean holding False where 0 is refused', () => {
		for (const target of TARGETS) {
			expect(errors(`Dim n As Long\n    n = -1\n    ActiveSheet.Range("A1").${target} = n`), target).toEqual(['host-property-value-out-of-range']);
		}
		expect(errors('Dim b As Boolean\n    b = False\n    ActiveSheet.Range("A1").Font.Size = b')).toEqual(['host-property-value-out-of-range']);
		expect(errors('Dim n As Integer\n    n = True\n    ActiveSheet.Range("A1").Font.Size = n')).toEqual(['host-property-value-out-of-range']);
	});
});
