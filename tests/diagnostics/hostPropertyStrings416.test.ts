// Diagnostics tests: Excel properties given a String that is no number
// (issue #416). Measured in Excel 16.0 64-bit (build 20430, 2026-10-02)
// through pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src, { host: 'excel' }).filter((diag) => diag.severity === 'error').map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a host property given a String', () => {
	it('raises for one that is no number', () => {
		for (const [target, error, takesBoolean] of [
			['Range("A1").Font.Bold', '1004', true], ['Range("A1").Font.Size', '1004', false], ['Range("A1").ColumnWidth', '1004', false],
			['Range("A1").RowHeight', '1004', false], ['Range("A1").HorizontalAlignment', '1004', false], ['Range("A1").WrapText', '1004', true],
			['ActiveWindow.Zoom', '1004', false], ['Application.ScreenUpdating', '13', true], ['Application.DisplayAlerts', '13', true],
			['Application.Calculation', '13', false], ['Range("A1").Interior.Color', '13', false], ['ActiveSheet.Tab.Color', '13', false],
		] as const) {
			for (const value of ['"abc"', '""']) {
				expect(errors(`${target} = ${value}`), `${target} = ${value}`).toEqual([expect.stringMatching(new RegExp(`^host-property-value-out-of-range: .*'${error}'`))]);
			}
			expect(errors(`${target} = "True"`), `${target} = "True"`).toHaveLength(takesBoolean ? 0 : 1);
		}
		expect(errors('Dim ws As Worksheet', 'Set ws = ActiveSheet', 'ws.Visible = "abc"')).toEqual([expect.stringContaining("'13'")]);
	});

	it('is quiet for a number in a String, and for Value, NumberFormat and Font.Name', () => {
		for (const line of [
			'Range("A1").Font.Size = "12"', 'Range("A1").ColumnWidth = "12"', 'Application.ScreenUpdating = "12"', 'ActiveWindow.Zoom = "12"',
			'Range("A1").Value = "abc"', 'Range("A1").NumberFormat = "abc"', 'Range("A1").Font.Name = "abc"', 'Range("A1").Value = ""',
		]) {
			expect(errors(line), line).toEqual([]);
		}
		expect(errors('Dim ws As Worksheet', 'Set ws = ActiveSheet', 'ws.Visible = "-1"')).toEqual([]);
	});
});
