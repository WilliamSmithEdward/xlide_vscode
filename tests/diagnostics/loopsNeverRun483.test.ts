// Diagnostics tests: loops that run no pass (issue #483, leftovers of
// #273). Each sample was measured through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const findings = (src: string): string[] => analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code);

describe('loops that run no pass (issue #483)', () => {
	it('skips a loop until an object never set is Nothing, or while it is not', () => {
		const bodies = [
			['Dim c As Collection', 'Do Until c Is Nothing', '    Main = c.Count', '    Set c = Nothing', 'Loop'],
			['Dim c As Collection', 'While Not c Is Nothing', '    Main = c.Count', '    Set c = Nothing', 'Wend'],
			['Dim c As Collection', 'Set c = New Collection', 'Set c = Nothing', 'Do While Not c Is Nothing', '    Main = c.Count', 'Loop'],
		];
		for (const body of bodies) {
			expect(findings(wrap(...body)), body.join(' / ')).toEqual([]);
		}
	});

	it('skips For Each over Array() in a local and over an empty Collection', () => {
		const bodies = [
			['Dim d As Long, x As Variant, v As Variant', 'v = Array()', 'For Each x In v', '    Main = 10 / d', 'Next'],
			['Dim c As Collection, d As Long, x As Variant', 'Set c = New Collection', 'For Each x In c', '    Main = 10 / d', 'Next'],
			['Dim e As New Collection, c As Collection, x As Variant', 'For Each x In e', '    Main = c.Count', 'Next'],
		];
		for (const body of bodies) {
			expect(findings(wrap(...body)), body.join(' / ')).toEqual([]);
		}
	});

	it('still reports a loop that runs, once the Collection has an item or the object is set', () => {
		const bodies = [
			['Dim c As Collection, d As Long, x As Variant', 'Set c = New Collection', 'c.Add 1', 'For Each x In c', '    Main = 10 / d', 'Next'],
			['Dim c As Collection, d As Long', 'Set c = New Collection', 'Do Until c Is Nothing', '    Main = 10 / d', 'Loop'],
			['Dim e As New Collection, d As Long, x As Variant', 'e.Add 1', 'For Each x In e', '    Main = 10 / d', 'Next'],
			['Dim c As Collection, d As Long, x As Variant', 'Set c = New Collection', 'If d = 0 Then c.Add 1', 'For Each x In c', '    Main = 10 / d', 'Next'],
			['Dim c As Collection, d As Long, k As Long, x As Variant', 'Set c = New Collection', 'For k = 1 To 2', '    c.Add k', 'Next', 'For Each x In c', '    Main = 10 / d', 'Next'],
		];
		for (const body of bodies) {
			expect(byCode(analyzeModule(wrap(...body)), 'division-by-zero'), body.join(' / ')).toHaveLength(1);
		}
	});
});
