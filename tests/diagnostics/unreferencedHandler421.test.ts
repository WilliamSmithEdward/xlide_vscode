// Diagnostics tests: a handler below Exit Function that nothing jumps to, as
// when its On Error line is commented out, never runs (issue #421). Measured
// in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nPublic Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

const RULES = ['resume-without-error', 'division-by-zero'];

function found(src: string): string[] {
	const diagnostics = analyzeModule(src);
	return RULES.flatMap((code) => byCode(diagnostics, code).map(() => code));
}

describe('a handler nothing jumps to', () => {
	it('raises nothing below Exit Function', () => {
		for (const lines of [
			["' On Error GoTo EH", 'Main = 1', 'Exit Function', 'EH:', 'Main = Err.Description', 'Resume Next'],
			["' On Error GoTo EH", 'Main = 1', 'ExitHere:', 'Exit Function', 'EH:', 'Resume ExitHere'],
			["' On Error GoTo EH", 'Main = 1', 'Exit Function', 'EH:', 'Resume'],
			['On Error GoTo 0', 'Main = 1', 'Exit Function', 'EH:', 'Resume Next'],
			['Dim d As Long', "' On Error GoTo EH", 'Main = 1', 'Exit Function', 'EH:', 'Main = 10 / d'],
		]) {
			expect(found(source(...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('is still judged where a jump reaches it, or in the main line', () => {
		for (const [lines, code] of [
			[['Main = 1', 'Resume Next'], 'resume-without-error'],
			[['Main = 1', 'GoTo EH', 'Exit Function', 'EH:', 'Resume Next'], 'resume-without-error'],
			[['Dim d As Long', 'Main = 1', 'GoTo EH', 'Exit Function', 'EH:', 'Main = 10 / d'], 'division-by-zero'],
		] as const) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), code), code, {});
		}
	});
});
