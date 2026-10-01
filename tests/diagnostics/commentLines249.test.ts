// Diagnostics tests: a comment line is no statement to the flow rules (issue
// #249). Each layout was measured in Excel 16.0 (build 20326, 2026-10-01) and
// runs without an error.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.join('\n')}\nEnd Function\n`;
}

function codes(src: string): string[] {
	return analyzeModule(src).map((diag) => diag.code);
}

describe('comment lines', () => {
	it.each([
		['the error-handler banner', ['    On Error GoTo Handler', '    Main = 1', '    Exit Function', "' ---- error handler ----", 'Handler:', '    Main = Err.Number']],
		['an indented banner', ['    On Error GoTo Handler', '    Main = 1', '    Exit Function', "    ' ---- error handler ----", 'Handler:', '    Main = Err.Number']],
		['a Rem banner', ['    On Error GoTo Handler', '    Main = 1', '    Exit Function', 'Rem ---- error handler ----', 'Handler:', '    Main = Err.Number']],
		['the last line', ['    Main = 1', '    Exit Function', "' the end"]],
		['a comment GoTo skips', ['    GoTo Skip', "    ' skipped", 'Skip:', '    Main = 3']],
		['a comment above a GoSub target', ['    GoSub Work', '    Exit Function', "    ' the subroutine", 'Work:', '    Main = 2', '    Return']],
		['a Rem and two comments above it', ['    GoSub Work', '    Exit Function', '    Rem one', "    ' two", 'Work:', '    Main = 2', '    Return']],
		['a comment after Return', ['    GoSub Work', '    Main = Main + 1', '    Exit Function', 'Work:', '    Main = 2', '    Return', "    ' after return"]],
	])('are not unreachable code or a fall into a GoSub target: %s', (_, lines) => {
		expect(codes(source(...lines))).toEqual([]);
	});

	it('leave a real statement after Exit unreachable', () => {
		const src = source('    Main = 1', '    Exit Function', "    ' note", '    Main = 2');
		expectDiagnostic(src, byCode(analyzeModule(src), 'unreachable-code'), 'unreachable-code', { span: 'Main = 2', message: "after 'Exit Function'" });
	});

	it('hide nothing a handler rule reads (issue199_01)', () => {
		const src = source('    Dim x As Double', '    On Error GoTo Handler', "    ' raise", '    Err.Raise 5', '    Exit Function', 'Handler:', "    ' keep going", '    On Error Resume Next', "    ' divide", '    x = 1 / 0', '    Main = x');
		expect(codes(src)).toContain('division-by-zero');
	});

	it('still separate a member Attribute line from its header, as before', () => {
		const attribute = 'Attribute Main.VB_Description = "metadata"';
		expect(codes(`Option Explicit\nPublic Sub Main()\n${attribute}\n    Debug.Print 1\nEnd Sub\n`)).not.toContain('module-declaration-in-procedure');
		expect(codes(`Option Explicit\nPublic Sub Main()\n' note\n${attribute}\n    Debug.Print 1\nEnd Sub\n`)).toContain('module-declaration-in-procedure');
	});

	it('are not body statements', () => {
		const mod = parseModule("Sub T()\n    x = 1 ' tail\n    ' note\n    Rem other\n    y = 2: ' after a colon\nEnd Sub\n");
		const body = (mod.members[0] as unknown as { body: Array<{ kind: string }> }).body;
		expect(body.map((node) => node.kind)).toEqual(['Assignment', 'Assignment']);
	});
});
