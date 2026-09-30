// Diagnostics tests: a line number followed by a label, `10 L1: x = 1`, and
// where a label or a line number may stand (issue #230). Measured through
// pyVBAharness in Excel 16.0 on 2026-09-30, with a full compile: the line
// holds both targets, GoTo 10 and Erl see the number, GoTo L1 and Resume L1
// the name. A label is one only at the start of its line, after the line
// number if any: in `10: L1:` and `10 L1: L2:` the word is a call, "Sub or
// Function not defined". A line number after a colon is "Syntax error".

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { collectProcedureLabels } from '../../src/analyzer/flow/procedureLabels';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import type { ProcedureNode } from '../../src/analyzer/parser/nodes';
import { byCode } from '../helpers/diagnostics';

function main(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.join('\n')}\nEnd Function\n`;
}

/** Codes found, with a bare call checked against Main alone. */
function codes(src: string): string[] {
	return analyzeModule(src, { knownProcedures: new Set(['main']), knownIdentifiers: new Set() })
		.filter((d) => d.severity === 'error')
		.map((d) => d.code ?? '');
}

describe('a line number followed by a label (issue #230)', () => {
	it.each([
		['GoTo L1', ['    GoTo L1', '    Main = 0', '10 L1: Main = 1']],
		['GoTo 10', ['    GoTo 10', '    Main = 0', '10 L1: Main = 1']],
		['the label alone on its line', ['    GoTo L1', '    Main = 0', '10 L1:', '    Main = 1']],
		['Resume L1', ['    On Error GoTo H', '    Error 5', '    Exit Function', 'H:', '    Resume L1', '    Main = 0', '10 L1: Main = 1']],
		['Erl', ['    On Error GoTo H', '10 L1: Error 5', '    Exit Function', 'H:', '    Main = Erl']],
		['an assignment after it', ['    Dim x As Long', '10 L1: x = 2', '    Main = x']],
		['a space before the colon', ['    GoTo L1', '10 L1 : Main = 1']],
		['a label named like the procedure', ['    GoTo 10', '10 Main: Main = 1']],
	])('declares both targets: %s', (_label, lines) => {
		expect(codes(main(...lines))).toEqual([]);
	});

	it('collects the number and the name', () => {
		const src = main('10 L1: Main = 1', '20', 'L2: Main = 2');
		const proc = parseModule(src).members.find((m): m is ProcedureNode => m.kind === 'Procedure');
		expect([...collectProcedureLabels(src, proc as ProcedureNode).keys()]).toEqual(['line:10', 'name:l1', 'line:20', 'name:l2']);
	});

	it('takes a word after a line number with no colon for a call, not a label', () => {
		const src = main('    GoTo Foo', '10 Foo');
		const proc = parseModule(src).members.find((m): m is ProcedureNode => m.kind === 'Procedure');
		expect([...collectProcedureLabels(src, proc as ProcedureNode).keys()]).toEqual(['line:10']);
		expect(codes(src)).toEqual(expect.arrayContaining(['undefined-label', 'unknown-call']));
	});

	it('knows a handler written after a line number, `10 H:`', () => {
		const src = main('    On Error GoTo H', '    Main = 1', '10 H: Err.Raise Err.Number');
		expect(byCode(analyzeModule(src), 'handler-fall-through')).toHaveLength(1);
	});

	it('reports the name used again as a duplicate', () => {
		const hits = byCode(analyzeModule(main('    GoTo 10', '10 L1: Main = 1', 'L1: Main = 2')), 'duplicate-label');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain("'L1'");
	});
});

describe('a label is one only at the start of its line (issue #230)', () => {
	it('reads a word after `10:` or after a label as a call', () => {
		expect(codes(main('    GoTo L1', '10: L1: Main = 1'))).toEqual(expect.arrayContaining(['undefined-label', 'unknown-call']));
		const second = byCode(analyzeModule(main('    GoTo L1', '10 L1: L2: Main = 1'), { knownProcedures: new Set(['main']), knownIdentifiers: new Set() }), 'unknown-call');
		expect(second.map((d) => d.message)).toEqual(["Sub or Function not defined: 'L2'."]);
		expect(codes(main('    Main = 0: L1: Main = 1'))).toEqual(['unknown-call']);
	});

	it('leaves a call to a Sub of that name after a colon alone', () => {
		const src = 'Option Explicit\nSub L1()\nEnd Sub\nFunction Main() As Variant\n    Main = 0: L1: Main = 1\nEnd Function\n';
		expect(analyzeModule(src, { knownProcedures: new Set(['main', 'l1']), knownIdentifiers: new Set() }).filter((d) => d.severity === 'error')).toEqual([]);
	});
});

describe('invalid-line-number: a line number after a colon (issue #230)', () => {
	it.each([
		['after a statement', '    Main = 0: 20 Main = 1'],
		['after a line number', '10: 20 Main = 1'],
		['after a label', 'L1: 20 Main = 1'],
		['after a line number and a label', '10 L1: 20 Main = 1'],
		['alone after them', '10 L1: 20'],
		["in a one-line If's tail", '    If True Then Main = 1: 20 Main = 2'],
		['after Then and a colon', '    If True Then: 20 Main = 2'],
		['after a line continuation', '    Main = 1 _\n    : 20 Main = 2'],
	])('%s', (_label, line) => {
		const hits = byCode(analyzeModule(main(line)), 'invalid-line-number');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Line number 20 is not at the start of its line');
	});

	it('leaves a line number at the start of a line, and Else 20, which is GoTo 20', () => {
		expect(byCode(analyzeModule(main('10 Main = 1', '  20 Main = 2', '\t30')), 'invalid-line-number')).toEqual([]);
		expect(byCode(analyzeModule(main('    If False Then Main = 0 Else 20', '    Main = 2', '20  Main = 1')), 'invalid-line-number')).toEqual([]);
	});
});
