// Diagnostics tests: a With block over an element, `With c(1)` on a
// Collection that holds Collections and `With a(0)` on an element never set
// (issue #295). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function raised(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

const INNER = 'Dim c As Collection, d As Collection\n    Set c = New Collection\n    Set d = New Collection\n    d.Add 1\n    c.Add d\n    ';

describe('With over an element (issue #295)', () => {
	it('reads `.Item` inside With c(1) as c(1)\'s', () => {
		expect(raised(`${INNER}With c(1)\n        Main = .Item(2)\n    End With`)).toEqual(['9']);
		expect(raised(`${INNER}With c(1)\n        Main = .Item("k")\n    End With`)).toEqual(['5']);
		expect(raised(`${INNER}With c(1)\n        .Remove 1\n    End With\n    Main = c(1).Item(1)`)).toEqual(['5']);
		for (const body of [`${INNER}With c(1)\n        Main = .Item(1)\n    End With`, `${INNER}With c(1)\n        .Add 2\n        Main = .Item(2)\n    End With`]) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('reports With a(0) on an element never set, once the body reaches it', () => {
		expect(raised('Dim a(1) As Collection\n    With a(0)\n        Main = .Count\n    End With')).toEqual(['91']);
		expect(raised('Dim a(1) As Object\n    With a(0)\n        Main = .Count\n    End With')).toEqual(['91']);
		for (const body of [
			'Dim a(1) As Collection\n    Set a(0) = New Collection\n    With a(0)\n        Main = .Count\n    End With',
			'Dim a(1) As Collection\n    With a(0)\n    End With\n    Main = 2',
			'Dim a(1) As Collection\n    With a(0)\n        Main = 3\n    End With',
			'Dim o As Collection\n    With o\n    End With\n    Main = 2',
		]) {
			expect(raised(body), body).toEqual([]);
		}
	});
});
