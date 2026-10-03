// Diagnostics tests: an error value in the uses and from the sources #310
// left (issue #607). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HEAD = 'Private Function TakeL(ByVal p As Long) As Long\n    TakeL = p\nEnd Function\nPrivate Function TakeV(p As Variant) As Variant\n    TakeV = 1\nEnd Function\n';
const E = 'Dim v\n    v = CVErr(2042)\n    ';

function raised(body: string): string[] {
	const source = `Option Explicit\n${HEAD}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('an error value where a number or text is needed (issue #607)', () => {
	it('reports each use Excel refuses', () => {
		const cases: Array<[string, string]> = [
			['Main = -v', '13'], ['Main = v And 1', '13'], ['Main = v Or 1', '13'], ['Main = v Like "a*"', '13'],
			['Main = IIf(v, 1, 2)', '13'], ['Do While v\n        Exit Do\n    Loop', '13'], ['Dim i As Long\n    For i = 1 To v\n    Next', '13'],
			['Main = Int(v)', '13'], ['Main = Round(v)', '13'], ['Main = CDate(v)', '13'], ['Main = CByte(v)', '6'], ['Main = Str(v)', '13'],
			['Main = Left(v, 1)', '13'], ['Main = InStr(v, "a")', '13'], ['Main = Space(v)', '13'], ['Dim a(3) As Long\n    Main = a(v)', '13'],
			['Main = Choose(v, 1, 2)', '13'], ['Main = TakeL(v)', '13'], ['Main = Join(Array(v, "a"), ",")', '13'],
			['Main = WorksheetFunction.Sum(v, 1)', '1004'],
		];
		for (const [use, error] of cases) {
			expect(raised(E + use), use).toEqual([error]);
		}
	});

	it('leaves alone what takes an error value', () => {
		for (const use of ['Main = CDbl(v)', 'Main = CLng(v)', 'Main = Application.Sum(v, 1)', 'Main = IsError(v)', 'Main = TakeV(v)', 'Main = CVar(v)']) {
			expect(raised(E + use), use).toEqual([]);
		}
	});

	it('follows an error value through a local from a cell or an array', () => {
		for (const body of [
			'Dim v, w\n    w = CVErr(2042)\n    v = w\n    Main = v + 1',
			'Dim v\n    Range("A1").Formula = "=1/0"\n    v = Range("A1").Value\n    Main = v + 1',
			'Dim v\n    Range("A1").Formula = "=NA()"\n    v = Range("A1")\n    Main = v + 1',
			'Dim v\n    ActiveSheet.Range("A1").Formula = "=1/0"\n    v = ActiveSheet.Range("A1").Value\n    Main = v + 1',
			'Dim v\n    Cells(1, 1).Formula = "=1/0"\n    v = Cells(1, 1).Value\n    Main = v + 1',
			'Dim v\n    Range("A1").Value = "=1/0"\n    v = Range("A1").Value\n    Main = v + 1',
			'Dim v, a\n    a = Array(1, CVErr(2007))\n    v = a(1)\n    Main = v + 1',
			'Dim v\n    v = Array(1, CVErr(2007))(1)\n    Main = v + 1',
			'Dim v\n    Range("A1").Formula = "=1/0"\n    v = Range("A1").Value\n    Range("A1").Value = 5\n    Main = v + 1',
		]) {
			expect(raised(body), body).toEqual(['13']);
		}
		expect(raised('Dim v\n    Range("A1").Formula = "=1/0"\n    v = Range("B1").Value\n    Main = v + 1')).toEqual([]);
	});
});
