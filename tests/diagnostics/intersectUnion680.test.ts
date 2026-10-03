// Diagnostics tests: Intersect and Union whose failure the literals prove
// (issue #680). Measured on 2026-10-03 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getWordObjectModel } from '../../src/analyzer/host/wordObjectModel';

function errors(body: string, after = ''): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n${after}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.message);
}

const WS = 'Dim ws As Worksheet, r As Range\n    Set ws = ActiveSheet\n    ';
const TWO = 'Dim ws As Worksheet, w2 As Worksheet, r As Range\n    Set ws = ActiveSheet\n    Set w2 = Worksheets.Add\n    ';
const N = 'Dim n As Range, r As Range\n    ';

describe('an Intersect of ranges that do not meet, held in a variable (issue #680)', () => {
	it('reports a member of it', () => {
		for (const body of [
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    Main = r.Address`,
			`${WS}Set r = Intersect(ws.Range("A1:B2"), ws.Range("C3:D4"))\n    Main = r.Address`,
			`${WS}Set r = Intersect(Range("A1"), Range("C3"))\n    Main = r.Address`,
			`${WS}Set r = Application.Intersect(ws.Range("A1"), ws.Range("C3"))\n    Main = r.Address`,
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    Main = r.Count`,
		]) {
			const found = errors(body);
			expect(found, body).toHaveLength(1);
			expect(found[0], body).toContain("Object variable 'r' is Nothing before member access");
		}
	});

	it('stays quiet when the ranges meet, the code checks, or the variable is Set again', () => {
		for (const body of [
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    Main = (r Is Nothing)`,
			`${WS}Set r = Intersect(ws.Range("A1:C3"), ws.Range("B2"))\n    Main = r.Address`,
			`${WS}Set r = Intersect(ws.Rows(2), ws.Columns(3))\n    Main = r.Address`,
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    If r Is Nothing Then Main = "none" Else Main = r.Address`,
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    If Not r Is Nothing Then Main = r.Address Else Main = "none"`,
			`${WS}Set r = Intersect(ws.Range("A1"), ws.Range("C3"))\n    Set r = ws.Range("A1")\n    Main = r.Address`,
			`${WS}Set r = Intersect(Range("A1"), ws.Range("C3"))\n    If r Is Nothing Then Main = "none" Else Main = r.Address`,
			`${WS}Set r = Intersect(ws.Range("A1:C3"), ws.Range("C3:D4"))\n    Main = r.Address`,
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});

	it('leaves a project Intersect of its own alone', () => {
		const own = 'Private Function Intersect(a As Range, b As Range) As Range\n    Set Intersect = b\nEnd Function\n';
		expect(errors('Dim r As Range\n    Set r = Intersect(Range("A1"), Range("C3"))\n    Main = r.Address', own)).toEqual([]);
	});
});

describe('Intersect and Union across two sheets (issue #680)', () => {
	it('reports Intersect as Union is, through Application too', () => {
		expect(errors(`${TWO}Set r = Intersect(ws.Range("A1"), w2.Range("A1"))\n    Main = "ok"`)).toEqual([
			"Intersect takes ranges of one sheet, and 'ws' and 'w2' are different sheets. This will raise Run-time error '1004': Method 'Intersect' of object '_Global' failed.",
		]);
		expect(errors(`${TWO}Set r = Application.Intersect(ws.Range("A1"), w2.Range("A1"))\n    Main = "ok"`)[0]).toContain("of object '_Application' failed");
		expect(errors(`${TWO}Main = (Intersect(ws.Range("A1"), w2.Range("A1")) Is Nothing)`)).toHaveLength(1);
		expect(errors(`${TWO}Set r = Union(ws.Range("A1"), w2.Range("A1"))\n    Main = "ok"`)[0]).toContain("Method 'Union' of object '_Global' failed");
		expect(errors(`${TWO}Set r = Application.Union(ws.Range("A1"), w2.Range("A1"))\n    Main = "ok"`)[0]).toContain("Method 'Union' of object '_Application' failed");
	});
});

describe('Union given Nothing (issue #680)', () => {
	it('reports a local still Nothing, in any place', () => {
		for (const body of [
			`${N}Set r = Union(n, ActiveSheet.Range("A1"))\n    Main = r.Address`,
			`${N}Set r = Union(ActiveSheet.Range("A1"), n)\n    Main = r.Address`,
			`${N}Set n = Nothing\n    Set r = Union(n, ActiveSheet.Range("A1"))\n    Main = r.Address`,
			`${N}Set r = Application.Union(n, ActiveSheet.Range("A1"))\n    Main = r.Address`,
			`${N}Set r = Union(ActiveSheet.Range("A1"), ActiveSheet.Range("A2"), n)\n    Main = r.Address`,
		]) {
			expect(errors(body), body).toEqual([
				"Object variable 'n' is Nothing, and Union takes no Nothing. This will raise Run-time error '5': Invalid procedure call or argument.",
			]);
		}
	});

	it('stays quiet on the usual idiom and a local Set first', () => {
		for (const body of [
			'Dim n As Range, c As Range\n    For Each c In ActiveSheet.Range("A1:A3").Cells\n        If n Is Nothing Then\n            Set n = c\n        Else\n            Set n = Union(n, c)\n        End If\n    Next c\n    Main = n.Address',
			'Dim n As Range, c As Range\n    For Each c In ActiveSheet.Range("A1:A3").Cells\n        If n Is Nothing Then Set n = c Else Set n = Union(n, c)\n    Next c\n    Main = n.Address',
			'Dim n As Range, c As Range\n    Set c = ActiveSheet.Range("A1")\n    If Not n Is Nothing Then Set n = Union(n, c)\n    Main = (n Is Nothing)',
			'Dim n As Range, c As Range\n    Set c = ActiveSheet.Range("A1")\n    If n Is Nothing Then\n        Set n = c\n    Else\n        Set n = Union(n, c)\n    End If\n    Main = n.Address',
			`${N}Set n = ActiveSheet.Range("B2")\n    Set r = Union(n, ActiveSheet.Range("A1"))\n    Main = r.Address`,
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});

	it('leaves a project Union of its own alone', () => {
		const own = 'Private Function Union(a As Range, b As Range) As Range\n    Set Union = b\nEnd Function\n';
		expect(errors(`${N}Set r = Union(n, ActiveSheet.Range("A1"))\n    Main = r.Address`, own)).toEqual([]);
	});

	it('leaves Word alone, which has no Union', () => {
		const src = `Option Explicit\nFunction Main() As Variant\n    Dim n As Object, r As Object\n    Set r = Union(n, n)\nEnd Function\n`;
		expect(analyzeModule(src).filter((diag) => diag.message.includes('Union takes no Nothing'))).toHaveLength(2);
		expect(analyzeModule(src, { hostModel: getWordObjectModel() }).filter((diag) => diag.message.includes('Union takes no Nothing'))).toEqual([]);
	});
});
