// A loop counter one step past what it indexes (issue #200). The measured
// cases ran in Excel 16.0 through pyVBAharness: each raising one raises the
// error named every time it runs, and each quiet one runs.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

const DECLS = '    Dim i As Long, r As Long, s As String, parts As Variant\n    Dim c As New Collection\n    c.Add "a": c.Add "b": c.Add "c"\n    s = "abc"';

function errors(body: string, decls = DECLS, signature = 'Function Main() As String'): Array<{ code: string; message: string }> {
	const source = `Option Explicit\n${signature}\n${decls}\n${body}\nEnd ${signature.split(' ')[0]}\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1' }).diagnostics
		.filter((d) => d.severity === 'error')
		.map((d) => ({ code: d.code ?? '', message: d.message }));
}

describe('loop counters past what they index (issue #200)', () => {
	const RAISES: ReadonlyArray<readonly [string, string, string, string]> = [
		['Mid$ from 0', '    For i = 0 To Len(s) - 1\n        Main = Main & Mid$(s, i, 1)\n    Next i', 'runtime-argument-value', '5'],
		['Cells from row 0', '    For r = 0 To 3\n        Main = Main & Cells(r, 1).Value\n    Next r', 'host-argument-out-of-range', '1004'],
		['Cells from column 0', '    For i = 0 To 2\n        Main = Main & Cells(1, i).Value\n    Next i', 'host-argument-out-of-range', '1004'],
		['a Collection from 0', '    For i = 0 To c.Count - 1\n        Main = Main & c(i)\n    Next i', 'collection-index-out-of-range', '9'],
		['c.Item from 0', '    For i = 0 To c.Count - 1\n        Main = Main & c.Item(i)\n    Next i', 'collection-index-out-of-range', '9'],
		['UBound + 1', '    Dim a(1 To 5) As Long\n    For i = LBound(a) To UBound(a) + 1\n        a(i) = i\n    Next i', 'array-subscript-out-of-bounds', '9'],
		['Step -1 to 0', '    Dim a(1 To 5) As Long\n    For i = 5 To 0 Step -1\n        a(i) = i\n    Next i', 'array-subscript-out-of-bounds', '9'],
		['Split UBound + 1', '    parts = Split("a,b,c", ",")\n    For i = 1 To UBound(parts) + 1\n        Main = Main & parts(i)\n    Next i', 'array-subscript-out-of-bounds', '9'],
		['c.Count + 1', '    For i = 1 To c.Count + 1\n        Main = Main & c(i)\n    Next i', 'collection-index-out-of-range', '9'],
		['Left$ to -1', '    For i = Len(s) To -1 Step -1\n        Main = Main & Left$(s, i)\n    Next i', 'runtime-argument-value', '5'],
		['InStr from 0', '    For i = 0 To 2\n        Main = Main & InStr(i, s, "b")\n    Next i', 'runtime-argument-value', '5'],
		['Do While to 4', '    Dim a(1 To 3) As Long\n    i = 1\n    Do While i <= 4\n        a(i) = i\n        i = i + 1\n    Loop', 'array-subscript-out-of-bounds', '9'],
		// Not measured: the same forms, reached another way.
		['a statement inside With', '    Dim a(1 To 5) As Long\n    With c\n        For i = 1 To 6\n            With c\n                a(i) = i\n            End With\n        Next i\n    End With', 'array-subscript-out-of-bounds', '9'],
		['Step 2 lands on 7', '    Dim a(1 To 6) As Long\n    For i = 1 To 8 Step 2\n        a(i) = i\n    Next i', 'array-subscript-out-of-bounds', '9'],
		['Do Until past 3', '    Dim a(1 To 3) As Long\n    i = 1\n    Do Until i > 4\n        a(i) = i\n        i = i + 1\n    Loop', 'array-subscript-out-of-bounds', '9'],
	];
	it.each(RAISES)('reports %s', (_name, body, code, number) => {
		const hits = errors(body);
		expect(hits).toHaveLength(1);
		expect(hits[0].code).toBe(code);
		expect(hits[0].message).toContain(`Run-time error '${number}'`);
	});

	const RUNS: ReadonlyArray<readonly [string, string]> = [
		['Mid$ from 1 to Len', '    For i = 1 To Len(s)\n        Main = Main & Mid$(s, i, 1)\n    Next i'],
		['LBound to UBound', '    Dim a(1 To 5) As Long\n    For i = LBound(a) To UBound(a)\n        a(i) = i\n    Next i'],
		['0 To UBound of Split', '    parts = Split("a,b,c", ",")\n    For i = 0 To UBound(parts)\n        Main = Main & parts(i)\n    Next i'],
		['a use under an If', '    For i = 0 To Len(s) - 1\n        If i > 0 Then Main = Main & Mid$(s, i, 1)\n    Next i'],
		['the last pass under an If', '    Dim a(1 To 5) As Long\n    For i = 1 To 6\n        If i < 6 Then a(i) = i\n    Next i'],
		['a loop an empty string never runs', '    s = ""\n    For i = 0 To Len(s) - 1\n        Main = Main & Mid$(s, i, 1)\n    Next i'],
		['an Exit For before the use', '    For i = 0 To 3\n        If i = 0 Then Exit For\n        Main = Main & Cells(i, 1).Value\n    Next i'],
		['a Collection 1 To Count', '    For i = 1 To c.Count\n        Main = Main & c(i)\n    Next i'],
		['Do While below 4', '    Dim a(1 To 3) As Long\n    i = 1\n    Do While i < 4\n        a(i) = i\n        i = i + 1\n    Loop'],
		['Right$ from Len down to 0', '    For i = Len(s) To 0 Step -1\n        Main = Main & Right$(s, i)\n    Next i'],
		['a counter the body changes', '    Dim a(1 To 5) As Long\n    For i = 1 To 6\n        a(i) = i\n        i = i + 1\n    Next i'],
		// Not measured: what the rules cannot know stays quiet.
		['a counter a call may change', '    Dim a(1 To 5) As Long\n    For i = 1 To 6\n        Bump i\n        a(i) = i\n    Next i'],
		['a label before the use', '    Dim a(1 To 5) As Long\n    For i = 1 To 6\nHere:\n        a(i) = i\n    Next i'],
		['an array the body resizes', '    Dim b() As Long\n    ReDim b(1 To 3)\n    For i = 1 To UBound(b) + 1\n        ReDim Preserve b(1 To i)\n        b(i) = i\n    Next i'],
		['a Collection the body adds to', '    For i = 1 To c.Count + 1\n        c.Add "x"\n        Main = Main & c(i)\n    Next i'],
		['For Each', '    Dim v As Variant\n    For Each v In c\n        Main = Main & Mid$(s, 0 + Len(v), 1)\n    Next v'],
		['Step 2 stopping at 5', '    Dim a(1 To 5) As Long\n    For i = 1 To 6 Step 2\n        a(i) = i\n    Next i'],
		["another array's UBound", '    Dim a() As Variant, b() As Variant\n    a = Array(1, 2)\n    b = Array(1, 2, 3, 4)\n    For i = 0 To UBound(a) + 1\n        Main = Main & b(i)\n    Next i'],
	];
	it.each(RUNS)('stays quiet for %s', (_name, body) => {
		expect(errors(`${body}\n    Exit Function\nEnd Function\nSub Bump(ByRef n As Long)\n    n = n + 1\nEnd Sub\nFunction Unused() As String`)).toEqual([]);
	});

	it('names the pass and the counter', () => {
		const [hit] = errors('    For i = 0 To Len(s) - 1\n        Main = Main & Mid$(s, i, 1)\n    Next i');
		expect(hit.message).toBe("On the first pass of the For loop, where 'i' is 0: Argument 'Start' of 'Mid$' is 0; this will raise Run-time error '5': Invalid procedure call or argument.");
	});

	it('says a Collection that may be empty raises 5 instead', () => {
		const hits = errors('    For i = 1 To items.Count + 1\n        Unused = Unused & items(i)\n    Next i', '    Dim i As Long', 'Function Unused(items As Collection) As String');
		expect(hits.map((hit) => hit.message)).toEqual([
			"Counter 'i' reaches items.Count + 1 on its last pass, and 'items' holds its elements at 1 to items.Count. This will raise Run-time error '9': Subscript out of range, or '5' if 'items' is empty.",
		]);
	});

	it('reads UBound of an array it knows nothing else about', () => {
		const hits = errors('    For i = 0 To UBound(values) + 1\n        Unused = Unused & values(i)\n    Next i', '    Dim i As Long', 'Function Unused(values As Variant) As String');
		expect(hits.map((hit) => hit.message)).toEqual([
			"Counter 'i' reaches UBound(values) + 1 on its last pass, which for array 'values' is above its upper bound. This will raise Run-time error '9': Subscript out of range.",
		]);
	});
});
