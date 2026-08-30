import { describe, expect, it } from 'vitest';
import { classifyReferenceKinds } from '../src/analyzer';
import { collectSymbolReferences } from '../src/vbaReferenceResolution';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';

// Issue #55: every reference says whether it READS or WRITES. The rules are
// syntactic - the assignment family writes its target's terminal name,
// declarations write the names they introduce, Mid and ReDim Preserve modify
// in place - and the one gray zone, a variable passed where a ByRef parameter
// might write it, deliberately stays a read.

/** Kind of the `nth` whole-identifier occurrence (0-based) of `name`. */
function kindOf(source: string, name: string, nth = 0): string {
	const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const re = new RegExp(`(?<![A-Za-z0-9_])${escaped}(?![A-Za-z0-9_])`, 'g');
	let at = -1;
	for (let i = 0; i <= nth; i++) {
		const m = re.exec(source);
		if (!m) { throw new Error(`${name} occurrence ${nth} not found`); }
		at = m.index;
	}
	const kinds = classifyReferenceKinds(source, [at]);
	return kinds.get(at)!;
}

describe('classifyReferenceKinds', () => {
	it('classifies the issue #55 procedure exactly', () => {
		const src = [
			'Public Sub PostInvoices()',
			'    Dim lastRow As Long',
			'    Dim total As Currency',
			'    Dim i As Long',
			'    lastRow = Sheet1.UsedRange.Rows.Count',
			'',
			'    total = 0',
			'    For i = 2 To lastRow',
			'        If Sheet1.Cells(i, 3).Value > 0 Then',
			'            total = total + Sheet1.Cells(i, 3).Value',
			'        End If',
			'    Next i',
			'',
			'    Debug.Print "Posted " & total',
			'End Sub',
		].join('\r\n');
		expect(kindOf(src, 'total', 0)).toBe('write');  // Dim total
		expect(kindOf(src, 'total', 1)).toBe('write');  // total = 0
		expect(kindOf(src, 'total', 2)).toBe('write');  // total = total + ... (left)
		expect(kindOf(src, 'total', 3)).toBe('read');   // ... = total + ... (right)
		expect(kindOf(src, 'total', 4)).toBe('read');   // Debug.Print
		expect(kindOf(src, 'lastRow', 1)).toBe('write'); // lastRow = ...
		expect(kindOf(src, 'lastRow', 2)).toBe('read');  // For ... To lastRow
		expect(kindOf(src, 'i', 1)).toBe('write');       // For i = ...
	});

	it('covers the assignment family and the declarations', () => {
		expect(kindOf('Set obj = New Collection', 'obj')).toBe('write');
		expect(kindOf('Let x = 1', 'x')).toBe('write');
		expect(kindOf('arr(i) = 5', 'arr')).toBe('write');
		expect(kindOf('arr(i) = 5', 'i')).toBe('read');
		expect(kindOf('thing.Prop = 1', 'thing')).toBe('read');
		expect(kindOf('thing.Prop = 1', 'Prop')).toBe('write');
		expect(kindOf('a.b(i).c = 1', 'c')).toBe('write');
		expect(kindOf('Sub P()\r\n    With o\r\n        .Field = 2\r\n    End With\r\nEnd Sub', 'Field')).toBe('write');
		expect(kindOf('For Each item In coll\r\nNext item', 'item', 0)).toBe('write');
		expect(kindOf('For Each item In coll\r\nNext item', 'coll')).toBe('read');
		expect(kindOf('Const LIMIT = 10', 'LIMIT')).toBe('write');
		expect(kindOf('Private Const A = 1, B = A + 1', 'B')).toBe('write');
		expect(kindOf('Private Const A = 1, B = A + 1', 'A', 1)).toBe('read');
		expect(kindOf('Dim x As Long, y(10) As String', 'y')).toBe('write');
		expect(kindOf('Dim x As Worksheet', 'Worksheet')).toBe('read'); // the type reads
	});

	it('covers the in-place modifiers and the file reads', () => {
		expect(kindOf('ReDim a(1 To n)', 'a')).toBe('write');
		expect(kindOf('ReDim a(1 To n)', 'n')).toBe('read');
		expect(kindOf('ReDim Preserve a(1 To n)', 'a')).toBe('readwrite');
		expect(kindOf('Erase a, b', 'b')).toBe('write');
		expect(kindOf('Mid(s, 2, 3) = "ab"', 's')).toBe('readwrite');
		expect(kindOf('LSet rec = other', 'rec')).toBe('write');
		expect(kindOf('Line Input #1, textLine', 'textLine')).toBe('write');
		expect(kindOf('Input #1, a, b', 'b')).toBe('write');
		expect(kindOf('Get #1, , record', 'record')).toBe('write');
	});

	it('keeps the honest reads honest', () => {
		expect(kindOf('If x = 1 Then y = 2', 'x')).toBe('read');       // comparison
		expect(kindOf('If x = 1 Then y = 2', 'y')).toBe('write');      // the colon-free inline branch... assignment after Then
		expect(kindOf('Call Update(total)', 'total')).toBe('read');     // ByRef stays a read
		expect(kindOf('foo bar:=value', 'value')).toBe('read');         // named argument
		expect(kindOf('foo bar:=value', 'bar')).toBe('read');
		expect(kindOf('Debug.Print total', 'total')).toBe('read');
	});

	it('writes the parameter names and the procedure name in a signature', () => {
		const src = 'Public Function Sum(ByVal a As Long, Optional b As Long = 0) As Long\r\n    Sum = a + b\r\nEnd Function';
		expect(kindOf(src, 'Sum', 0)).toBe('write');  // declaration
		expect(kindOf(src, 'a', 0)).toBe('write');    // parameter
		expect(kindOf(src, 'b', 0)).toBe('write');
		expect(kindOf(src, 'Sum', 1)).toBe('write');  // return assignment
		expect(kindOf(src, 'a', 1)).toBe('read');
	});
});

describe('collectSymbolReferences carries kinds', () => {
	it('returns the five total references with their kinds', () => {
		const source = [
			'Public Sub PostInvoices()',
			'    Dim total As Currency',
			'    total = 0',
			'    total = total + 1',
			'    Debug.Print total',
			'End Sub',
			'',
		].join('\r\n');
		const mods = [{ moduleName: 'Invoices', source, type: 'standard', documentType: undefined }];
		const project = buildVbaProjectIndex(mods.map((m) => ({ moduleName: m.moduleName, source: m.source, type: m.type })));
		const byModule = new Map(mods.map((m) => [m.moduleName.toLowerCase(), m]));
		const at = source.indexOf('total');
		const result = collectSymbolReferences(
			byModule, project, mods, source, 'Invoices', mods[0],
			'total', at + 'total'.length, at, true,
		);
		expect(result.hasSymbol).toBe(true);
		const kinds = result.references
			.sort((a, b) => a.line - b.line || a.column - b.column)
			.map((r) => `${r.line}:${r.kind}`);
		expect(kinds).toEqual(['1:write', '2:write', '3:write', '3:read', '4:read']);
	});
});
