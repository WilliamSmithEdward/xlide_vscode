// Diagnostics tests: #338's follow-ups (issue #613), a class member that
// calls itself through Me, With Me, another name for Me, a new instance of
// its class, or CallByName. Measured on 2026-10-02 in Excel 16.0 (build
// 20430), each class called from a standard module.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function codes(cls: string): string[] {
	return analyzeModule(`Option Explicit\n${cls}`, { moduleName: 'Class1', moduleKind: 'class' })
		.filter((diag) => diag.code === 'recursive-property-accessor' || diag.code === 'unbounded-recursion')
		.map((diag) => diag.code);
}

describe('a class member that calls itself (issue #613)', () => {
	it('takes a Get that assigns through Me as calling the Let', () => {
		expect(codes('Private m As Long\nPublic Property Let Value(ByVal v As Long)\n    m = v\nEnd Property\nPublic Property Get Value() As Long\n    If m = 0 Then Me.Value = 7\n    Value = m\nEnd Property\n')).toEqual([]);
	});

	it('reports the ways #338 missed', () => {
		const cases: Array<[string, string]> = [
			['Public Property Get Item(ByVal i As Long) As Long\n    Item = Me.Item(i)\nEnd Property\n', 'recursive-property-accessor'],
			['Public Property Get Value() As Long\n    With Me\n        Value = .Value\n    End With\nEnd Property\n', 'recursive-property-accessor'],
			['Public Property Let Value(ByVal v As Long)\n    With Me\n        .Value = v\n    End With\nEnd Property\n', 'recursive-property-accessor'],
			['Public Property Let Value(ByVal v As Long)\n    Dim o As Class1\n    Set o = Me\n    o.Value = v\nEnd Property\n', 'recursive-property-accessor'],
			['Public Function Twice() As Long\n    Dim o As Class1\n    Set o = New Class1\n    Twice = o.Twice\nEnd Function\n', 'unbounded-recursion'],
			['Public Sub Go()\n    CallByName Me, "Go", VbMethod\nEnd Sub\n', 'unbounded-recursion'],
		];
		for (const [cls, code] of cases) {
			expect(codes(cls), cls).toEqual([code]);
		}
	});

	it('stays quiet where the call ends', () => {
		for (const cls of [
			'Public Property Get Item(ByVal i As Long) As Long\n    If i > 0 Then Item = Me.Item(i - 1)\nEnd Property\n',
			'Private busy As Boolean\nPublic Sub Go()\n    If busy Then Exit Sub\n    busy = True\n    Me.Go\nEnd Sub\n',
			'Public Sub Go()\n    If False Then Me.Go\nEnd Sub\n',
			'Private m As Long\nPublic Property Get Value() As Long\n    Value = m\nEnd Property\nPublic Property Let Value(ByVal v As Long)\n    m = v + Me.Value\nEnd Property\n',
			'Public Function Twice() As Long\n    Dim o As Class1\n    Set o = New Class1\n    Twice = 2\nEnd Function\n',
			'Public Sub Go()\n    CallByName Me, "Stop2", VbMethod\nEnd Sub\nPublic Sub Stop2()\nEnd Sub\n',
			'Public Property Get Item(ByVal i As Long) As Long\n    Item = i\nEnd Property\nPublic Property Get Twice(ByVal i As Long) As Long\n    Twice = Me.Item(i)\nEnd Property\n',
		]) {
			expect(codes(cls), cls).toEqual([]);
		}
	});
});
