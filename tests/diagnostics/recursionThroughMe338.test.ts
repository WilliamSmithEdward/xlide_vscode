// Diagnostics tests: class members that call themselves through Me or with
// empty parentheses (issue #338). Each case was run through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430): the raising ones raise error 28.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function inClass(code: string): string[] {
	return analyzeModule(`Option Explicit\n${code}`, { moduleKind: 'class', moduleName: 'Class1' })
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code);
}

describe('a class member that calls itself (issue #338)', () => {
	it('is reported through Me or with empty parentheses', () => {
		expect(inClass('Public Property Let Value(ByVal v As Long)\n    Me.Value = v\nEnd Property\nPublic Property Get Value() As Long\nEnd Property\n')).toEqual(['recursive-property-accessor']);
		expect(inClass('Public Property Set Items(ByVal v As Collection)\n    Set Me.Items = v\nEnd Property\nPublic Property Get Items() As Collection\nEnd Property\n')).toEqual(['recursive-property-accessor']);
		expect(inClass('Public Property Get Value() As Long\n    Value = Value() + 1\nEnd Property\n')).toEqual(['recursive-property-accessor']);
		expect(inClass('Public Function Twice() As Long\n    Twice = Me.Twice\nEnd Function\n')).toEqual(['unbounded-recursion']);
		expect(inClass('Public Function Twice() As Long\n    Twice = Me.Twice()\nEnd Function\n')).toEqual(['unbounded-recursion']);
		expect(inClass('Public Sub Go()\n    Me.Go\nEnd Sub\n')).toEqual(['unbounded-recursion']);
	});

	it('stays quiet on the return variable, another member, or a guarded call', () => {
		expect(inClass('Public Property Get Value() As Long\n    Value = Value + 1\nEnd Property\n')).toEqual([]);
		expect(inClass('Private m As Long\nPublic Property Let Value(ByVal v As Long)\n    m = v\nEnd Property\nPublic Property Get Value() As Long\n    Value = m\nEnd Property\nPublic Sub Bump()\n    Me.Value = Me.Value + 1\nEnd Sub\n')).toEqual([]);
		expect(inClass('Public Function Twice(Optional ByVal n As Long = 0) As Long\n    If n < 1 Then Twice = Me.Twice(n + 1)\nEnd Function\n')).toEqual([]);
	});
});
