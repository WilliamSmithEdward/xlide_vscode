// Diagnostics tests: a Loop While or Loop Until line reads its condition after
// the body, so an object the body sets is set by then (issue #560). The first
// case was run in Excel 16.0 64-bit (build 20430) and returns 1.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a bottom-tested loop condition reads what the body set (issue #560)', () => {
	it('stays quiet when the body sets the object the condition reads', () => {
		expect(errors('Dim c As Collection\n    Do\n        Set c = New Collection\n        c.Add 1\n    Loop While c.Count < 1\n    Main = c.Count')).toEqual([]);
		expect(errors('Dim c As Collection\n    Do\n        Set c = New Collection\n    Loop Until c.Count = 0\n    Main = 1')).toEqual([]);
	});

	it('still reports an object the body never sets', () => {
		expect(errors('Dim c As Collection, i As Long\n    Do\n        i = i + 1\n    Loop While c.Count < 1\n    Main = i')).toEqual(['object-variable-not-set']);
	});
});
