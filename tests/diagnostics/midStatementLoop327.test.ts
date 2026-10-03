// Diagnostics tests: the Mid statement inside a For loop, its start the
// counter (issue #327). Each case was run through pyVBAharness on 2026-10-02
// in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim s As String, i As Long\n    s = "abc"\n    ${body}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('the Mid statement in a loop (issue #327)', () => {
	it('starts past the end on the last pass', () => {
		expect(errors('For i = 1 To 4\n        Mid(s, i, 1) = "x"\n    Next')).toEqual(['runtime-argument-value']);
		expect(errors('For i = 1 To 4\n        Mid$(s, i) = "x"\n    Next')).toEqual(['runtime-argument-value']);
		expect(errors('For i = 0 To 2\n        Mid(s, i, 1) = "x"\n    Next')).toEqual(['runtime-argument-value']);
	});

	it('stays quiet when every pass starts inside the string', () => {
		expect(errors('For i = 1 To 3\n        Mid(s, i, 1) = "x"\n    Next')).toEqual([]);
	});
});
