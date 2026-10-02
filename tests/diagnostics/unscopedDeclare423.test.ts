// Diagnostics tests: a Declare with no Public or Private in a standard module
// is Public (issue #423). Measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

function errors(declare: string, ...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	const module2 = `Option Explicit\n${declare}\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, { moduleName: 'Module2', source: module2 }], 'Module1')
		.filter((diag) => diag.severity === 'error')
		.map((diag) => `${diag.code}: ${diag.message}`);
}

const UNSCOPED = 'Declare PtrSafe Function M Lib "kernel32" Alias "GetTickCount" () As Long';

describe('a Declare with no scope keyword', () => {
	it('is called from another module, bare or qualified', () => {
		for (const line of ['Call M', 'Module2.M', 'Main = M', 'Main = Module2.M']) {
			expect(errors(UNSCOPED, line, 'Main = 1'), line).toEqual([]);
		}
	});

	it('is still hidden when Private', () => {
		expect(errors(`Private ${UNSCOPED}`, 'Main = M').length).toBeGreaterThan(0);
	});
});
