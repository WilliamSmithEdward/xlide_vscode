// Diagnostics tests: Application.Run of a procedure whose name reads as a
// cell address (issue #468). Each sample was measured through pyVBAharness on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const NAMES = ['Pub2', 'AB12', 'XFD1', 'R1C1', 'R2', 'C3', 'XFE1', 'Calc1', 'Test1', 'Pubx', 'Run_1', 'Main2', 'Go'];
const MODULE2 = `Option Explicit\n${NAMES.map((name) => `Public Sub ${name}()\nEnd Sub\n`).join('')}`;

function found(call: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${call}\n    Main = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }, { moduleName: 'Module2', source: MODULE2 }], 'Module1')
		.filter((d) => d.severity === 'error').map((d) => d.message);
}

describe('Application.Run of a name that reads as a cell address (issue #468)', () => {
	it('reports the bare name, and names the module-qualified fix', () => {
		for (const name of ['Pub2', 'AB12', 'XFD1', 'R1C1', 'R2', 'C3']) {
			const hits = found(`Application.Run "${name}"`);
			expect(hits, name).toHaveLength(1);
			expect(hits[0], name).toContain(`"Module2.${name}"`);
		}
		expect(found('Run "Pub2"')).toHaveLength(1);
	});

	it('stays quiet on names past the sheet, other names, and qualified ones', () => {
		for (const call of [...['XFE1', 'Calc1', 'Test1', 'Pubx', 'Run_1', 'Main2', 'Go'].map((name) => `Application.Run "${name}"`), 'Application.Run "Module2.Pub2"', 'Application.Run "Module2.R1C1"']) {
			expect(found(call), call).toEqual([]);
		}
	});
});
