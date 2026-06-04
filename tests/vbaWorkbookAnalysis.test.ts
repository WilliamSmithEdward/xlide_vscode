import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
	workspace: {
		getConfiguration: () => ({
			get: (_key: string, fallback: unknown) => fallback,
		}),
	},
}));

import { analyzeWorkbook } from '../src/vbaWorkbookAnalysis';
import type { PythonBridge } from '../src/pythonBridge';

interface TestModule {
	name: string;
	type: string;
	source: string;
}

function bridgeForModules(modules: TestModule[]): PythonBridge {
	const byName = new Map(modules.map((mod) => [mod.name, mod]));
	return {
		call: vi.fn(async (method: string, payload: { module?: string }) => {
			if (method === 'listModules') {
				return modules.map(({ name, type }) => ({ name, type }));
			}
			if (method === 'readModule' && payload.module) {
				const mod = byName.get(payload.module);
				if (!mod) {
					throw new Error(`Unknown module ${payload.module}`);
				}
				return { source: mod.source };
			}
			throw new Error(`Unexpected bridge call ${method}`);
		}),
	} as unknown as PythonBridge;
}

describe('analyzeWorkbook metadata summary', () => {
	it('attaches shared rule metadata to structural and semantic workbook problems', async () => {
		const bridge = bridgeForModules([
			{
				name: 'BlockBroken',
				type: 'standard',
				source: 'Option Explicit\nSub Broken()\n\tMsgBox 1\n',
			},
			{
				name: 'Declarations',
				type: 'standard',
				source: 'Option Explicit\nSub Assigns()\n\tNotDeclared = 1\nEnd Sub\n',
			},
			{
				name: 'RuntimeRisk',
				type: 'standard',
				source: 'Option Explicit\nFunction NeedsReturn()\nEnd Function\n',
			},
		]);

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');
		const byCode = new Map(result.problems.map((problem) => [problem.code, problem]));

		expect([...byCode.keys()].sort()).toEqual([
			'missing-block-closer',
			'missing-return-assignment',
			'undeclared-variable',
		]);
		expect(byCode.get('missing-block-closer')).toMatchObject({
			ruleTitle: 'Missing block closer',
			category: 'syntax',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(byCode.get('undeclared-variable')).toMatchObject({
			category: 'project-symbol',
			vbeCompileEquivalent: true,
			diagnosticKind: 'compile-error',
		});
		expect(byCode.get('missing-return-assignment')).toMatchObject({
			category: 'semantic',
			vbeCompileEquivalent: false,
			diagnosticKind: 'runtime-risk',
		});
		expect(result.summary.byCategory).toMatchObject({
			'project-symbol': 1,
			semantic: 1,
			syntax: 1,
		});
		expect(result.summary.byDiagnosticKind).toMatchObject({
			'compile-error': 2,
			'runtime-risk': 1,
		});
		expect(result.summary.vbeCompileEquivalentCount).toBe(2);
		expect(result.summary.nonVbeCompileEquivalentCount).toBe(1);
	});

	it('surfaces replacement quick fixes for mismatched procedure closers', async () => {
		const bridge = bridgeForModules([
			{
				name: 'Performance',
				type: 'class',
				source:
					'Public Property Get Measurement() As Double\n' +
					'    Measurement = 1\n' +
					'End Function\n',
			},
		]);

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');
		const problem = result.problems.find((item) => item.code === 'missing-block-closer');

		expect(problem).toMatchObject({
			expectedClose: 'End Property',
			quickFixAvailable: true,
			quickFixTitles: ["Replace 'End Function' with 'End Property'"],
		});
	});
});
