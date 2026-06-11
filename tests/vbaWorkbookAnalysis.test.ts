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
import { BridgeError, JSONRPC_METHOD_NOT_FOUND } from '../src/pythonBridgeErrors';
import { fakePythonBridge } from './helpers/fakePythonBridge';
import { deferred, flushPromises } from './helpers/async';

describe('analyzeWorkbook metadata summary', () => {
	it('loads workbook modules through the batch read endpoint', async () => {
		const bridge = fakePythonBridge([
			{
				name: 'Module1',
				type: 'standard',
				source: 'Option Explicit\nSub T()\nEnd Sub\n',
			},
		]);

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');

		expect(result.moduleCount).toBe(1);
		expect(vi.mocked(bridge.call).mock.calls.map(([method]) => method)).toEqual(['readModules']);
	});

	it('falls back to legacy list/read calls when the backend lacks batch reads', async () => {
		const bridge = fakePythonBridge([
			{
				name: 'Module1',
				type: 'standard',
				source: 'Option Explicit\nSub T()\nEnd Sub\n',
			},
		], { supportsBatchRead: false });

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');

		expect(result.moduleCount).toBe(1);
		expect(vi.mocked(bridge.call).mock.calls.map(([method]) => method)).toEqual([
			'readModules',
			'listModules',
			'readModule',
		]);
	});

	it('reads legacy fallback modules concurrently when batch reads are unavailable', async () => {
		const module1 = deferred<{ source: string }>();
		const module2 = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((method: string, payload: { module?: string }) => {
				if (method === 'readModules') {
					return Promise.reject(new BridgeError('Method not found: readModules', JSONRPC_METHOD_NOT_FOUND));
				}
				if (method === 'listModules') {
					return Promise.resolve([
						{ name: 'Module1', type: 'standard' },
						{ name: 'Module2', type: 'standard' },
					]);
				}
				if (method === 'readModule' && payload.module === 'Module1') {
					return module1.promise;
				}
				if (method === 'readModule' && payload.module === 'Module2') {
					return module2.promise;
				}
				return Promise.reject(new Error(`Unexpected bridge call ${method}`));
			}),
		} as unknown as PythonBridge;

		const pending = analyzeWorkbook(bridge, 'Book.xlsm');
		await flushPromises();

		expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'readModule')).toHaveLength(2);

		module1.resolve({ source: 'Option Explicit\nSub First()\nEnd Sub\n' });
		module2.resolve({ source: 'Option Explicit\nSub Second()\nEnd Sub\n' });

		await expect(pending).resolves.toMatchObject({ moduleCount: 2 });
	});

	it('attaches shared rule metadata to structural and semantic workbook problems', async () => {
		const bridge = fakePythonBridge([
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
		const bridge = fakePythonBridge([
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

	it('uses project ByRef helper signatures for workbook return-assignment analysis', async () => {
		const bridge = fakePythonBridge([
			{
				name: 'Runner',
				type: 'standard',
				source:
					'Option Explicit\n' +
					'Public Function Run()\n' +
					'    Helpers.CopyVariant Run, 1\n' +
					'End Function\n',
			},
			{
				name: 'Helpers',
				type: 'standard',
				source:
					'Option Explicit\n' +
					'Public Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)\n' +
					'    dest = value\n' +
					'End Sub\n',
			},
		]);

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');

		expect(result.problems.filter((item) => item.code === 'missing-return-assignment')).toEqual([]);
	});

	it('uses source bindings before host globals in workbook member diagnostics', async () => {
		const bridge = fakePythonBridge([
			{
				name: 'Caller',
				type: 'standard',
				source:
					'Option Explicit\n' +
					'Public Sub T()\n' +
					'    Dim ActiveSheet\n' +
					'    Set ActiveSheet = ThisWorkbook\n' +
					'    ActiveSheet.DoesNotExist\n' +
					'    Dim typedSheet As Person\n' +
					'    typedSheet.Range\n' +
					'End Sub\n',
			},
			{
				name: 'Person',
				type: 'class',
				source:
					'Option Explicit\n' +
					'Public Sub Save()\n' +
					'End Sub\n',
			},
		]);

		const result = await analyzeWorkbook(bridge, 'Book.xlsm');
		const memberHits = result.problems.filter((item) => item.code === 'member-not-found');

		expect(memberHits).toHaveLength(1);
		expect(memberHits[0]).toMatchObject({
			moduleName: 'Caller',
			severity: 'error',
		});
		expect(memberHits[0].message).toContain('Person.Range');
		expect(result.problems.map((item) => item.message).join('\n')).not.toContain('Excel.Worksheet.DoesNotExist');
	});
});
