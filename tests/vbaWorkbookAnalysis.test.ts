import { afterEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { analyzeWorkbook, setWorkbookAnalysisWorker, type WorkbookAnalysisWorker } from '../src/vbaWorkbookAnalysis';
import type { WorkbookEngine } from '../src/workbookEngine';
import { fakeWorkbookEngine } from './helpers/fakeWorkbookEngine';
import { deferred, flushPromises } from './helpers/async';

describe('analyzeWorkbook metadata summary', () => {
	it('loads workbook modules through the batch read endpoint', async () => {
		const bridge = fakeWorkbookEngine([
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

	it('attaches shared rule metadata to structural and semantic workbook problems', async () => {
		const bridge = fakeWorkbookEngine([
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
		const bridge = fakeWorkbookEngine([
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
		const problem = result.problems.find((item) => item.code === 'mismatched-end-keyword');

		expect(problem).toMatchObject({
			severity: 'warning',
			expectedClose: 'End Property',
			quickFixAvailable: true,
			quickFixTitles: ["Replace 'End Function' with 'End Property'"],
		});
	});

	it('uses project ByRef helper signatures for workbook return-assignment analysis', async () => {
		const bridge = fakeWorkbookEngine([
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
		const bridge = fakeWorkbookEngine([
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

describe('analyzeWorkbook worker routing', () => {
	afterEach(() => {
		setWorkbookAnalysisWorker(undefined);
	});

	const MODULES = [
		{ name: 'ModA', type: 'standard', source: 'Option Explicit\nSub A()\n\tmissingA = 1\nEnd Sub\n' },
		{ name: 'ModB', type: 'standard', source: 'Option Explicit\nSub B()\n\tmissingB = 1\nEnd Sub\n' },
	];

	function cannedWorker() {
		const analyzed: string[] = [];
		const seeded: Array<{ key: string; generation: number; count: number }> = [];
		const worker: WorkbookAnalysisWorker = {
			available: true,
			ensureSeeded(key, generation, modules) {
				seeded.push({ key, generation, count: modules().length });
			},
			analyze(request) {
				analyzed.push(request.moduleName);
				return Promise.resolve({
					diagnostics: [{
						code: 'undeclared-variable',
						message: `worker saw ${request.moduleName}`,
						severity: 'error' as const,
						span: { start: 0, end: 1 },
					}],
					suppressedDiagnostics: [{
						code: 'undeclared-variable',
						message: `worker suppressed ${request.moduleName}`,
						severity: 'error' as const,
						span: { start: 0, end: 1 },
					}],
				});
			},
		};
		return { worker, analyzed, seeded };
	}

	it('routes every module through the worker and keeps suppressed findings', async () => {
		const { worker, analyzed, seeded } = cannedWorker();
		setWorkbookAnalysisWorker(worker);

		const result = await analyzeWorkbook(fakeWorkbookEngine(MODULES), 'Book.xlsm');

		expect(analyzed.sort()).toEqual(['ModA', 'ModB']);
		expect(seeded).toHaveLength(1);
		expect(seeded[0].count).toBe(2);
		expect(seeded[0].key).toContain('workbook-analysis:');
		// The worker's findings are what the result reports - the in-host
		// analyzer (which would say "missingA") never ran.
		expect(result.problems.map((p) => p.message).sort()).toEqual([
			'worker saw ModA',
			'worker saw ModB',
		]);
		expect(result.suppressedProblems.map((p) => p.message).sort()).toEqual([
			'worker suppressed ModA',
			'worker suppressed ModB',
		]);
	});

	it('falls back to the identical in-host pass when the worker rejects', async () => {
		const worker: WorkbookAnalysisWorker = {
			available: true,
			ensureSeeded() { /* accepted */ },
			analyze() { return Promise.reject(new Error('worker crashed')); },
		};
		setWorkbookAnalysisWorker(worker);

		const result = await analyzeWorkbook(fakeWorkbookEngine(MODULES), 'Book.xlsm');

		// Real analysis findings, produced in-host.
		expect(result.problems.some((p) => p.message.includes('missingA'))).toBe(true);
		expect(result.problems.some((p) => p.message.includes('missingB'))).toBe(true);
	});

	it('never touches an unavailable worker', async () => {
		const analyzed: string[] = [];
		const worker: WorkbookAnalysisWorker = {
			available: false,
			ensureSeeded() { analyzed.push('seed'); },
			analyze(request) {
				analyzed.push(request.moduleName);
				return Promise.resolve({ diagnostics: [], suppressedDiagnostics: [] });
			},
		};
		setWorkbookAnalysisWorker(worker);

		const result = await analyzeWorkbook(fakeWorkbookEngine(MODULES), 'Book.xlsm');

		expect(analyzed).toEqual([]);
		expect(result.problems.some((p) => p.message.includes('missingA'))).toBe(true);
	});

	it('seeds with the same fingerprint for unchanged sources across runs', async () => {
		const { worker, seeded } = cannedWorker();
		setWorkbookAnalysisWorker(worker);

		await analyzeWorkbook(fakeWorkbookEngine(MODULES), 'Book.xlsm');
		await analyzeWorkbook(fakeWorkbookEngine(MODULES), 'Book.xlsm');
		expect(seeded).toHaveLength(2);
		expect(seeded[0].generation).toBe(seeded[1].generation);

		const edited = MODULES.map((m) => m.name === 'ModA'
			? { ...m, source: m.source.replace('missingA', 'missingEdited') }
			: m);
		await analyzeWorkbook(fakeWorkbookEngine(edited), 'Book.xlsm');
		expect(seeded[2].generation).not.toBe(seeded[0].generation);
	});
});

describe('analyzeWorkbook progress', () => {
	afterEach(() => {
		setWorkbookAnalysisWorker(undefined);
	});

	it('reports each module completion, forced past the start-report throttle', async () => {
		// All start reports fire within milliseconds now that analysis is
		// async, so the throttle drops them and only completion reports keep
		// the notification moving.
		const worker: WorkbookAnalysisWorker = {
			available: true,
			ensureSeeded() { /* accepted */ },
			analyze() { return Promise.resolve({ diagnostics: [], suppressedDiagnostics: [] }); },
		};
		setWorkbookAnalysisWorker(worker);
		const messages: string[] = [];

		await analyzeWorkbook(fakeWorkbookEngine([
			{ name: 'ModA', type: 'standard', source: 'Option Explicit\nSub A()\nEnd Sub\n' },
			{ name: 'ModB', type: 'standard', source: 'Option Explicit\nSub B()\nEnd Sub\n' },
		]), 'Book.xlsm', { progress: (message) => messages.push(message) });

		const done = messages.filter((m) => m.startsWith('Analyzed '));
		expect(done).toHaveLength(2);
		expect(done.some((m) => m.includes('(1/2)'))).toBe(true);
		expect(done.some((m) => m.includes('(2/2)'))).toBe(true);
	});
});
