import { describe, expect, it } from 'vitest';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

const MOD_A = [
	'Option Explicit',
	'Sub Alpha()',
	'    undeclaredThing = 1',
	'End Sub',
].join('\n');

const MOD_B = [
	'Option Explicit',
	'Public Sub Helper(ByVal n As Long)',
	'End Sub',
].join('\n');

describe('AnalysisWorkerState', () => {
	it('analyzes with seeded project context and matches the in-host pass', () => {
		const state = new AnalysisWorkerState();
		state.handle({
			kind: 'seed',
			workbookKey: 'wb1',
			generation: 1,
			modules: [
				{ moduleName: 'ModA', source: MOD_A, type: 'standard' },
				{ moduleName: 'ModB', source: MOD_B, type: 'standard' },
			],
		});
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: MOD_A, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		expect(response.diagnostics.some((d) => d.message.includes('undeclaredThing'))).toBe(true);
		// Second analyze with a body-only change engages incremental.
		const edited = MOD_A.replace('    undeclaredThing = 1', '    undeclaredThing = 1\n    alsoUndeclared = 2');
		const second = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: edited, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(second?.kind).toBe('result');
		if (second?.kind !== 'result') { return; }
		expect(second.incrementalMode).toBe('incremental');
		expect(second.diagnostics.some((d) => d.message.includes('alsoUndeclared'))).toBe(true);
	});

	it('requests a reseed for an unknown or stale workbook generation', () => {
		const state = new AnalysisWorkerState();
		const unknown = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'nope', generation: 1,
			source: MOD_A, moduleName: 'ModA',
		});
		expect(unknown?.kind).toBe('needSeed');
		state.handle({ kind: 'seed', workbookKey: 'wb1', generation: 1, modules: [{ moduleName: 'ModA', source: MOD_A }] });
		const stale = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 2,
			source: MOD_A, moduleName: 'ModA',
		});
		expect(stale?.kind).toBe('needSeed');
	});

	it('analyzes standalone (no workbook) and matches direct analysis', () => {
		const state = new AnalysisWorkerState();
		const response = state.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc1',
			source: MOD_A, moduleName: 'ModA', moduleType: 'standard',
		});
		expect(response?.kind).toBe('result');
		if (response?.kind !== 'result') { return; }
		const direct = analyzeVbaModuleSource({ source: MOD_A, moduleName: 'ModA', moduleType: 'standard' });
		expect(response.diagnostics.map((d) => `${d.code}:${d.span.start}`).sort())
			.toEqual(direct.diagnostics.map((d) => `${d.code}:${d.span.start}`).sort());
	});

	it('forget clears per-document incremental state', () => {
		const state = new AnalysisWorkerState();
		state.handle({ kind: 'seed', workbookKey: 'wb1', generation: 1, modules: [{ moduleName: 'ModA', source: MOD_A }] });
		state.handle({ kind: 'analyze', requestId: 1, docKey: 'doc1', workbookKey: 'wb1', generation: 1, source: MOD_A, moduleName: 'ModA' });
		state.handle({ kind: 'forget', docKey: 'doc1' });
		const after = state.handle({
			kind: 'analyze', requestId: 2, docKey: 'doc1', workbookKey: 'wb1', generation: 1,
			source: MOD_A.replace('= 1', '= 2'), moduleName: 'ModA',
		});
		expect(after?.kind).toBe('result');
		if (after?.kind !== 'result') { return; }
		expect(after.incrementalMode).toBe('full');
	});
});
