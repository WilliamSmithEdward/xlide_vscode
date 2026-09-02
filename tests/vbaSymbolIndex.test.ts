import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import type { ProjectEngine } from '../src/projectEngine';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';
import { fakeProjectEngine } from './helpers/fakeProjectEngine';
import { deferred, flushPromises } from './helpers/async';

describe('VbaSymbolIndex project identity', () => {
	it('keeps identical module names siloed by project path', async () => {
		const bridge = fakeProjectEngine({
			'C:/One/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub FromOne()\nEnd Sub\n' }],
			'C:/Two/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub FromTwo()\nEnd Sub\n' }],
		});
		const index = new VbaSymbolIndex(bridge);

		const first = await index.getModule('C:/One/Book.xlsm', 'Module1');
		const second = await index.getModule('C:/Two/Book.xlsm', 'Module1');

		expect(first.source).toContain('FromOne');
		expect(second.source).toContain('FromTwo');
	});

	it('uses one case-insensitive module cache key within a project', async () => {
		const bridge = fakeProjectEngine({
			'C:/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub Cached()\nEnd Sub\n' }],
		});
		const index = new VbaSymbolIndex(bridge);

		await index.getModule('C:/Book.xlsm', 'Module1');
		await index.getModule('C:/Book.xlsm', 'module1');

		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);
	});

	it('updates a cached module directly from saved editor text', async () => {
		const bridge = fakeProjectEngine({});
		const index = new VbaSymbolIndex(bridge);

		index.updateModuleSource('C:/Book.xlsm', 'Module1', 'Sub Saved()\nEnd Sub\n');
		const mod = await index.getModule('C:/Book.xlsm', 'module1');

		expect(mod.source).toContain('Saved');
		expect(vi.mocked(bridge.call)).not.toHaveBeenCalled();
	});

	it('shares concurrent reads for the same module', async () => {
		const read = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((_method: string, _payload: { path: string; module?: string }) => read.promise),
		} as unknown as ProjectEngine;
		const index = new VbaSymbolIndex(bridge);

		const first = index.getModule('C:/Book.xlsm', 'Module1');
		const second = index.getModule('C:/Book.xlsm', 'module1');
		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);

		read.resolve({ source: 'Sub Shared()\nEnd Sub\n' });
		const [firstModule, secondModule] = await Promise.all([first, second]);

		expect(firstModule).toBe(secondModule);
		expect(firstModule.source).toContain('Shared');
	});

	it('shares project indexing and reuses the cached module list', async () => {
		const bridge = fakeProjectEngine({
			'C:/Book.xlsm': [
				{ name: 'Module1', type: 'standard', source: 'Sub First()\nEnd Sub\n' },
				{ name: 'Module2', type: 'standard', source: 'Sub Second()\nEnd Sub\n' },
			],
		});
		const index = new VbaSymbolIndex(bridge);

		const [first, second] = await Promise.all([
			index.getAllModules('C:/Book.xlsm'),
			index.getAllModules('C:/Book.xlsm'),
		]);
		const third = await index.getAllModules('C:/Book.xlsm');

		expect(first.map((mod) => mod.moduleName)).toEqual(['Module1', 'Module2']);
		expect(second.map((mod) => mod.moduleName)).toEqual(['Module1', 'Module2']);
		expect(third.map((mod) => mod.moduleName)).toEqual(['Module1', 'Module2']);
		expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'readModules')).toHaveLength(1);
		expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'listModules')).toHaveLength(0);
		expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'readModule')).toHaveLength(0);
	});

	it('keeps saved editor text when a stale bridge read finishes late', async () => {
		const read = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((_method: string, _payload: { path: string; module?: string }) => read.promise),
		} as unknown as ProjectEngine;
		const index = new VbaSymbolIndex(bridge);

		const pending = index.getModule('C:/Book.xlsm', 'Module1');
		index.updateModuleSource('C:/Book.xlsm', 'Module1', 'Sub Saved()\nEnd Sub\n');
		read.resolve({ source: 'Sub Stale()\nEnd Sub\n' });

		const resolved = await pending;
		const cached = await index.getModule('C:/Book.xlsm', 'module1');

		expect(resolved.source).toContain('Saved');
		expect(cached.source).toContain('Saved');
	});
});
