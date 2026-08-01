import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import type { WorkbookEngine } from '../src/workbookEngine';
import { BridgeError, JSONRPC_METHOD_NOT_FOUND } from '../src/workbookEngineErrors';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';
import { fakeWorkbookEngine } from './helpers/fakeWorkbookEngine';
import { deferred, flushPromises } from './helpers/async';

describe('VbaSymbolIndex workbook identity', () => {
	it('keeps identical module names siloed by workbook path', async () => {
		const bridge = fakeWorkbookEngine({
			'C:/One/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub FromOne()\nEnd Sub\n' }],
			'C:/Two/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub FromTwo()\nEnd Sub\n' }],
		});
		const index = new VbaSymbolIndex(bridge);

		const first = await index.getModule('C:/One/Book.xlsm', 'Module1');
		const second = await index.getModule('C:/Two/Book.xlsm', 'Module1');

		expect(first.source).toContain('FromOne');
		expect(second.source).toContain('FromTwo');
	});

	it('uses one case-insensitive module cache key within a workbook', async () => {
		const bridge = fakeWorkbookEngine({
			'C:/Book.xlsm': [{ name: 'Module1', type: 'standard', source: 'Sub Cached()\nEnd Sub\n' }],
		});
		const index = new VbaSymbolIndex(bridge);

		await index.getModule('C:/Book.xlsm', 'Module1');
		await index.getModule('C:/Book.xlsm', 'module1');

		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);
	});

	it('updates a cached module directly from saved editor text', async () => {
		const bridge = fakeWorkbookEngine({});
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
		} as unknown as WorkbookEngine;
		const index = new VbaSymbolIndex(bridge);

		const first = index.getModule('C:/Book.xlsm', 'Module1');
		const second = index.getModule('C:/Book.xlsm', 'module1');
		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);

		read.resolve({ source: 'Sub Shared()\nEnd Sub\n' });
		const [firstModule, secondModule] = await Promise.all([first, second]);

		expect(firstModule).toBe(secondModule);
		expect(firstModule.source).toContain('Shared');
	});

	it('shares workbook indexing and reuses the cached module list', async () => {
		const bridge = fakeWorkbookEngine({
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

	it('falls back to list/read calls when batch workbook reads are unavailable', async () => {
		const bridge = fakeWorkbookEngine([
			{ name: 'Module1', type: 'standard', source: 'Sub Module1Proc()\nEnd Sub\n' },
			{ name: 'Module2', type: 'standard', source: 'Sub Module2Proc()\nEnd Sub\n' },
		], { supportsBatchRead: false });
		const index = new VbaSymbolIndex(bridge);

		const modules = await index.getAllModules('C:/Book.xlsm');

		expect(modules.map((mod) => mod.moduleName)).toEqual(['Module1', 'Module2']);
		expect(vi.mocked(bridge.call).mock.calls.map(([method]) => method)).toEqual([
			'readModules',
			'listModules',
			'readModule',
			'readModule',
		]);
	});

	it('reads fallback workbook modules concurrently', async () => {
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
		} as unknown as WorkbookEngine;
		const index = new VbaSymbolIndex(bridge);

		const pending = index.getAllModules('C:/Book.xlsm');
		await flushPromises();

		expect(vi.mocked(bridge.call).mock.calls.filter(([method]) => method === 'readModule')).toHaveLength(2);

		module1.resolve({ source: 'Sub First()\nEnd Sub\n' });
		module2.resolve({ source: 'Sub Second()\nEnd Sub\n' });

		await expect(pending).resolves.toMatchObject([
			{ moduleName: 'Module1' },
			{ moduleName: 'Module2' },
		]);
	});

	it('keeps saved editor text when a stale bridge read finishes late', async () => {
		const read = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((_method: string, _payload: { path: string; module?: string }) => read.promise),
		} as unknown as WorkbookEngine;
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
