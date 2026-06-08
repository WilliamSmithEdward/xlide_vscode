import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => {
	class Disposable {
		dispose(): void { /* no-op */ }
	}
	class EventEmitter<T> {
		readonly event = () => new Disposable();
		fire(_value: T): void { /* no-op */ }
		dispose(): void { /* no-op */ }
	}
	return { EventEmitter };
});
vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));

import type { PythonBridge } from '../src/pythonBridge';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';

function bridgeForSources(
	sources: Record<string, string>,
	moduleLists: Record<string, Array<{ name: string; type: string }>> = {},
): PythonBridge {
	return {
		call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
			if (method === 'readModules') {
				const list = moduleLists[payload.path];
				if (!list) {
					throw new Error('Method not found: readModules');
				}
				return list.map((entry) => {
					const key = `${payload.path}::${entry.name.toLowerCase()}`;
					const source = sources[key];
					if (source === undefined) {
						throw new Error(`Unknown module ${key}`);
					}
					return { ...entry, source };
				});
			}
			if (method === 'listModules') {
				const list = moduleLists[payload.path];
				if (!list) {
					throw new Error(`Unknown workbook ${payload.path}`);
				}
				return list;
			}
			if (method === 'readModule' && payload.module) {
				const key = `${payload.path}::${payload.module.toLowerCase()}`;
				const source = sources[key];
				if (source === undefined) {
					throw new Error(`Unknown module ${key}`);
				}
				return { source };
			}
			throw new Error(`Unexpected bridge call ${method}`);
		}),
	} as unknown as PythonBridge;
}

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (error: unknown) => void } {
	let resolve!: (value: T) => void;
	let reject!: (error: unknown) => void;
	const promise = new Promise<T>((promiseResolve, promiseReject) => {
		resolve = promiseResolve;
		reject = promiseReject;
	});
	return { promise, resolve, reject };
}

describe('VbaSymbolIndex workbook identity', () => {
	it('keeps identical module names siloed by workbook path', async () => {
		const bridge = bridgeForSources({
			'C:/One/Book.xlsm::module1': 'Sub FromOne()\nEnd Sub\n',
			'C:/Two/Book.xlsm::module1': 'Sub FromTwo()\nEnd Sub\n',
		});
		const index = new VbaSymbolIndex(bridge);

		const first = await index.getModule('C:/One/Book.xlsm', 'Module1');
		const second = await index.getModule('C:/Two/Book.xlsm', 'Module1');

		expect(first.source).toContain('FromOne');
		expect(second.source).toContain('FromTwo');
		expect(first.symbols[0].name).toBe('FromOne');
		expect(second.symbols[0].name).toBe('FromTwo');
	});

	it('uses one case-insensitive module cache key within a workbook', async () => {
		const bridge = bridgeForSources({
			'C:/Book.xlsm::module1': 'Sub Cached()\nEnd Sub\n',
		});
		const index = new VbaSymbolIndex(bridge);

		await index.getModule('C:/Book.xlsm', 'Module1');
		await index.getModule('C:/Book.xlsm', 'module1');

		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);
	});

	it('updates a cached module directly from saved editor text', async () => {
		const bridge = bridgeForSources({});
		const index = new VbaSymbolIndex(bridge);

		index.updateModuleSource('C:/Book.xlsm', 'Module1', 'Sub Saved()\nEnd Sub\n');
		const mod = await index.getModule('C:/Book.xlsm', 'module1');

		expect(mod.source).toContain('Saved');
		expect(mod.symbols.map((symbol) => symbol.name)).toEqual(['Saved']);
		expect(vi.mocked(bridge.call)).not.toHaveBeenCalled();
	});

	it('shares concurrent reads for the same module', async () => {
		const read = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((_method: string, _payload: { path: string; module?: string }) => read.promise),
		} as unknown as PythonBridge;
		const index = new VbaSymbolIndex(bridge);

		const first = index.getModule('C:/Book.xlsm', 'Module1');
		const second = index.getModule('C:/Book.xlsm', 'module1');
		expect(vi.mocked(bridge.call)).toHaveBeenCalledTimes(1);

		read.resolve({ source: 'Sub Shared()\nEnd Sub\n' });
		const [firstModule, secondModule] = await Promise.all([first, second]);

		expect(firstModule).toBe(secondModule);
		expect(firstModule.symbols.map((symbol) => symbol.name)).toEqual(['Shared']);
	});

	it('shares workbook indexing and reuses the cached module list', async () => {
		const bridge = bridgeForSources(
			{
				'C:/Book.xlsm::module1': 'Sub First()\nEnd Sub\n',
				'C:/Book.xlsm::module2': 'Sub Second()\nEnd Sub\n',
			},
			{
				'C:/Book.xlsm': [
					{ name: 'Module1', type: 'standard' },
					{ name: 'Module2', type: 'standard' },
				],
			},
		);
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
		const bridge = {
			call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
				if (method === 'readModules') {
					throw new Error('Method not found: readModules');
				}
				if (method === 'listModules') {
					return [
						{ name: 'Module1', type: 'standard' },
						{ name: 'Module2', type: 'standard' },
					];
				}
				if (method === 'readModule' && payload.module) {
					return { source: `Sub ${payload.module}Proc()\nEnd Sub\n` };
				}
				throw new Error(`Unexpected bridge call ${method}`);
			}),
		} as unknown as PythonBridge;
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

	it('keeps saved editor text when a stale bridge read finishes late', async () => {
		const read = deferred<{ source: string }>();
		const bridge = {
			call: vi.fn((_method: string, _payload: { path: string; module?: string }) => read.promise),
		} as unknown as PythonBridge;
		const index = new VbaSymbolIndex(bridge);

		const pending = index.getModule('C:/Book.xlsm', 'Module1');
		index.updateModuleSource('C:/Book.xlsm', 'Module1', 'Sub Saved()\nEnd Sub\n');
		read.resolve({ source: 'Sub Stale()\nEnd Sub\n' });

		const resolved = await pending;
		const cached = await index.getModule('C:/Book.xlsm', 'module1');

		expect(resolved.source).toContain('Saved');
		expect(cached.source).toContain('Saved');
		expect(cached.symbols.map((symbol) => symbol.name)).toEqual(['Saved']);
	});
});
