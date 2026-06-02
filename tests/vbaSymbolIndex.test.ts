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

function bridgeForSources(sources: Record<string, string>): PythonBridge {
	return {
		call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
			if (method !== 'readModule' || !payload.module) {
				throw new Error(`Unexpected bridge call ${method}`);
			}
			const key = `${payload.path}::${payload.module.toLowerCase()}`;
			const source = sources[key];
			if (source === undefined) {
				throw new Error(`Unknown module ${key}`);
			}
			return { source };
		}),
	} as unknown as PythonBridge;
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
});
