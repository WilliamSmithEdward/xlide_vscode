import { vi } from 'vitest';
import type { PythonBridge } from '../../src/pythonBridge';

export interface FakeBridgeModule {
	name: string;
	type: string;
	documentType?: string;
	source: string;
}

export interface FakeBridgeOptions {
	/** When false, readModules rejects so callers exercise the listModules/readModule fallback. */
	supportsBatchRead?: boolean;
}

/** Single-workbook module list (any path) or per-workbook-path module lists. */
type FakeBridgeWorkbooks = FakeBridgeModule[] | Record<string, FakeBridgeModule[]>;

/**
 * Builds a fake PythonBridge covering the module-read RPC surface
 * (readModules/listModules/readModule). Calls are recorded via vi.fn so tests
 * can assert on the call sequence.
 */
export function fakePythonBridge(
	workbooks: FakeBridgeWorkbooks,
	options: FakeBridgeOptions = {},
): PythonBridge {
	const modulesFor = (workbookPath: string): FakeBridgeModule[] | undefined =>
		Array.isArray(workbooks) ? workbooks : workbooks[workbookPath];
	return {
		call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
			const modules = modulesFor(payload.path);
			if (method === 'readModules') {
				if (options.supportsBatchRead === false) {
					throw new Error('Method not found: readModules');
				}
				if (!modules) {
					throw new Error(`Unknown workbook ${payload.path}`);
				}
				return modules.map(({ name, type, documentType, source }) => ({ name, type, documentType, source }));
			}
			if (method === 'listModules') {
				if (!modules) {
					throw new Error(`Unknown workbook ${payload.path}`);
				}
				return modules.map(({ name, type, documentType }) => ({ name, type, documentType }));
			}
			if (method === 'readModule' && payload.module) {
				const moduleName = payload.module.toLowerCase();
				const mod = modules?.find((candidate) => candidate.name.toLowerCase() === moduleName);
				if (!mod) {
					throw new Error(`Unknown module ${payload.path}::${payload.module}`);
				}
				return { source: mod.source };
			}
			throw new Error(`Unexpected bridge call ${method}`);
		}),
	} as unknown as PythonBridge;
}
