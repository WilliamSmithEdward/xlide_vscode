import { vi } from 'vitest';
import type { WorkbookEngine } from '../../src/workbookEngine';
import { WorkbookEngineError, JSONRPC_METHOD_NOT_FOUND } from '../../src/workbookEngineErrors';

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
 * Builds a fake WorkbookEngine covering the module-read RPC surface
 * (readModules/listModules/readModule). Calls are recorded via vi.fn so tests
 * can assert on the call sequence.
 */
export function fakeWorkbookEngine(
	workbooks: FakeBridgeWorkbooks,
	options: FakeBridgeOptions = {},
): WorkbookEngine {
	const modulesFor = (workbookPath: string): FakeBridgeModule[] | undefined =>
		Array.isArray(workbooks) ? workbooks : workbooks[workbookPath];
	return {
		call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
			const modules = modulesFor(payload.path);
			if (method === 'readModules') {
				if (options.supportsBatchRead === false) {
					throw new WorkbookEngineError('Method not found: readModules', JSONRPC_METHOD_NOT_FOUND);
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
	} as unknown as WorkbookEngine;
}
