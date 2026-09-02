import { vi } from 'vitest';
import type { ProjectEngine } from '../../src/projectEngine';

export interface FakeBridgeModule {
	name: string;
	type: string;
	documentType?: string;
	source: string;
	/** A VB6 project's module: the file it is. */
	filePath?: string;
}

/** Single-project module list (any path) or per-project-path module lists. */
type FakeBridgeProjects = FakeBridgeModule[] | Record<string, FakeBridgeModule[]>;

/**
 * Builds a fake ProjectEngine covering the module-read RPC surface
 * (readModules/listModules/readModule). Calls are recorded via vi.fn so tests
 * can assert on the call sequence.
 */
export function fakeProjectEngine(
	projects: FakeBridgeProjects,
): ProjectEngine {
	const modulesFor = (projectPath: string): FakeBridgeModule[] | undefined =>
		Array.isArray(projects) ? projects : projects[projectPath];
	return {
		call: vi.fn(async (method: string, payload: { path: string; module?: string }) => {
			const modules = modulesFor(payload.path);
			if (method === 'readModules') {
				if (!modules) {
					throw new Error(`Unknown workbook ${payload.path}`);
				}
				return modules.map(({ name, type, documentType, source, filePath }) => ({ name, type, documentType, source, filePath }));
			}
			if (method === 'listModules') {
				if (!modules) {
					throw new Error(`Unknown workbook ${payload.path}`);
				}
				return modules.map(({ name, type, documentType, filePath }) => ({ name, type, documentType, filePath }));
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
	} as unknown as ProjectEngine;
}
