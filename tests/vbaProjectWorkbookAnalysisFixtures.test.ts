import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
	workspace: {
		getConfiguration: () => ({
			get: (_key: string, fallback: unknown) => fallback,
		}),
		textDocuments: [],
	},
}));

import { analyzeWorkbook } from '../src/vbaWorkbookAnalysis';
import type { PythonBridge } from '../src/pythonBridge';
import {
	fixtureModules,
	loadVbaProjectFixtures,
	type VbaProjectFixture,
} from './helpers/vbaProjectFixtures';

function bridgeForFixture(fixture: VbaProjectFixture): PythonBridge {
	const modules = fixtureModules(fixture);
	const byName = new Map(modules.map((mod) => [mod.moduleName, mod]));
	return {
		call: vi.fn(async (method: string, payload: { module?: string }) => {
			if (method === 'listModules') {
				return modules.map((mod) => ({
					name: mod.moduleName,
					type: mod.type ?? 'standard',
					documentType: mod.documentType,
				}));
			}
			if (method === 'readModule' && payload.module) {
				const mod = byName.get(payload.module);
				if (!mod) {
					throw new Error(`Unknown module ${payload.module}`);
				}
				return { source: mod.source };
			}
			throw new Error(`Unexpected bridge call ${method}`);
		}),
	} as unknown as PythonBridge;
}

describe('machine-readable VBA workbook analysis fixtures', () => {
	for (const fixture of loadVbaProjectFixtures().filter((item) => item.assertions.workbookAnalysis)) {
		it(`matches workbook analysis expectations for ${fixture.id}`, async () => {
			const assertion = fixture.assertions.workbookAnalysis;
			const result = await analyzeWorkbook(bridgeForFixture(fixture), `${fixture.id}.xlsm`);
			const problemSummary = result.problems
				.map((problem) => `${problem.moduleName}:${problem.line}:${problem.code}: ${problem.message}`)
				.join('\n');

			if (assertion?.problemCount !== undefined) {
				expect(result.problems, problemSummary).toHaveLength(assertion.problemCount);
			}
			for (const codeAssertion of assertion?.codes ?? []) {
				const hits = result.problems.filter((problem) => problem.code === codeAssertion.code);
				expect(hits, `${codeAssertion.code}\n${problemSummary}`).toHaveLength(codeAssertion.count);
				const messages = hits.map((hit) => hit.message).join('\n');
				for (const text of codeAssertion.messagesContain ?? []) {
					expect(messages, text).toContain(text);
				}
			}
			for (const code of assertion?.absentCodes ?? []) {
				expect(result.problems.filter((problem) => problem.code === code), code).toHaveLength(0);
			}
		});
	}
});
