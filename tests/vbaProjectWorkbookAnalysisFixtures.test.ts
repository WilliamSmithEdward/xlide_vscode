import * as path from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import type * as VscodeType from 'vscode';

vi.mock('vscode', () => ({
	workspace: {
		getConfiguration: () => ({
			get: (_key: string, fallback: unknown) => fallback,
		}),
		textDocuments: [],
	},
}));

import * as vscode from 'vscode';
import { analyzeWorkbook } from '../src/vbaWorkbookAnalysis';
import type { PythonBridge } from '../src/pythonBridge';
import type { WorkbookAnalysisProblem } from '../src/vbaWorkbookAnalysis';
import {
	fixtureModules,
	loadVbaProjectFixtures,
	type VbaProjectFixtureWorkbookProblemAssertion,
	type VbaProjectFixtureOpenDocumentAssertion,
	type VbaProjectFixture,
} from './helpers/vbaProjectFixtures';

function bridgeForFixture(fixture: VbaProjectFixture): PythonBridge {
	const modules = fixtureModules(fixture);
	const byName = new Map(modules.map((mod) => [mod.moduleName, mod]));
	return {
		call: vi.fn(async (method: string, payload: { module?: string }) => {
			if (method === 'readModules') {
				return modules.map((mod) => ({
					name: mod.moduleName,
					type: mod.type ?? 'standard',
					documentType: mod.documentType,
					source: mod.source,
				}));
			}
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

function fixtureWorkbookPath(fixture: VbaProjectFixture): string {
	return path.join(path.sep, 'fixtures', `${fixture.id}.xlsm`);
}

function documentPath(workbookPath: string, moduleName: string): string {
	return `${workbookPath.replace(/\\/g, '/')}/${encodeURIComponent(moduleName)}.bas`;
}

function openDocument(
	fixture: VbaProjectFixture,
	doc: VbaProjectFixtureOpenDocumentAssertion,
) {
	const workbookPath = doc.workbookPath ?? fixtureWorkbookPath(fixture);
	const uriPath = documentPath(workbookPath, doc.moduleName);
	return {
		uri: {
			scheme: 'xlide-vba',
			path: uriPath,
			toString: () => `xlide-vba:${uriPath}`,
		} as VscodeType.Uri,
		getText: () => doc.sourceLines.join('\n'),
	};
}

function setOpenDocuments(documents: ReturnType<typeof openDocument>[]): void {
	(vscode.workspace as unknown as { textDocuments: typeof documents }).textDocuments = documents;
}

function matchingProblems(
	problems: readonly WorkbookAnalysisProblem[],
	expected: VbaProjectFixtureWorkbookProblemAssertion,
): WorkbookAnalysisProblem[] {
	return problems.filter((problem) => {
		if (expected.moduleName && problem.moduleName !== expected.moduleName) {
			return false;
		}
		if (expected.code && problem.code !== expected.code) {
			return false;
		}
		if (expected.line !== undefined && problem.line !== expected.line) {
			return false;
		}
		if (expected.column !== undefined && problem.column !== expected.column) {
			return false;
		}
		if (expected.severity && problem.severity !== expected.severity) {
			return false;
		}
		for (const text of expected.messageContains ?? []) {
			if (!problem.message.includes(text)) {
				return false;
			}
		}
		for (const title of expected.quickFixTitles ?? []) {
			if (!(problem.quickFixTitles ?? []).includes(title)) {
				return false;
			}
		}
		return true;
	});
}

describe('machine-readable VBA workbook analysis fixtures', () => {
	for (const fixture of loadVbaProjectFixtures().filter((item) => item.assertions.workbookAnalysis)) {
		it(`matches workbook analysis expectations for ${fixture.id}`, async () => {
			const assertion = fixture.assertions.workbookAnalysis;
			setOpenDocuments((assertion?.openDocuments ?? []).map((doc) => openDocument(fixture, doc)));
			let result: Awaited<ReturnType<typeof analyzeWorkbook>> | undefined;
			try {
				result = await analyzeWorkbook(bridgeForFixture(fixture), fixtureWorkbookPath(fixture));
			} finally {
				setOpenDocuments([]);
			}
			expect(result).toBeDefined();
			const analysis = result!;
			const problemSummary = analysis.problems
				.map((problem) => `${problem.moduleName}:${problem.line}:${problem.code}: ${problem.message}`)
				.join('\n');

			if (assertion?.problemCount !== undefined) {
				expect(analysis.problems, problemSummary).toHaveLength(assertion.problemCount);
			}
			for (const expected of assertion?.problems ?? []) {
				expect(matchingProblems(analysis.problems, expected), `${JSON.stringify(expected)}\n${problemSummary}`)
					.toHaveLength(1);
			}
			for (const codeAssertion of assertion?.codes ?? []) {
				const hits = analysis.problems.filter((problem) => problem.code === codeAssertion.code);
				expect(hits, `${codeAssertion.code}\n${problemSummary}`).toHaveLength(codeAssertion.count);
				const messages = hits.map((hit) => hit.message).join('\n');
				for (const text of codeAssertion.messagesContain ?? []) {
					expect(messages, text).toContain(text);
				}
			}
			for (const code of assertion?.absentCodes ?? []) {
				expect(analysis.problems.filter((problem) => problem.code === code), code).toHaveLength(0);
			}
		});
	}
});
