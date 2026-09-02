import * as path from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import type * as VscodeType from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import * as vscode from 'vscode';
import { analyzeProject } from '../src/vbaProjectWideAnalysis';
import type { ProjectEngine } from '../src/projectEngine';
import type { ProjectAnalysisProblem } from '../src/vbaProjectWideAnalysis';
import {
	fixtureModules,
	loadVbaProjectFixtures,
	type VbaProjectFixtureProblemAssertion,
	type VbaProjectFixtureOpenDocumentAssertion,
	type VbaProjectFixture,
} from './helpers/vbaProjectFixtures';
import { fakeProjectEngine } from './helpers/fakeProjectEngine';

function bridgeForFixture(fixture: VbaProjectFixture): ProjectEngine {
	return fakeProjectEngine(fixtureModules(fixture).map((mod) => ({
		name: mod.moduleName,
		type: mod.type ?? 'standard',
		documentType: mod.documentType,
		source: mod.source,
	})));
}

function fixtureWorkbookPath(fixture: VbaProjectFixture): string {
	return path.join(path.sep, 'fixtures', `${fixture.id}.xlsm`);
}

function documentPath(projectPath: string, moduleName: string): string {
	return `${projectPath.replace(/\\/g, '/')}/${encodeURIComponent(moduleName)}.bas`;
}

function openDocument(
	fixture: VbaProjectFixture,
	doc: VbaProjectFixtureOpenDocumentAssertion,
) {
	const projectPath = doc.projectPath ?? fixtureWorkbookPath(fixture);
	const uriPath = documentPath(projectPath, doc.moduleName);
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
	problems: readonly ProjectAnalysisProblem[],
	expected: VbaProjectFixtureProblemAssertion,
): ProjectAnalysisProblem[] {
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

describe('machine-readable VBA project analysis fixtures', () => {
	for (const fixture of loadVbaProjectFixtures().filter((item) => item.assertions.projectAnalysis)) {
		it(`matches project analysis expectations for ${fixture.id}`, async () => {
			const assertion = fixture.assertions.projectAnalysis;
			setOpenDocuments((assertion?.openDocuments ?? []).map((doc) => openDocument(fixture, doc)));
			let result: Awaited<ReturnType<typeof analyzeProject>> | undefined;
			try {
				result = await analyzeProject(bridgeForFixture(fixture), fixtureWorkbookPath(fixture));
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
