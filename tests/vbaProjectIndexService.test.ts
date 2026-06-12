import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	workspace: {
		onDidCloseTextDocument: vi.fn(() => ({ dispose: () => undefined })),
	},
}));
vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));

import * as vscode from 'vscode';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';
import { VbaProjectIndexService } from '../src/vbaProjectIndexService';
import { fakePythonBridge, type FakeBridgeModule } from './helpers/fakePythonBridge';

const BOOK = 'C:/Book.xlsm';

function service(modules: FakeBridgeModule[]): {
	index: VbaSymbolIndex;
	projectIndexService: VbaProjectIndexService;
	callCount: () => number;
} {
	const bridge = fakePythonBridge({ [BOOK]: modules });
	const index = new VbaSymbolIndex(bridge);
	return {
		index,
		projectIndexService: new VbaProjectIndexService(index),
		callCount: () => vi.mocked(bridge.call).mock.calls.length,
	};
}

interface FakeOpenDocument {
	uri: { scheme: string; path: string; toString: () => string };
	version: number;
	getText: () => string;
}

function openXlideDocument(moduleName: string, source: string): FakeOpenDocument {
	const value = `xlide-vba:/${BOOK}/${moduleName}.bas`;
	return {
		uri: {
			scheme: 'xlide-vba',
			path: `/${BOOK}/${moduleName}.bas`,
			toString: () => value,
		},
		version: 1,
		getText: () => source,
	};
}

beforeEach(() => {
	(vscode.workspace.textDocuments as unknown[]).length = 0;
});

describe('VbaProjectIndexService', () => {
	it('builds one shared workbook project and reuses it across requests', async () => {
		const { projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
			{ name: 'Sheet1', type: 'document', source: 'Private Sub Worksheet_Activate()\nEnd Sub\n' },
		]);

		const first = await projectIndexService.contextForWorkbook(BOOK);
		const second = await projectIndexService.contextForWorkbook(BOOK);

		expect(second.project).toBe(first.project);
		expect(callCount()).toBe(1);
		expect(first.project.visibleProcedureNames('Sheet1').has('alpha')).toBe(true);
		expect(first.byModule.get('module1')?.source).toContain('Alpha');
		expect(first.moduleMetadata.get('sheet1')?.moduleKind).toBe('document');
		expect(first.modules.map((mod) => mod.moduleName)).toEqual(['Module1', 'Sheet1']);
	});

	it('folds saved module changes incrementally without a workbook rebuild', async () => {
		const { index, projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
			{ name: 'Module2', type: 'standard', source: 'Public Sub Beta()\nEnd Sub\n' },
		]);
		const first = await projectIndexService.contextForWorkbook(BOOK);

		index.updateModuleSource(BOOK, 'Module1', 'Public Sub Gamma()\nEnd Sub\n');
		const second = await projectIndexService.contextForWorkbook(BOOK);

		expect(second.project).toBe(first.project);
		expect(callCount()).toBe(1);
		expect(second.project.visibleProcedureNames('Module2').has('gamma')).toBe(true);
		expect(second.project.visibleProcedureNames('Module2').has('alpha')).toBe(false);
		expect(second.byModule.get('module1')?.source).toContain('Gamma');
	});

	it('rebuilds after a workbook-level symbol index invalidation', async () => {
		const { index, projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const first = await projectIndexService.contextForWorkbook(BOOK);

		index.invalidate(BOOK);
		const second = await projectIndexService.contextForWorkbook(BOOK);

		expect(second.project).not.toBe(first.project);
		expect(callCount()).toBe(2);
	});

	it('applies open-document text per version on access', async () => {
		const { projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const document = openXlideDocument('Module1', 'Public Sub FromEditor()\nEnd Sub\n');
		(vscode.workspace.textDocuments as unknown[]).push(document);

		const first = await projectIndexService.contextForWorkbook(BOOK);
		expect(first.byModule.get('module1')?.source).toContain('FromEditor');
		expect(first.project.visibleProcedureNames('Module1').has('fromeditor')).toBe(true);

		document.version = 2;
		document.getText = () => 'Public Sub Edited()\nEnd Sub\n';
		const second = await projectIndexService.contextForWorkbook(BOOK);

		expect(second.project).toBe(first.project);
		expect(callCount()).toBe(1);
		expect(second.byModule.get('module1')?.source).toContain('Edited');
		expect(second.project.visibleProcedureNames('Module1').has('edited')).toBe(true);
	});

	it('serves last-good content in live mode and rethrows in strict mode', async () => {
		// A non-string source is the one input that makes the module build
		// throw (the VBA parser itself recovers from any malformed text), so
		// it stands in for an internal indexing failure here.
		const { index, projectIndexService } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
			{ name: 'Bad', type: 'standard', source: undefined as unknown as string },
		]);

		const live = await projectIndexService.contextForWorkbook(BOOK, 'live');
		expect(live.byModule.has('bad')).toBe(false);
		expect(live.moduleMetadata.has('bad')).toBe(true);
		expect(live.project.visibleProcedureNames('Bad').has('alpha')).toBe(true);

		await expect(projectIndexService.contextForWorkbook(BOOK, 'strict')).rejects.toBeTruthy();

		index.updateModuleSource(BOOK, 'Bad', 'Public Sub Fixed()\nEnd Sub\n');
		const strict = await projectIndexService.contextForWorkbook(BOOK, 'strict');
		expect(strict.byModule.get('bad')?.source).toContain('Fixed');
		expect(strict.project.visibleProcedureNames('Module1').has('fixed')).toBe(true);
	});

	it('resets the consumer-memoized procedure signatures on module change', async () => {
		const { index, projectIndexService } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const context = await projectIndexService.contextForWorkbook(BOOK);
		context.projectProcedures = new Map();

		index.updateModuleSource(BOOK, 'Module1', 'Public Sub Beta()\nEnd Sub\n');
		const next = await projectIndexService.contextForWorkbook(BOOK);

		expect(next.projectProcedures).toBeUndefined();
	});
});
