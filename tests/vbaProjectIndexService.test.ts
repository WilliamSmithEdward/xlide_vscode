import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	workspace: {
		onDidCloseTextDocument: vi.fn(() => ({ dispose: () => undefined })),
	},
}));

import * as vscode from 'vscode';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';
import { VbaProjectIndexService } from '../src/vbaProjectIndexService';
import { fakeProjectEngine, type FakeBridgeModule } from './helpers/fakeProjectEngine';
import { ownersFromListings, setVb6ModuleOwnersForTests } from '../src/vb6ProjectLocator';

// Platform-appropriate absolute path: decodeModuleUri and path.resolve are
// deliberately platform-sensitive, so a hardcoded Windows path never matches
// on POSIX runners.
const BOOK = process.platform === 'win32' ? 'C:/Book.xlsm' : '/work/Book.xlsm';
const BOOK_URI_PATH = BOOK.startsWith('/') ? BOOK : `/${BOOK}`;

function service(modules: FakeBridgeModule[]): {
	index: VbaSymbolIndex;
	projectIndexService: VbaProjectIndexService;
	callCount: () => number;
} {
	const bridge = fakeProjectEngine({ [BOOK]: modules });
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
	const value = `xlide-vba:${BOOK_URI_PATH}/${moduleName}.bas`;
	return {
		uri: {
			scheme: 'xlide-vba',
			path: `${BOOK_URI_PATH}/${moduleName}.bas`,
			toString: () => value,
		},
		version: 1,
		getText: () => source,
	};
}

beforeEach(() => {
	(vscode.workspace.textDocuments as unknown[]).length = 0;
});

describe('a VB6 project in the project index', () => {
	const VBP = process.platform === 'win32' ? 'C:/proj/App.vbp' : '/proj/App.vbp';
	const FRM = process.platform === 'win32' ? 'C:/proj/Form1.frm' : '/proj/Form1.frm';
	const BAS = process.platform === 'win32' ? 'C:/proj/modMain.bas' : '/proj/modMain.bas';

	it('folds an open VB6 file into its project, designer header blanked, at its own offsets', async () => {
		const header = 'VERSION 5.00\r\nBegin VB.Form Form1 \r\n   Caption = "F"\r\n   BeginProperty Font \r\n      Name = "Tahoma"\r\n   EndProperty\r\nEnd\r\n';
		const code = 'Attribute VB_Name = "Form1"\r\nPrivate Sub Form_Load()\r\nEnd Sub\r\n';
		const bridge = fakeProjectEngine({ [VBP]: [
			{ name: 'Form1', type: 'userform', source: header.replace(/[^\r\n]/g, ' ') + code, filePath: FRM },
			{ name: 'modMain', type: 'standard', source: 'Attribute VB_Name = "modMain"\r\nPublic Sub Alpha()\r\nEnd Sub\r\n', filePath: BAS },
		] });
		setVb6ModuleOwnersForTests(ownersFromListings([
			{ vbpPath: VBP, modules: [
				{ name: 'Form1', type: 'userform', filePath: FRM },
				{ name: 'modMain', type: 'standard', filePath: BAS },
			] },
		]));
		const index = new VbaSymbolIndex(bridge);
		const projectIndexService = new VbaProjectIndexService(index);

		// The editor holds the WHOLE file, header included, with one new procedure.
		const edited = header + code + 'Private Sub Command1_Click()\r\nEnd Sub\r\n';
		(vscode.workspace.textDocuments as unknown[]).push({
			uri: { scheme: 'file', fsPath: FRM.replace(/\//g, process.platform === 'win32' ? '\\' : '/'), path: FRM.startsWith('/') ? FRM : `/${FRM}`, toString: () => `file:///${FRM}` },
			version: 2,
			getText: () => edited,
		});

		const context = await projectIndexService.contextForProject(VBP, 'live');
		const form = context.byModule.get('form1');
		expect(form?.source.length).toBe(edited.length);
		expect(form?.source).not.toMatch(/Begin VB\.Form/);
		expect(form?.source).toMatch(/Private Sub Command1_Click/);
		expect(context.byModule.get('modmain')?.filePath).toBe(BAS);
		expect(context.project.visibleProcedureNames('Form1').has('alpha')).toBe(true);
		setVb6ModuleOwnersForTests(new Map());
	});
});

describe('VbaProjectIndexService', () => {
	it('builds one shared project index and reuses it across requests', async () => {
		const { projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
			{ name: 'Sheet1', type: 'document', source: 'Private Sub Worksheet_Activate()\nEnd Sub\n' },
		]);

		const first = await projectIndexService.contextForProject(BOOK);
		const second = await projectIndexService.contextForProject(BOOK);

		expect(second.project).toBe(first.project);
		expect(callCount()).toBe(1);
		expect(first.project.visibleProcedureNames('Sheet1').has('alpha')).toBe(true);
		expect(first.byModule.get('module1')?.source).toContain('Alpha');
		expect(first.moduleMetadata.get('sheet1')?.moduleKind).toBe('document');
		expect(first.modules.map((mod) => mod.moduleName)).toEqual(['Module1', 'Sheet1']);
	});

	it('folds saved module changes incrementally without a project rebuild', async () => {
		const { index, projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
			{ name: 'Module2', type: 'standard', source: 'Public Sub Beta()\nEnd Sub\n' },
		]);
		const first = await projectIndexService.contextForProject(BOOK);

		index.updateModuleSource(BOOK, 'Module1', 'Public Sub Gamma()\nEnd Sub\n');
		const second = await projectIndexService.contextForProject(BOOK);

		expect(second.project).toBe(first.project);
		expect(callCount()).toBe(1);
		expect(second.project.visibleProcedureNames('Module2').has('gamma')).toBe(true);
		expect(second.project.visibleProcedureNames('Module2').has('alpha')).toBe(false);
		expect(second.byModule.get('module1')?.source).toContain('Gamma');
	});

	it('rebuilds after a project-level symbol index invalidation', async () => {
		const { index, projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const first = await projectIndexService.contextForProject(BOOK);

		index.invalidate(BOOK);
		const second = await projectIndexService.contextForProject(BOOK);

		expect(second.project).not.toBe(first.project);
		expect(callCount()).toBe(2);
	});

	it('applies open-document text per version on access', async () => {
		const { projectIndexService, callCount } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const document = openXlideDocument('Module1', 'Public Sub FromEditor()\nEnd Sub\n');
		(vscode.workspace.textDocuments as unknown[]).push(document);

		const first = await projectIndexService.contextForProject(BOOK);
		expect(first.byModule.get('module1')?.source).toContain('FromEditor');
		expect(first.project.visibleProcedureNames('Module1').has('fromeditor')).toBe(true);

		document.version = 2;
		document.getText = () => 'Public Sub Edited()\nEnd Sub\n';
		const second = await projectIndexService.contextForProject(BOOK);

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

		const live = await projectIndexService.contextForProject(BOOK, 'live');
		expect(live.byModule.has('bad')).toBe(false);
		expect(live.moduleMetadata.has('bad')).toBe(true);
		expect(live.project.visibleProcedureNames('Bad').has('alpha')).toBe(true);

		await expect(projectIndexService.contextForProject(BOOK, 'strict')).rejects.toBeTruthy();

		index.updateModuleSource(BOOK, 'Bad', 'Public Sub Fixed()\nEnd Sub\n');
		const strict = await projectIndexService.contextForProject(BOOK, 'strict');
		expect(strict.byModule.get('bad')?.source).toContain('Fixed');
		expect(strict.project.visibleProcedureNames('Module1').has('fixed')).toBe(true);
	});

	it('resets the consumer-memoized procedure signatures on module change', async () => {
		const { index, projectIndexService } = service([
			{ name: 'Module1', type: 'standard', source: 'Public Sub Alpha()\nEnd Sub\n' },
		]);
		const context = await projectIndexService.contextForProject(BOOK);
		context.projectProcedures = new Map();

		index.updateModuleSource(BOOK, 'Module1', 'Public Sub Beta()\nEnd Sub\n');
		const next = await projectIndexService.contextForProject(BOOK);

		expect(next.projectProcedures).toBeUndefined();
	});
});
