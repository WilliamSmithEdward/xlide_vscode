import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { WorkbookEngine } from '../src/workbookEngine';
import { WorkbookEngineError, JSONRPC_METHOD_NOT_FOUND } from '../src/workbookEngineErrors';
import {
	buildExportModuleSyncPlan,
	buildImportModuleSyncPlan,
	buildSideBySideDiff,
	classifyModuleType,
	editorPreviewSource,
} from '../src/moduleSyncPlan';

interface FakeModule {
	name: string;
	type: string;
	source: string;
	documentType?: string;
}

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string; repo: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-sync-plan-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	const repo = path.join(root, 'repo');
	fs.writeFileSync(workbook, '', 'utf8');
	fs.mkdirSync(repo, { recursive: true });
	return { root, workbook, repo };
}

function fakeBridge(modules: readonly FakeModule[]): WorkbookEngine {
	return {
		async call<T>(method: string, args: Record<string, unknown>): Promise<T> {
			if (method === 'listModules') {
				return modules.map((mod) => ({
					name: mod.name,
					type: mod.type,
					documentType: mod.documentType,
				})) as T;
			}
			if (method === 'readModule') {
				const moduleName = String(args.module ?? '').toLowerCase();
				const mod = modules.find((candidate) => candidate.name.toLowerCase() === moduleName);
				if (!mod) {
					throw new Error(`Unknown module ${String(args.module)}`);
				}
				return { source: mod.source } as T;
			}
			throw new WorkbookEngineError(`Method not found: ${method}`, JSONRPC_METHOD_NOT_FOUND);
		},
	} as WorkbookEngine;
}

function batchFakeBridge(modules: readonly FakeModule[], calls: string[]): WorkbookEngine {
	return {
		async call<T>(method: string): Promise<T> {
			calls.push(method);
			if (method === 'readModules') {
				return modules.map((mod) => ({
					name: mod.name,
					type: mod.type,
					documentType: mod.documentType,
					source: mod.source,
				})) as T;
			}
			throw new Error(`Unexpected bridge call ${method}`);
		},
	} as WorkbookEngine;
}

// Pins classifyModuleType (src/vba/workbookService.ts) row by row. This table
// began life shared with the retired Python backend's classifier; it survives
// because the rows encode VBE behavior worth keeping pinned, not because a
// second implementation still mirrors it.
const sharedClassificationTable: ReadonlyArray<readonly [string, string, string]> = [
	['Module1', 'Option Explicit\nSub Hello()\nEnd Sub\n', 'standard'],
	['Globals', 'Attribute VB_PredeclaredId = True\n', 'standard'],
	['stdArray', 'Attribute VB_PredeclaredId = True\nOption Explicit\n', 'standard'],
	['MyClass', 'Attribute VB_Base = "{CC27B1A4-1234-1234-1234-000000000000}"\n', 'standard'],
	['UserForm1', 'Attribute VB_Base = "0{11111111-0000-0000-0000-000000000000};{22222222-0000-0000-0000-000000000000}"\n', 'userform'],
	['UserForm2', 'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046};{22222222-0000-0000-0000-000000000000}"\n', 'userform'],
	['ThisWorkbook', 'Attribute VB_Base = "{00020819-0000-0000-C000-000000000046}"\n', 'document'],
	['Sheet1', 'Attribute VB_Base = "{00020820-0000-0000-C000-000000000046}"\n', 'document'],
	['Chart1', 'Attribute VB_Base = "{00020821-0000-0000-C000-000000000046}"\n', 'document'],
	['CustomDoc', 'Attribute VB_Base = "{00020820-0000-0000-c000-000000000046}"\n', 'document'],
	['ThisWorkbook', 'Option Explicit\n', 'document'],
	['Sheet3', 'Option Explicit\n', 'document'],
	['Feuil2', 'Option Explicit\n', 'document'],
	['thisworkbook', 'Option Explicit\n', 'standard'],
];

describe('module sync plan', () => {
	it('classifies module types per the pinned classification table', () => {
		for (const [name, source, expected] of sharedClassificationTable) {
			expect(classifyModuleType(name, source), `${name}: ${JSON.stringify(source)}`).toBe(expected);
		}
	});

	it('imports a predeclared class module as a creatable class, not a skipped document', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'stdArray.cls'), [
			'VERSION 1.0 CLASS',
			'BEGIN',
			'  MultiUse = -1  \'True',
			'END',
			'Attribute VB_Name = "stdArray"',
			'Attribute VB_PredeclaredId = True',
			'Attribute VB_Exposed = False',
			'Option Explicit',
			'',
		].join('\r\n'), 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		expect(plan.items[0]).toMatchObject({
			moduleName: 'stdArray',
			moduleType: 'class',
			status: 'will-create',
			checked: true,
			unsupportedDirectCreation: false,
		});
		expect(plan.warnings).toEqual([]);
	});

	it('plans export row statuses from one workbook-vs-repo comparison path', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Existing.bas'), 'Sub Same()\nEnd Sub\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'Changed.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');

		const plan = await buildExportModuleSyncPlan(fakeBridge([
			{ name: 'Existing', type: 'standard', source: 'Sub Same()\nEnd Sub\n' },
			{ name: 'Changed', type: 'standard', source: 'Sub Newer()\nEnd Sub\n' },
			{ name: 'NewModule', type: 'standard', source: 'Sub NewModule()\nEnd Sub\n' },
		]), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'exportAll',
		});

		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('Existing')).toMatchObject({
			status: 'unchanged',
			checked: false,
			selectable: true,
		});
		expect(byName.get('Changed')).toMatchObject({
			status: 'will-write',
			checked: true,
			selectable: true,
			rightTitle: 'Repo: Changed.bas (will overwrite)',
		});
		expect(byName.get('NewModule')).toMatchObject({
			status: 'will-create',
			checked: true,
			selectable: true,
			rightTitle: 'Repo: NewModule.bas (will create)',
		});
		expect(byName.get('NewModule')?.warning).toBeUndefined();
		expect(byName.get('NewModule')?.diff.filter((line) => line.left).every((line) => line.kind === 'added')).toBe(true);
	});

	it('builds the export plan from a single batch readModules call when available', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Existing.bas'), 'Sub Same()\nEnd Sub\n', 'utf8');
		const calls: string[] = [];

		const plan = await buildExportModuleSyncPlan(batchFakeBridge([
			{ name: 'Existing', type: 'standard', source: 'Sub Same()\nEnd Sub\n' },
			{ name: 'Changed', type: 'standard', source: 'Sub Newer()\nEnd Sub\n' },
		], calls), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'exportAll',
		});

		expect(calls).toEqual(['readModules']);
		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('Existing')?.status).toBe('unchanged');
		expect(byName.get('Changed')?.status).toBe('will-create');
	});

	it('surfaces true-up stale root module files as removable preview rows', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Stale.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');

		const plan = await buildExportModuleSyncPlan(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'trueUp',
		});

		const stale = plan.items.find((item) => item.relativeName === 'Stale.bas');
		expect(stale).toMatchObject({
			moduleName: 'Stale',
			status: 'will-remove',
			checked: true,
			selectable: true,
			existsInWorkbook: false,
			existsInRepo: true,
			leftTitle: 'Repo: Stale.bas (will remove)',
			rightTitle: 'Workbook: missing module',
		});
		expect(stale?.warning).toContain('stale .bas/.cls repo module file');
	});

	it('does not preview non-module or nested files as export deletions', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Stale.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'StaleClass.cls'), 'VERSION 1.0 CLASS\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'Notes.txt'), 'keep', 'utf8');
		fs.writeFileSync(path.join(repo, 'UserForm1.frm'), 'keep', 'utf8');
		fs.mkdirSync(path.join(repo, 'nested'));
		fs.writeFileSync(path.join(repo, 'nested', 'Stale.cls'), 'keep', 'utf8');

		const plan = await buildExportModuleSyncPlan(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'trueUp',
		});

		const staleRows = plan.items.filter((item) => item.status === 'will-remove').map((item) => item.relativeName);
		expect(staleRows).toEqual(['Stale.bas', 'StaleClass.cls']);
	});

	it('warns and skips missing document/userform cls code-behind modules while allowing existing name-match updates', async () => {
		const { workbook, repo } = tempWorkbook();
		const sheetBase = 'Attribute VB_Base = "{00020820-0000-0000-C000-000000000046}"\n';
		const formBase = 'Attribute VB_Base = "{00000000-0000-0000-0000-000000000001}{00000000-0000-0000-0000-000000000002}"\n';
		fs.writeFileSync(path.join(repo, 'Sheet1.cls'), `${sheetBase}Private Sub Worksheet_Change()\nEnd Sub\n`, 'utf8');
		fs.writeFileSync(path.join(repo, 'Sheet2.cls'), `${sheetBase}Private Sub Worksheet_Activate()\nEnd Sub\n`, 'utf8');
		fs.writeFileSync(path.join(repo, 'UserForm1.cls'), `${formBase}Private Sub CommandButton1_Click()\nEnd Sub\n`, 'utf8');
		fs.writeFileSync(path.join(repo, 'UserForm2.cls'), `${formBase}Private Sub CommandButton1_Click()\nEnd Sub\n`, 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([
			{
				name: 'Sheet1',
				type: 'document',
				documentType: 'worksheet',
				source: `${sheetBase}Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n`,
			},
			{
				name: 'UserForm1',
				type: 'userform',
				source: `${formBase}Private Sub UserForm_Initialize()\nEnd Sub\n`,
			},
		]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('Sheet1')).toMatchObject({
			status: 'will-update',
			checked: true,
			selectable: true,
			unsupportedDirectCreation: false,
			rightTitle: 'Workbook: Sheet1 (will update)',
		});
		expect(byName.get('Sheet1')?.warning).toContain('code can be updated');
		expect(byName.get('Sheet2')).toMatchObject({
			status: 'skipping-import',
			checked: false,
			selectable: true,
			unsupportedDirectCreation: true,
			rightTitle: 'Workbook: Sheet2 (cannot create)',
		});
		expect(byName.get('Sheet2')?.warning).toContain('cannot be created directly');
		expect(byName.get('UserForm1')).toMatchObject({
			status: 'will-update',
			checked: true,
			selectable: true,
			unsupportedDirectCreation: false,
		});
		expect(byName.get('UserForm1')?.warning).toContain('code can be updated');
		expect(byName.get('UserForm2')).toMatchObject({
			status: 'skipping-import',
			checked: false,
			selectable: true,
			unsupportedDirectCreation: true,
		});
		expect(byName.get('UserForm2')?.warning).toContain('cannot be created directly');
		expect(plan.warnings).toEqual([
			'Sheet2: skipping import unless the module already exists in the workbook.',
			'UserForm2: skipping import unless the module already exists in the workbook.',
		]);
	});

	it('plans import true-up deletions for workbook-only standard and class modules only', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Module1.bas'), 'Sub T()\nEnd Sub\n', 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
			{ name: 'StaleStandard', type: 'standard', source: 'Sub Old()\nEnd Sub\n' },
			{ name: 'StaleClass', type: 'class', source: 'VERSION 1.0 CLASS\n' },
			{
				name: 'Sheet1',
				type: 'document',
				documentType: 'worksheet',
				source: 'Private Sub Worksheet_Activate()\nEnd Sub\n',
			},
			{ name: 'UserForm1', type: 'userform', source: 'VERSION 5.00\n' },
		]), {
			workbookPath: workbook,
			importFolder: repo,
			importMode: 'trueUpStandardClass',
		});

		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('StaleStandard')).toMatchObject({
			status: 'will-remove',
			checked: true,
			selectable: true,
			existsInWorkbook: true,
			existsInRepo: false,
			detail: 'Will delete workbook module',
			rightTitle: 'Workbook: StaleStandard (will delete)',
		});
		expect(byName.get('StaleClass')).toMatchObject({
			status: 'will-remove',
			checked: true,
			selectable: true,
		});
		expect(byName.get('StaleClass')?.diff.filter((line) => line.right).every((line) => line.kind === 'removed')).toBe(true);
		expect(byName.has('Sheet1')).toBe(false);
		expect(byName.has('UserForm1')).toBe(false);
	});

	it('tones import-created source lines as additions', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'NewModule.bas'), 'Sub T()\nEnd Sub\n', 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		const item = plan.items[0];
		expect(item).toMatchObject({
			moduleName: 'NewModule',
			status: 'will-create',
			rightTitle: 'Workbook: NewModule (will create)',
		});
		expect(item.diff.filter((line) => line.left).every((line) => line.kind === 'added')).toBe(true);
	});

	it('ignores frm files in import planning', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'UserForm1.frm'), 'VERSION 5.00\nBegin VB.UserForm UserForm1\nEnd\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'Module1.bas'), 'Sub T()\nEnd Sub\n', 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		expect(plan.items.map((item) => item.relativeName)).toEqual(['Module1.bas']);
	});

	it('builds a side-by-side diff with changed line metadata', () => {
		const diff = buildSideBySideDiff('A\nB\nC', 'A\nX\nC');
		expect(diff.map((line) => line.kind)).toContain('changed');
		expect(diff.find((line) => line.kind === 'changed')).toMatchObject({
			left: 'B',
			right: 'X',
		});
	});

	it('builds preview source without VBA attribute headers', () => {
		expect(editorPreviewSource([
			'Attribute VB_Name = "Class1"',
			'Attribute VB_GlobalNameSpace = False',
			'',
			'Option Explicit',
			'Public Property Get Name() As String',
			'End Property',
			'Attribute Name.VB_Description = "Hidden"',
		].join('\n'))).toBe([
			'Option Explicit',
			'Public Property Get Name() As String',
			'End Property',
		].join('\n'));
	});

	it('keeps raw source available for optional header diff display', async () => {
		const { workbook, repo } = tempWorkbook();
		const source = [
			'Attribute VB_Name = "Module1"',
			'',
			'Option Explicit',
			'Sub T()',
			'End Sub',
		].join('\n');
		fs.writeFileSync(path.join(repo, 'Module1.bas'), source, 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([]), {
			workbookPath: workbook,
			importFolder: repo,
		});
		const item = plan.items[0];

		expect(item.leftCode).toBe([
			'Option Explicit',
			'Sub T()',
			'End Sub',
		].join('\n'));
		expect(item.leftRawCode).toBe(source);
		expect(item.diff.map((line) => line.left).join('\n')).not.toContain('Attribute VB_Name');
		expect(item.diffWithHeaders.map((line) => line.left).join('\n')).toContain('Attribute VB_Name');
	});
});
