import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { WorkbookEngine } from '../src/workbookEngine';
import { synthesizeClassHeader } from '../src/vba/vbaProject';
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
			if (method === 'readModules') {
				return modules.map((mod) => ({ ...mod })) as T;
			}
			throw new Error(`Unexpected bridge call ${method}`);
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
			rightTitle: 'File: missing module',
		});
		expect(stale?.warning).toContain('stale .bas/.cls/.frm repo module file');
	});

	it('does not preview non-module or nested files as export deletions', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Stale.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'StaleClass.cls'), 'VERSION 1.0 CLASS\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'Notes.txt'), 'keep', 'utf8');
		// A .frx is the form's binary designer sidecar, not a module (#21).
		fs.writeFileSync(path.join(repo, 'UserForm1.frx'), 'keep', 'utf8');
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
			rightTitle: 'File: Sheet1 (will update)',
		});
		expect(byName.get('Sheet1')?.warning).toContain('code can be updated');
		expect(byName.get('Sheet2')).toMatchObject({
			status: 'skipping-import',
			checked: false,
			selectable: true,
			unsupportedDirectCreation: true,
			rightTitle: 'File: Sheet2 (cannot create)',
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
			'Sheet2: skipping import unless the module already exists in the file.',
			'UserForm2: skipping import unless the module already exists in the file.',
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
			rightTitle: 'File: StaleStandard (will delete)',
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
			rightTitle: 'File: NewModule (will create)',
		});
		expect(item.diff.filter((line) => line.left).every((line) => line.kind === 'added')).toBe(true);
	});

	it('plans .frm files as forms and ignores their .frx sidecars (#21)', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'UserForm1.frm'), 'VERSION 5.00\nBegin VB.UserForm UserForm1\nEnd\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'UserForm1.frx'), 'binary', 'utf8');
		fs.writeFileSync(path.join(repo, 'Module1.bas'), 'Sub T()\nEnd Sub\n', 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		expect(plan.items.map((item) => item.relativeName)).toEqual(['Module1.bas', 'UserForm1.frm']);
		const form = plan.items.find((item) => item.relativeName === 'UserForm1.frm');
		// A form still cannot be created from its file alone.
		expect(form?.moduleType).toBe('userform');
		expect(form?.status).toBe('skipping-import');
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

describe('module sync preview stays responsive on huge modules', () => {
	// A 26,000-line class shipped 130,000 DOM nodes into the preview on every
	// click - activating a row, and even ticking a checkbox on the row already
	// shown, cost well over a second each. The preview keeps only the visible
	// rows in the DOM and skips rebuilds that cannot change what is displayed;
	// these are the load-bearing pieces of that contract.
	const script = fs.readFileSync(
		path.join(__dirname, '..', 'assets', 'webview', 'moduleSync.js'), 'utf8',
	);
	const styles = fs.readFileSync(
		path.join(__dirname, '..', 'assets', 'webview', 'moduleSync.css'), 'utf8',
	);

	it('renders only the visible diff rows', () => {
		expect(script).toContain('function paintDiffWindow');
		// Repainting has to be driven by scrolling, or rows below the fold
		// would never appear.
		expect(script).toMatch(/el\('diff'\)\.addEventListener\('scroll', paintDiffWindow/);
		// Absolute placement on a full-height canvas is what keeps the
		// scrollbar honest while most rows are absent.
		expect(styles).toContain('.diffCanvas { position: relative; }');
		expect(styles).toContain('.diffCanvas > .line { position: absolute;');
	});

	it('skips diff rebuilds that cannot change what is shown', () => {
		expect(script).toContain('if (!force && key === diffRenderedKey) return;');
		// A refreshed plan can reuse an item id with new content, so that path
		// must force past the guard.
		expect(script).toMatch(/renderDiff\(true\);/);
	});

	it('updates list rows in place instead of rebuilding them', () => {
		expect(script).toContain('function syncListState');
		const rowClick = script.slice(script.indexOf("row.addEventListener('click'"));
		expect(rowClick.slice(0, 200)).toContain('syncListState();');
	});
});

describe('hiding headers hides the whole header, not just the attribute lines', () => {
	// Built from synthesizeClassHeader rather than a hand-written header: the
	// existing preview test used a fixture named for a class that carried no
	// class preamble, so the one case that failed was the one it did not cover.
	const BODY = 'Option Explicit\r\n\r\nPublic Sub Go()\r\nEnd Sub\r\n';

	it('drops the VERSION/BEGIN/END block a class header opens with', () => {
		const preview = editorPreviewSource(synthesizeClassHeader('Widget') + BODY);
		expect(preview.split('\n')[0]).toBe('Option Explicit');
		for (const leaked of ['VERSION 1.0 CLASS', 'BEGIN', 'MultiUse', 'END']) {
			expect(preview, leaked).not.toContain(leaked);
		}
		// The body is untouched.
		expect(preview).toContain('Public Sub Go()');
		expect(preview).toContain('End Sub');
	});

	it('drops a UserForm designer block including its nested controls', () => {
		const header = [
			'VERSION 5.00',
			'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ',
			'   Caption         =   "UserForm1"',
			'   Begin Forms.CommandButton CommandButton1 ',
			'      Caption         =   "OK"',
			'   End',
			'End',
			'Attribute VB_Name = "UserForm1"',
			'',
		].join('\r\n');
		const preview = editorPreviewSource(header + BODY);
		expect(preview.split('\n')[0]).toBe('Option Explicit');
		for (const leaked of ['VERSION 5.00', 'Begin ', 'CommandButton', 'Caption']) {
			expect(preview, leaked).not.toContain(leaked);
		}
	});

	it('leaves a standard module unchanged', () => {
		const preview = editorPreviewSource('Attribute VB_Name = "Module1"\r\n' + BODY);
		expect(preview.split('\n')[0]).toBe('Option Explicit');
		expect(preview).toContain('Public Sub Go()');
	});

	it('leaves source alone when a header block never closes', () => {
		// Showing too much beats hiding code.
		const truncated = 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\n' + BODY;
		expect(editorPreviewSource(truncated)).toContain('Public Sub Go()');
	});

	it('does not mistake a body End for the header block end', () => {
		const preview = editorPreviewSource(
			synthesizeClassHeader('Widget')
			+ 'Option Explicit\r\n\r\nPublic Sub Go()\r\n    With Me\r\n    End With\r\nEnd Sub\r\n',
		);
		expect(preview).toContain('End With');
		expect(preview).toContain('End Sub');
		expect(preview).not.toContain('VERSION');
	});
});

describe('a UserForm syncs as a .frm (#21)', () => {
	const FORM_SOURCE = [
		'Attribute VB_Name = "EntryForm"',
		'Attribute VB_Base = "0{11111111-0000-0000-0000-000000000000};{22222222-0000-0000-0000-000000000000}"',
		'Option Explicit',
		'',
		'Private Sub UserForm_Initialize()',
		'End Sub',
		'',
	].join('\r\n');

	it('exports a form under Name.frm, the name the VBE itself writes', async () => {
		const { workbook, repo } = tempWorkbook();
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: FORM_SOURCE }]);
		const plan = await buildExportModuleSyncPlan(bridge, { workbookPath: workbook, exportFolder: repo });
		const item = plan.items.find((candidate) => candidate.moduleName === 'EntryForm');
		expect(item?.relativeName).toBe('EntryForm.frm');
	});

	it('classifies a .frm as a userform by its name alone', async () => {
		const { workbook, repo } = tempWorkbook();
		// Deliberately headerless: a real Excel form's module text carries no
		// designer block, so the extension must be enough.
		fs.writeFileSync(path.join(repo, 'EntryForm.frm'), 'Option Explicit\r\n', 'utf8');
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: FORM_SOURCE }]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		const item = plan.items.find((candidate) => candidate.moduleName === 'EntryForm');
		expect(item?.moduleType).toBe('userform');
		expect(item?.status).toBe('will-update');
	});

	it('still cannot create a form from a repo file alone', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'NewForm.frm'), 'Option Explicit\r\n', 'utf8');
		const bridge = fakeBridge([]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		const item = plan.items.find((candidate) => candidate.moduleName === 'NewForm');
		expect(item?.status).toBe('skipping-import');
	});

	it('never treats a .frx sidecar as a module, in either direction', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'EntryForm.frm'), 'Option Explicit\r\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'EntryForm.frx'), Buffer.from([0x01, 0x02, 0x03]));
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: FORM_SOURCE }]);
		const importPlan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		expect(importPlan.items.some((item) => /\.frx$/i.test(item.relativeName))).toBe(false);
		// A trueUp export must not list the sidecar as a stale module either -
		// the .frx belongs to the .frm and is the VBE importer's to read.
		const exportPlan = await buildExportModuleSyncPlan(bridge, {
			workbookPath: workbook, exportFolder: repo, exportMode: 'trueUp',
		});
		expect(exportPlan.items.some((item) => /\.frx$/i.test(item.relativeName))).toBe(false);
	});

	it('retires an old .cls export of the same form on trueUp', async () => {
		const { workbook, repo } = tempWorkbook();
		// What 3.8.0 and earlier wrote: the form under a .cls name.
		fs.writeFileSync(path.join(repo, 'EntryForm.cls'), FORM_SOURCE, 'utf8');
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: FORM_SOURCE }]);
		const plan = await buildExportModuleSyncPlan(bridge, {
			workbookPath: workbook, exportFolder: repo, exportMode: 'trueUp',
		});
		const create = plan.items.find((item) => item.relativeName === 'EntryForm.frm');
		const stale = plan.items.find((item) => item.relativeName === 'EntryForm.cls');
		expect(create?.status).toBe('will-create');
		expect(stale?.status).toBe('will-remove');
	});

	it('a legacy .cls carrying a form header still classifies as a form', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'OldForm.cls'), FORM_SOURCE, 'utf8');
		const bridge = fakeBridge([{ name: 'OldForm', type: 'userform', source: FORM_SOURCE }]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		const item = plan.items.find((candidate) => candidate.moduleName === 'OldForm');
		expect(item?.moduleType).toBe('userform');
	});
});

describe('a form compares on the half its text can say (issue #36)', () => {
	// The .frm file carries the designer header the VBE's export writes; a
	// live module whose designer cannot be composed (readFormExport
	// unavailable or failing - the fakeBridge here, an engine without the
	// call, a form without designer storage) never has one. Raw equality was
	// false forever, so every form read as a perpetual "Will update".
	const CODE = [
		'Option Explicit',
		'',
		'Private Sub UserForm_Initialize()',
		'    Caption = "Entry"',
		'End Sub',
		'',
	].join('\r\n');
	const LIVE_FORM = [
		'Attribute VB_Name = "EntryForm"',
		'Attribute VB_Base = "0{11111111-0000-0000-0000-000000000000};{22222222-0000-0000-0000-000000000000}"',
		CODE,
	].join('\r\n');
	const REPO_FRM = [
		'VERSION 5.00',
		'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm ',
		'   Caption         =   "Entry"',
		'   ClientHeight    =   3000',
		'   OleObjectBlob   =   "EntryForm.frx":0000',
		'End',
		'Attribute VB_Name = "EntryForm"',
		'Attribute VB_Base = "0{11111111-0000-0000-0000-000000000000};{22222222-0000-0000-0000-000000000000}"',
		CODE,
	].join('\r\n');

	it('a clean round trip reads unchanged on the import plan', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'EntryForm.frm'), REPO_FRM, 'utf8');
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: LIVE_FORM }]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		const item = plan.items.find((candidate) => candidate.moduleName === 'EntryForm');
		expect(item?.status).toBe('unchanged');
		expect(item?.checked).toBe(false);
	});

	it('a code edit still reads will-update on the import plan', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'EntryForm.frm'),
			REPO_FRM + 'Public Sub Added()\r\nEnd Sub\r\n', 'utf8');
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: LIVE_FORM }]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		expect(plan.items.find((candidate) => candidate.moduleName === 'EntryForm')?.status)
			.toBe('will-update');
	});

	it('the export plan reads the same way', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'EntryForm.frm'), REPO_FRM, 'utf8');
		const bridge = fakeBridge([{ name: 'EntryForm', type: 'userform', source: LIVE_FORM }]);
		const clean = await buildExportModuleSyncPlan(bridge, { workbookPath: workbook, exportFolder: repo });
		expect(clean.items.find((candidate) => candidate.moduleName === 'EntryForm')?.status)
			.toBe('unchanged');

		const edited = fakeBridge([{
			name: 'EntryForm',
			type: 'userform',
			source: LIVE_FORM.replace('"Entry"', '"Changed"'),
		}]);
		const dirty = await buildExportModuleSyncPlan(edited, { workbookPath: workbook, exportFolder: repo });
		expect(dirty.items.find((candidate) => candidate.moduleName === 'EntryForm')?.status)
			.toBe('will-write');
	});

	it('non-form kinds keep the raw comparison, headers included', async () => {
		// A .cls attribute difference is a real pending change: class headers
		// round-trip, so raw equality stays the honest test there.
		const { workbook, repo } = tempWorkbook();
		const liveClass = [
			'Attribute VB_Name = "Person"',
			'Option Explicit',
			'',
		].join('\r\n');
		fs.writeFileSync(path.join(repo, 'Person.cls'), [
			'Attribute VB_Name = "Person"',
			'Attribute VB_Description = "A person."',
			'Option Explicit',
			'',
		].join('\r\n'), 'utf8');
		const bridge = fakeBridge([{ name: 'Person', type: 'class', source: liveClass }]);
		const plan = await buildImportModuleSyncPlan(bridge, { workbookPath: workbook, importFolder: repo });
		expect(plan.items.find((candidate) => candidate.moduleName === 'Person')?.status)
			.toBe('will-update');
	});
});
