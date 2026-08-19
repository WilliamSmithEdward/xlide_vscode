import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterAll, describe, expect, it } from 'vitest';
import {
	deleteModule,
	getProtectionInfo,
	getWorkbookInfo,
	listModules,
	listSubs,
	readCells,
	readModule,
	renameModule,
	validateWorkbook,
	writeCells,
	writeModule,
} from '../src/vba/workbookService';
import { openMacroContainer } from '../src/vba/macroContainer';

// Issues #24/#25, engine half: every macro container Office writes, read (and
// written where the write path is sound) natively. Each fixture here was
// authored by the live application itself through pyVBAharness, so the bytes
// are Office ground truth, not synthetic approximations.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const fixture = (name: string): string => path.join(FIXTURES, name);

const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-containers-'));
afterAll(() => {
	fs.rmSync(tempRoot, { recursive: true, force: true });
});

/** A scratch copy: mutations never touch the checked-in fixture. */
function copyOf(name: string): string {
	const target = path.join(tempRoot, `${Math.random().toString(36).slice(2)}-${name}`);
	fs.copyFileSync(fixture(name), target);
	return target;
}

interface ExpectedModule {
	name: string;
	type: string;
	documentType?: string;
	sourceContains?: string;
}

/** Fixture facts: the modules the authoring script created, per container. */
const READ_CASES: Array<{ file: string; modules: ExpectedModule[] }> = [
	{
		file: 'WordFixture.docm',
		modules: [
			{ name: 'ThisDocument', type: 'document', documentType: 'document', sourceContains: 'Private Sub Document_New()' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Public Sub GreetDocument()' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Selection.TypeText mGreeting' },
		],
	},
	{
		file: 'WordFixture.doc',
		modules: [
			{ name: 'ThisDocument', type: 'document', documentType: 'document', sourceContains: 'Private Sub Document_New()' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Public Sub GreetDocument()' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Selection.TypeText mGreeting' },
		],
	},
	{
		file: 'PowerPointFixture.pptm',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'ActivePresentation.Slides.Count' },
			{ name: 'CDeck', type: 'class', sourceContains: 'ActivePresentation.Name' },
		],
	},
	{
		file: 'PowerPointFixture.ppt',
		modules: [
			// The VBE case-normalized `Count` to `count` in this save; assert
			// the call chain, not the identifier casing it chose that day.
			{ name: 'Module1', type: 'standard', sourceContains: 'ActivePresentation.Slides.' },
			{ name: 'CDeck', type: 'class', sourceContains: 'ActivePresentation.Name' },
		],
	},
	{
		file: 'XlsFixture.xls',
		modules: [
			{ name: 'ThisWorkbook', type: 'document', documentType: 'workbook' },
			{ name: 'Sheet1', type: 'document', documentType: 'worksheet' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Hello from XlsFixture' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Debug.Print mGreeting' },
		],
	},
	{
		file: 'AccessFixture.accdb',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'DoCmd.OpenForm "frmGreeting"' },
			{ name: 'CAudit', type: 'class', sourceContains: 'Debug.Print action' },
			{ name: 'MBig', type: 'standard', sourceContains: 'step 800 of the long fixture module' },
		],
	},
	{
		file: 'AccessFixture.mdb',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'DoCmd.OpenForm "frmGreeting"' },
			{ name: 'CAudit', type: 'class', sourceContains: 'Debug.Print action' },
		],
	},
];

describe.each(READ_CASES)('reading $file', ({ file, modules }) => {
	it('lists the authored modules with their kinds', () => {
		const entries = listModules(fixture(file));
		for (const expected of modules) {
			const entry = entries.find((candidate) => candidate.name === expected.name);
			expect(entry, `${file} should carry ${expected.name}`).toBeDefined();
			expect(entry?.type, `${expected.name} type`).toBe(expected.type);
			if (expected.documentType) {
				expect(entry?.documentType, `${expected.name} documentType`).toBe(expected.documentType);
			}
		}
	});

	it('reads each module source', () => {
		for (const expected of modules) {
			const { source } = readModule(fixture(file), expected.name, true);
			expect(source).toContain(`Attribute VB_Name = "${expected.name}"`);
			if (expected.sourceContains) {
				expect(source).toContain(expected.sourceContains);
			}
		}
	});

	it('answers protection and structural validation', () => {
		const protection = getProtectionInfo(fixture(file));
		expect(protection.isPasswordProtected).toBe(false);
		expect(validateWorkbook(fixture(file)).issues).toEqual([]);
	});
});

describe('the containers behave as their formats require', () => {
	it('finds procedures inside an Access module whose stream chains across LVAL pages', () => {
		const subs = listSubs(fixture('AccessFixture.accdb'), 'MBig');
		expect(subs.map((sub) => sub.name)).toContain('BigTally');
		const { source } = readModule(fixture('AccessFixture.accdb'), 'MBig', true);
		// 800 generated lines force a multi-chunk stream held in a chained
		// long value; every line surviving proves the chain walk.
		for (const step of [1, 400, 800]) {
			expect(source).toContain(`step ${step} of the long fixture module`);
		}
	});

	it('getWorkbookInfo answers modules without a sheet surface for a Word document', () => {
		const info = getWorkbookInfo(fixture('WordFixture.docm'));
		expect(info.sheets).toEqual([]);
		expect(info.namedRanges).toEqual([]);
		expect(info.modules.map((module) => module.name)).toContain('ThisDocument');
	});

	it('refuses sheet reads on containers without worksheets, by name', () => {
		expect(() => readCells(fixture('WordFixture.docm'), 'Sheet1', 'A1'))
			.toThrow(/Word macro-enabled document.*no worksheet surface/);
		expect(() => readCells(fixture('AccessFixture.accdb'), 'Sheet1', 'A1'))
			.toThrow(/Access database.*no worksheet surface/);
		expect(() => readCells(fixture('XlsFixture.xls'), 'Sheet1', 'A1'))
			.toThrow(/legacy Excel workbook.*no worksheet surface/);
	});

	it('classifies containers from content, not extension', () => {
		const disguised = path.join(tempRoot, 'disguised.xlsm');
		fs.copyFileSync(fixture('WordFixture.docm'), disguised);
		const container = openMacroContainer(fs.readFileSync(disguised));
		expect(container.kind).toBe('word');
	});
});

const WRITE_CASES = [
	'WordFixture.docm',
	'PowerPointFixture.pptm',
	'WordFixture.doc',
	'XlsFixture.xls',
	'PowerPointFixture.ppt',
];

describe.each(WRITE_CASES.map((file) => ({ file })))('writing $file', ({ file }) => {
	it('adds a module, preserves the rest, and round-trips through a fresh parse', () => {
		const target = copyOf(file);
		const before = listModules(target).map((module) => module.name).sort();
		const beforeSources = new Map(
			before.map((name) => [name, readModule(target, name, true).source]),
		);

		writeModule(target, 'MAdded', 'Public Sub AddedProc()\r\n    Debug.Print "added"\r\nEnd Sub\r\n');

		const after = listModules(target);
		expect(after.map((module) => module.name)).toContain('MAdded');
		const added = readModule(target, 'MAdded', true).source;
		expect(added).toContain('Attribute VB_Name = "MAdded"');
		expect(added).toContain('Public Sub AddedProc()');
		expect(listSubs(target, 'MAdded').map((sub) => sub.name)).toEqual(['AddedProc']);
		for (const name of before) {
			expect(readModule(target, name, true).source, `${name} must survive the write`)
				.toBe(beforeSources.get(name));
		}
		expect(validateWorkbook(target).issues).toEqual([]);
	});

	it('edits an existing module in place', () => {
		const target = copyOf(file);
		writeModule(target, 'Module1', 'Public Sub Rewritten()\r\nEnd Sub\r\n');
		const { source } = readModule(target, 'Module1', true);
		expect(source).toContain('Public Sub Rewritten()');
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});

describe('rename and delete on a non-Excel host', () => {
	it('round-trips on a Word document', () => {
		const target = copyOf('WordFixture.docm');
		renameModule(target, 'CGreeter', 'CRenamed');
		expect(listModules(target).map((module) => module.name)).toContain('CRenamed');
		expect(readModule(target, 'CRenamed', true).source).toContain('Attribute VB_Name = "CRenamed"');
		deleteModule(target, 'CRenamed');
		expect(listModules(target).map((module) => module.name)).not.toContain('CRenamed');
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});

describe('Access stays read-only, with the reason stated', () => {
	it.each([['AccessFixture.accdb'], ['AccessFixture.mdb']])('%s refuses writes', (file) => {
		const target = copyOf(file);
		expect(() => writeModule(target, 'MNew', 'Public Sub P()\r\nEnd Sub\r\n'))
			.toThrow(/read-only: Access runs VBA from its compiled p-code/);
		expect(fs.readFileSync(target).equals(fs.readFileSync(fixture(file)))).toBe(true);
	});

	it('cell writes name the actual container', () => {
		const target = copyOf('WordFixture.docm');
		expect(() => writeCells(target, 'Sheet1', 'A1', [['x']]))
			.toThrow(/Word macro-enabled document.*cell writes need an OOXML Excel workbook/);
	});
});

// ---------------------------------------------------------------------------
// The export / import planning surface over real non-Excel containers: the
// same plan builders the GUI commands call, driven by the real engine service
// through a minimal bridge shim (the plan layer is bridge-driven and never
// looks at the container itself).

import type { WorkbookEngine } from '../src/workbookEngine';
import { buildExportModuleSyncPlan, buildImportModuleSyncPlan } from '../src/moduleSyncPlan';
import * as svc from '../src/vba/workbookService';

function realServiceBridge(): WorkbookEngine {
	return {
		async call<T>(method: string, args: Record<string, unknown>): Promise<T> {
			const p = String(args.path);
			switch (method) {
				case 'listModules': return svc.listModules(p) as T;
				case 'readModules': return svc.readModules(p, Boolean(args.full)) as T;
				case 'readModule': return svc.readModule(p, String(args.module), Boolean(args.full)) as T;
				case 'readFormExport': return svc.readFormExport(p, String(args.module)) as T;
				case 'writeModule': return svc.writeModule(p, String(args.module), String(args.source ?? ''), args.kind === 'class' ? 'class' : 'standard') as T;
				default: throw new Error(`bridge shim: unhandled ${method}`);
			}
		},
	} as WorkbookEngine;
}

describe('the module sync plans cover the non-Excel containers', () => {
	it('plans a Word document export with every module surfaced', async () => {
		const target = copyOf('WordFixture.docm');
		const exportFolder = path.join(tempRoot, 'word-export');
		const plan = await buildExportModuleSyncPlan(realServiceBridge(), {
			workbookPath: target,
			exportFolder,
		});
		const names = plan.items.map((item) => item.moduleName).sort();
		expect(names).toEqual(expect.arrayContaining(['CGreeter', 'Module1', 'ThisDocument']));
		expect(plan.items.every((item) => item.status !== undefined)).toBe(true);
	});

	it('plans an Access database export (read-only container, read-only operation)', async () => {
		const exportFolder = path.join(tempRoot, 'access-export');
		const plan = await buildExportModuleSyncPlan(realServiceBridge(), {
			workbookPath: fixture('AccessFixture.accdb'),
			exportFolder,
		});
		const names = plan.items.map((item) => item.moduleName).sort();
		expect(names).toEqual(expect.arrayContaining(['CAudit', 'MBig', 'Module1']));
	});

	it('plans and understands an import back into a Word document', async () => {
		const target = copyOf('WordFixture.docm');
		const importFolder = path.join(tempRoot, 'word-import');
		fs.mkdirSync(importFolder, { recursive: true });
		fs.writeFileSync(
			path.join(importFolder, 'Module1.bas'),
			'Attribute VB_Name = "Module1"\r\nPublic Sub Reworked()\r\nEnd Sub\r\n',
		);
		const plan = await buildImportModuleSyncPlan(realServiceBridge(), {
			workbookPath: target,
			importFolder,
		});
		const item = plan.items.find((candidate) => candidate.moduleName === 'Module1');
		expect(item).toBeDefined();
		expect(item?.status).toBe('will-update');
	});
});

// ---------------------------------------------------------------------------
// The unit-test staging pipeline over a non-Excel container: the same staging
// the Run VBA Tests command performs, against a real Word fixture.

import { discoverVbaTestsFromModule } from '../src/vbaTestRunner';
import { stageOwnedReadOnlyExcelTestHost } from '../src/vbaTestHostStaging';

describe('test staging covers Word containers', () => {
	it('stages assert, runner and dispatcher modules into a .docm copy', async () => {
		const target = copyOf('WordFixture.docm');
		const testSource = [
			"' @xlide-test",
			'Public Sub ChecksArithmetic()',
			'    XlideAssert.AreEqual 4, 2 + 2',
			'End Sub',
			'',
			'Public Sub HelperTarget()',
			'End Sub',
			'',
		].join('\r\n');
		svc.writeModule(target, 'ZzTests', testSource);
		const tests = discoverVbaTestsFromModule({
			name: 'ZzTests',
			type: 'standard',
			source: svc.readModule(target, 'ZzTests', true).source,
		});
		expect(tests.map((test) => test.qualifiedName)).toEqual(['ZzTests.ChecksArithmetic']);

		const staging = await stageOwnedReadOnlyExcelTestHost(
			realServiceBridge(), target, tests, { hostApp: 'word', log: () => undefined },
		);
		try {
			const staged = svc.listModules(staging.tempWorkbookPath).map((module) => module.name);
			expect(staged).toEqual(expect.arrayContaining(['XlideAssert', 'ZzTests', 'XlideTestDispatch']));
			expect(staged.some((name) => name.startsWith('XlideRun'))).toBe(true);
			const dispatcher = svc.readModule(staging.tempWorkbookPath, 'XlideTestDispatch', true).source;
			expect(dispatcher).toContain('ZzTests.HelperTarget');
			expect(dispatcher).toContain('XlideAssert.RecordTargetOutcome');
			const script = fs.readFileSync(staging.hostScriptPath, 'utf8');
			expect(script).toContain("$hostKind = 'word'");
			expect(script).toContain('$excel.Documents.Open($targetPath, $false, $true, $false)');
			expect(validateWorkbook(staging.tempWorkbookPath).issues).toEqual([]);
		} finally {
			staging.dispose();
		}
	});
});

// ---------------------------------------------------------------------------
// UserForms inside a WORD container: the designer storage machinery is
// container-agnostic, and this fixture (a form authored by Word's own VBE,
// with two controls and code-behind) proves it outside Excel.

import { readFormExport, writeFormDesigner } from '../src/vba/workbookService';
import { parseFormFrx } from '../src/vba/formDesigner';

describe('a Word document with a UserForm', () => {
	it('classifies the form and reads its designer-declared controls', () => {
		const entries = listModules(fixture('WordFormFixture.docm'));
		const form = entries.find((entry) => entry.name === 'FrmNotice');
		expect(form?.type).toBe('userform');
		const controls = (form?.implicitMembers ?? []).map((member) => member.name);
		expect(controls).toEqual(expect.arrayContaining(['LblMessage', 'BtnClose']));
		const { source } = readModule(fixture('WordFormFixture.docm'), 'FrmNotice', true);
		expect(source).toContain('Private Sub BtnClose_Click()');
	});

	it('exports the .frm/.frx pair and writes the designer back', () => {
		const target = copyOf('WordFormFixture.docm');
		const { frm, frx } = readFormExport(target, 'FrmNotice');
		expect(frm).toContain('Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmNotice');
		expect(frm).toContain('OleObjectBlob');
		expect(frm).toContain('"FrmNotice.frx":0000');
		const streams = parseFormFrx(frx);
		expect(streams).toBeDefined();

		writeFormDesigner(target, 'FrmNotice', frx, undefined);
		const after = listModules(target).find((entry) => entry.name === 'FrmNotice');
		expect(after?.type).toBe('userform');
		expect((after?.implicitMembers ?? []).map((member) => member.name))
			.toEqual(expect.arrayContaining(['LblMessage', 'BtnClose']));
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});

// ---------------------------------------------------------------------------
// Malformed input stays graceful: honest errors, never hangs or crashes.

import { accessVbaCfb, isAccessDatabase } from '../src/vba/accessDatabase';
import { MacroContainerError } from '../src/vba/macroContainer';

describe('malformed containers fail honestly', () => {
	it('rejects non-container bytes with a clear message', () => {
		for (const bytes of [Buffer.alloc(0), Buffer.from('hello world'), Buffer.alloc(4096)]) {
			expect(() => openMacroContainer(bytes)).toThrow(MacroContainerError);
			expect(() => openMacroContainer(bytes)).toThrow(/Not a macro-enabled Office file/);
		}
	});

	it('rejects a CFB with no recognizable Office host', () => {
		// A bare vbaProject.bin IS a CFB, but not an Office document.
		const bin = fs.readFileSync(fixture('FormFixture.xlsm'));
		const container = openMacroContainer(bin);
		const vbaBin = container.vbaCfb().toBytes();
		expect(() => openMacroContainer(vbaBin)).toThrow(/Compound file without a recognizable Office host/);
	});

	it('survives a corrupted Access database without hanging', () => {
		const data = Buffer.from(fs.readFileSync(fixture('AccessFixture.accdb')));
		expect(isAccessDatabase(data)).toBe(true);
		// Scramble every LVAL page's slot table: chains and rows become
		// nonsense, and the reader must fail (or answer empty) rather than
		// loop or crash.
		for (let page = 1; page < data.length / 4096; page++) {
			const base = page * 4096;
			if (data[base] === 0x01 && data.subarray(base + 4, base + 8).toString('latin1') === 'LVAL') {
				data.writeUInt16LE(0x0fff, base + 12);
				for (let i = 0; i < 64; i++) {
					data.writeUInt16LE((i * 7919) & 0xffff, base + 14 + 2 * i);
				}
			}
		}
		try {
			const cfb = accessVbaCfb(data);
			// A parse that survives must still answer as a project (possibly
			// module-less), not garbage.
			expect(cfb.hasStream('dir')).toBe(true);
		} catch (err) {
			expect(String(err)).toMatch(/No VBA project catalog|LVAL|Access/);
		}
	});

	it('truncated databases are not Access databases', () => {
		const data = fs.readFileSync(fixture('AccessFixture.accdb'));
		expect(isAccessDatabase(data.subarray(0, 4000))).toBe(false);
		expect(isAccessDatabase(data.subarray(0, 8191))).toBe(false);
	});
});

describe('repeated saves keep the .ppt persist machinery consistent', () => {
	it('survives three consecutive writes with the project intact', () => {
		const target = copyOf('PowerPointFixture.ppt');
		writeModule(target, 'MFirst', 'Public Sub One()\r\nEnd Sub\r\n');
		writeModule(target, 'MSecond', 'Public Sub Two()\r\nEnd Sub\r\n');
		writeModule(target, 'MFirst', 'Public Sub OneRewritten()\r\nEnd Sub\r\n');
		const names = listModules(target).map((module) => module.name);
		expect(names).toEqual(expect.arrayContaining(['Module1', 'CDeck', 'MFirst', 'MSecond']));
		expect(readModule(target, 'MFirst', true).source).toContain('OneRewritten');
		expect(readModule(target, 'Module1', true).source).toContain('ActivePresentation.Slides.');
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});
