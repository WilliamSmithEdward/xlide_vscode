import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { parseFormDesignerStreams, parseFormFrx, splitFrmSource, type DesignerControl } from '../src/vba/formDesigner';
import {
	listModules,
	readFormExport,
	readModule,
	resetWorkbookCacheForTests,
	writeFormDesigner,
	writeModule,
} from '../src/vba/workbookService';
import { readModules as readModulesFull } from '../src/vba/workbookService';
import { Cfb } from '../src/vba/cfb';
import { buildExportModuleSyncPlan, buildImportModuleSyncPlan } from '../src/moduleSyncPlan';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectEditorSymbolContextForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { resolveMemberCompletions } from '../src/analyzer';

// A real Excel-authored workbook: a four-control UserForm built through the
// VBE and saved by Excel itself. The designer reader's claims are pinned to
// what Excel actually wrote, not to bytes this repo synthesized.
const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixture.xlsm');

const EXPECTED: DesignerControl[] = [
	{ name: 'RegionPick', type: 'MSForms.ComboBox' },
	{ name: 'Taxable', type: 'MSForms.CheckBox' },
	{ name: 'NameBox', type: 'MSForms.TextBox' },
	{ name: 'OkButton', type: 'MSForms.CommandButton' },
];

const ascii = (bytes: Buffer): string => bytes.toString('latin1');

const tempDirs: string[] = [];
afterEach(() => {
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function fixtureStreams(): { f: Buffer; o: Buffer } {
	const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(FIXTURE)).readVbaProject());
	return {
		f: cfb.getStreamInStorage('FrmPicker', 'f'),
		o: cfb.getStreamInStorage('FrmPicker', 'o'),
	};
}

describe('the designer storage of a real Excel form', () => {
	it('yields every control with its name and type', () => {
		const { f, o } = fixtureStreams();
		expect(parseFormDesignerStreams(f, o, ascii)).toEqual(EXPECTED);
	});

	it('needs no o stream when the cache indices are specific', () => {
		// Excel persists 25/26/23/17 directly rather than generic MorphData,
		// so typing this form never opens the o stream.
		const { f } = fixtureStreams();
		expect(parseFormDesignerStreams(f, undefined, ascii)).toEqual(EXPECTED);
	});

	it('returns undefined for a truncated stream rather than guessing', () => {
		const { f } = fixtureStreams();
		for (const cut of [1, 3, 40, 80, f.length - 1]) {
			expect(parseFormDesignerStreams(f.subarray(0, cut), undefined, ascii), `cut at ${cut}`)
				.toBeUndefined();
		}
	});

	it('returns undefined for a stream that is not a FormControl', () => {
		expect(parseFormDesignerStreams(Buffer.from('VERSION 5.00\r\n'), undefined, ascii))
			.toBeUndefined();
	});

	it('types a generic MorphData site through DisplayStyle in the o stream', () => {
		// Rewrite the real form's first site (ComboBox, index 25) as generic
		// MorphData (15) and hand it an o stream whose first record stores
		// DisplayStyle = combo. Same answer, different route.
		const { f } = fixtureStreams();
		const patched = Buffer.from(f);
		// Site 1's ClsidCacheIndex sits at 0x48 (walked by hand against
		// [MS-OFORMS] before this reader was written).
		expect(patched.readUInt16LE(0x48)).toBe(25);
		patched.writeUInt16LE(15, 0x48);
		const morphO = Buffer.alloc(64);
		morphO.writeUInt8(0x00, 0); // MinorVersion
		morphO.writeUInt8(0x02, 1); // MajorVersion
		morphO.writeUInt16LE(12, 2); // cbMorphData
		morphO.writeUInt32LE(1 << 6, 4); // PropMask low: fDisplayStyle only
		morphO.writeUInt32LE(0, 8); // PropMask high
		morphO.writeUInt8(0x03, 12); // fmDisplayStyleCombo
		const controls = parseFormDesignerStreams(patched, morphO, ascii);
		expect(controls?.[0]).toEqual({ name: 'RegionPick', type: 'MSForms.ComboBox' });
	});

	it('falls back to the base control surface for a MorphData site with no o stream', () => {
		const { f } = fixtureStreams();
		const patched = Buffer.from(f);
		patched.writeUInt16LE(15, 0x48);
		const controls = parseFormDesignerStreams(patched, undefined, ascii);
		expect(controls?.[0]).toEqual({ name: 'RegionPick', type: 'MSForms.Control' });
	});
});

describe('a workbook form knows its controls with no host at all', () => {
	// The end of the story #17 started: Excel keeps a form's control tree in
	// a binary designer blob, so a workbook-backed form's controls were
	// unknowable without a host. The engine reads them natively now.
	it('listModules carries the controls on the form entry', () => {
		resetWorkbookCacheForTests();
		const entries = listModules(FIXTURE);
		const form = entries.find((entry) => entry.name === 'FrmPicker');
		expect(form?.type).toBe('userform');
		expect(form?.implicitMembers).toEqual(EXPECTED);
		// And only on the form: no other entry invents members.
		for (const entry of entries.filter((candidate) => candidate.name !== 'FrmPicker')) {
			expect(entry.implicitMembers, entry.name).toBeUndefined();
		}
	});

	it('the code-behind analyzes clean and a caller resolves the controls', () => {
		const entries = listModules(FIXTURE).map((entry) => ({
			moduleName: entry.name,
			type: entry.type,
			documentType: entry.documentType,
			source: readModule(FIXTURE, entry.name, true).source,
			implicitMembers: entry.implicitMembers,
		}));
		const project = buildVbaProjectIndex(entries);

		// Inside the form: RegionPick.AddItem "West" is real code in the
		// fixture, and its module text declares no RegionPick anywhere.
		const form = entries.find((entry) => entry.moduleName === 'FrmPicker')!;
		const options = projectAnalysisOptionsForModule(project, 'FrmPicker');
		const diagnostics = analyzeVbaModuleSource({
			source: form.source,
			moduleName: 'FrmPicker',
			moduleType: 'userform',
			moduleKind: 'userform',
			...options,
		} as never).diagnostics;
		expect(diagnostics.filter((d) => d.code === 'undeclared-variable')).toEqual([]);

		// From outside: a standard module chains through the default instance.
		const context = projectEditorSymbolContextForModule(project, 'XlideFormProbe');
		const source = 'Private Sub T()\r\n    FrmPicker.RegionPick.\r\nEnd Sub\r\n';
		const members = resolveMemberCompletions(
			source,
			source.indexOf('RegionPick.') + 'RegionPick.'.length,
			{ projectClassMembers: context.analysisOptions.projectClassMembers } as never,
		).map((member) => member.name);
		expect(members).toContain('AddItem');
	});
});

describe('the form export pair, composed and round-tripped natively', () => {
	function tempCopy(): string {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-form-rt-'));
		tempDirs.push(dir);
		const copy = path.join(dir, 'FormFixture.xlsm');
		fs.copyFileSync(FIXTURE, copy);
		return copy;
	}

	it('composes the .frm from the designer block plus the module text', () => {
		const { frm } = readFormExport(FIXTURE, 'FrmPicker');
		const lines = frm.split('\r\n');
		expect(lines[0]).toBe('VERSION 5.00');
		const blob = lines.find((line) => line.includes('OleObjectBlob'));
		expect(blob).toBe('   OleObjectBlob   =   "FrmPicker.frx":0000');
		// Alphabetical among the properties, the way the VBE writes it.
		expect(lines.indexOf(blob!)).toBeGreaterThan(lines.findIndex((l) => l.includes('ClientWidth')));
		expect(lines.indexOf(blob!)).toBeLessThan(lines.findIndex((l) => l.includes('StartUpPosition')));
		expect(frm).not.toContain('TypeInfoVer');
		expect(frm).toContain('Attribute VB_Name = "FrmPicker"');
		expect(frm).toContain('RegionPick.AddItem');
	});

	it('the .frx carries the designer streams byte-for-byte', () => {
		const { frx } = readFormExport(FIXTURE, 'FrmPicker');
		const unpacked = parseFormFrx(frx)!;
		const { f, o } = fixtureStreams();
		expect(Buffer.compare(unpacked.f, f)).toBe(0);
		expect(Buffer.compare(unpacked.o, o)).toBe(0);
	});

	it('splitFrmSource partitions the pair the way writeModule needs', () => {
		const { frm } = readFormExport(FIXTURE, 'FrmPicker');
		const split = splitFrmSource(frm)!;
		expect(split.designerBlock.startsWith('VERSION 5.00')).toBe(true);
		expect(split.designerBlock.trimEnd().endsWith('End')).toBe(true);
		expect(split.moduleText).toContain('Attribute VB_Name');
		expect(split.moduleText).not.toContain('OleObjectBlob');
	});

	it('writes the designer back and the workbook round-trips identically', () => {
		const copy = tempCopy();
		const { frm, frx } = readFormExport(FIXTURE, 'FrmPicker');
		const result = writeFormDesigner(copy, 'FrmPicker', frx, splitFrmSource(frm)!.designerBlock);
		expect(result.ok).toBe(true);
		resetWorkbookCacheForTests();

		// The written workbook's designer streams equal the original's exactly -
		// the sidecar round-trip loses nothing - and the VBFrame merge restores
		// the workbook shape: TypeInfoVer back, OleObjectBlob gone.
		const original = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(FIXTURE)).readVbaProject());
		const written = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(copy)).readVbaProject());
		for (const stream of ['f', 'o']) {
			expect(
				Buffer.compare(
					written.getStreamInStorage('FrmPicker', stream),
					original.getStreamInStorage('FrmPicker', stream),
				),
				stream,
			).toBe(0);
		}
		const vbFrameName = String.fromCharCode(3) + 'VBFrame';
		expect(written.getStreamInStorage('FrmPicker', vbFrameName).toString('latin1'))
			.toBe(original.getStreamInStorage('FrmPicker', vbFrameName).toString('latin1'));

		// And the engine still reads the same controls out of the written copy.
		const entries = listModules(copy);
		expect(entries.find((entry) => entry.name === 'FrmPicker')?.implicitMembers).toEqual(EXPECTED);
	});

	it('importing the composed .frm keeps the designer block out of the code module', () => {
		const copy = tempCopy();
		const { frm } = readFormExport(FIXTURE, 'FrmPicker');
		const edited = frm.replace('RegionPick.AddItem "West"', 'RegionPick.AddItem "East"');
		const result = writeModule(copy, 'FrmPicker', edited, 'class');
		expect(result.ok).toBe(true);
		resetWorkbookCacheForTests();
		const source = readModule(copy, 'FrmPicker', true).source;
		expect(source).toContain('RegionPick.AddItem "East"');
		expect(source).not.toContain('VERSION 5.00');
		expect(source).not.toContain('OleObjectBlob');
		expect(source).toContain('Attribute VB_Name = "FrmPicker"');
	});
});

describe('sync plan symmetry over the real form workbook', () => {
	// The plan compares live against repo; the exporter writes the same
	// composed .frm, so one export must make the form read as unchanged.
	const engineBridge = {
		call: async <T>(method: string, args: Record<string, unknown>): Promise<T> => {
			const p = String(args.path);
			switch (method) {
				case 'listModules': return listModules(p) as T;
				case 'readModules': return readModulesFull(p) as T;
				case 'readModule': return readModule(p, String(args.module), Boolean(args.full)) as T;
				case 'readFormExport': return readFormExport(p, String(args.module)) as T;
				default: throw new Error(`unexpected: ${method}`);
			}
		},
	} as unknown as import('../src/workbookEngine').WorkbookEngine;

	it('a fresh export leaves the form unchanged in the next plan', async () => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-form-plan-'));
		tempDirs.push(dir);
		const workbook = path.join(dir, 'FormFixture.xlsm');
		fs.copyFileSync(FIXTURE, workbook);
		const repo = path.join(dir, 'repo');
		fs.mkdirSync(repo);
		resetWorkbookCacheForTests();

		const before = await buildExportModuleSyncPlan(engineBridge, {
			workbookPath: workbook, exportFolder: repo,
		});
		const create = before.items.find((item) => item.moduleName === 'FrmPicker');
		expect(create?.relativeName).toBe('FrmPicker.frm');
		expect(create?.status).toBe('will-create');

		// Write exactly what the plan promised: the composed pair.
		const pair = readFormExport(workbook, 'FrmPicker');
		fs.writeFileSync(path.join(repo, 'FrmPicker.frm'), pair.frm, 'utf8');
		fs.writeFileSync(path.join(repo, 'FrmPicker.frx'), pair.frx);

		const after = await buildExportModuleSyncPlan(engineBridge, {
			workbookPath: workbook, exportFolder: repo,
		});
		expect(after.items.find((item) => item.moduleName === 'FrmPicker')?.status).toBe('unchanged');

		// And the import plan agrees from the other direction.
		const importPlan = await buildImportModuleSyncPlan(engineBridge, {
			workbookPath: workbook, importFolder: repo,
		});
		const item = importPlan.items.find((candidate) => candidate.moduleName === 'FrmPicker');
		expect(item?.moduleType).toBe('userform');
		expect(item?.status).toBe('unchanged');
	});
});
