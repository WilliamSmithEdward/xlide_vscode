import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { parseFormDesignerStreams, type DesignerControl } from '../src/vba/formDesigner';
import { listModules, readModule, resetWorkbookCacheForTests } from '../src/vba/workbookService';
import { Cfb } from '../src/vba/cfb';
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
