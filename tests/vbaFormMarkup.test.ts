import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
	readFormMarkup,
	applyFormMarkup,
	addFormModule,
	readModules,
	readFormExport,
	resetWorkbookCacheForTests,
} from '../src/vba/workbookService';
import { parseFormFrx, parseFormDesignerStreams } from '../src/vba/formDesigner';
import { parseFormMarkup, formatOleColor, parseOleColor } from '../src/vba/oforms/markup';

// The form-as-text projection, whole loop: print a real Excel-authored form
// to markup, edit the document, apply it back as a name-keyed diff, and
// author a new form from nothing. The dialect is xlide_vbide's, adopted here
// so both products speak the same document.
//
// Every mutation asserted below was ALSO verified against live Excel once,
// during development: a workbook carrying the caption change, the moved
// button, the added Label, and a from-scratch FrmFresh compiled and
// instantiated its forms under Application.Run. That proof cannot run in CI;
// what stands in for it here is byte-identity on no-op plus re-read of every
// mutation, and one cross-check through the independent formDesigner reader.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');
const CRLF = '\r\n';

const tempDirs: string[] = [];
afterEach(() => {
	resetWorkbookCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function workbook(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-form-'));
	tempDirs.push(dir);
	const wb = path.join(dir, 'Forms.xlsm');
	fs.copyFileSync(FIXTURE, wb);
	return wb;
}

describe('the printed document', () => {
	it('projects the whole control tree in the shared dialect', () => {
		const markup = readFormMarkup(workbook(), 'EntryForm').markup;
		expect(markup).toContain('<Form Name="EntryForm" Caption="Quarter Entry" Width="348" Height="291">');
		expect(markup).toContain('<Label Name="NameLabel" Left="12" Top="14" Width="66" Height="16" Caption="Customer"');
		expect(markup).toContain('<Frame Name="Options" Caption="Freight"');
		expect(markup).toContain('<OptionButton Name="PickGround"');
		expect(markup).toContain('<MultiPage Name="Wizard"');
		expect(markup).toContain('<Page Name="Page1" Caption="Page1">');
		expect(markup).toContain('<Page Name="Page2" Caption="Page2" />');
		expect(markup).toContain('<Tab Caption="Tab1" />');
		expect(markup).toContain('Font.Name="Tahoma" Font.Size="8.25"');
		expect(markup).toContain('BackColor="ButtonFace"');
	});

	it('prints geometry as the shortest points that re-encode identically', () => {
		// Excel stores 12pt as round(12 * 2540/72) = 423 HIMETRIC; the naive
		// back-conversion says 11.99, the designer says 12, and so does this.
		const markup = readFormMarkup(workbook(), 'EntryForm').markup;
		expect(markup).toContain('Left="12" Top="14"');
		expect(markup).not.toContain('11.99');
	});
});

describe('colors', () => {
	it('spells a system color by name and a literal as hex', () => {
		expect(formatOleColor(0x8000000f)).toBe('ButtonFace');
		expect(formatOleColor(0x00c0dcc0)).toBe('#c0dcc0');
	});

	it('parses all three spellings back', () => {
		expect(parseOleColor('ButtonFace')).toBe(0x8000000f);
		expect(parseOleColor('#c0dcc0')).toBe(0x00c0dcc0);
		expect(parseOleColor('&H8000000F&')).toBe(0x8000000f);
		expect(parseOleColor('nonsense')).toBeUndefined();
	});
});

describe('a no-op apply', () => {
	it('changes nothing and reports nothing', () => {
		const wb = workbook();
		const before = fs.readFileSync(wb);
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		const result = applyFormMarkup(wb, 'EntryForm', markup);
		expect(result.applied).toEqual([]);
		// Nothing applied means nothing SAVED: the file bytes are untouched.
		expect(fs.readFileSync(wb).equals(before)).toBe(true);
	});
});

describe('mutations, re-read from the workbook', () => {
	it('moves a control, recaptions the form, adds and removes controls', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		const edited = markup
			.replace('Caption="Quarter Entry"', 'Caption="Quarter Entry v2"')
			.replace('<CommandButton Name="OkButton" Left="262"', '<CommandButton Name="OkButton" Left="250"')
			.replace(/<TabStrip Name="Views"[\s\S]*?<\/TabStrip>\r\n/, '')
			.replace('</Form>', `    <Label Name="AddedNote" Left="12" Top="270" Width="150" Height="14" Caption="added by markup" />${CRLF}</Form>`);
		const result = applyFormMarkup(wb, 'EntryForm', edited);
		expect(result.applied).toContain('Caption of the form');
		expect(result.applied).toContain('position of OkButton');
		expect(result.applied).toContain('removed Views');
		expect(result.applied).toContain('added Label AddedNote');

		resetWorkbookCacheForTests();
		const after = readFormMarkup(wb, 'EntryForm').markup;
		expect(after).toContain('Caption="Quarter Entry v2"');
		expect(after).toContain('<CommandButton Name="OkButton" Left="250"');
		expect(after).not.toContain('<TabStrip Name="Views"');
		expect(after).toContain('<Label Name="AddedNote"');
		// The untouched neighbourhood survived whole.
		expect(after).toContain('<Frame Name="Options" Caption="Freight"');
		expect(after).toContain('<Page Name="Page1" Caption="Page1">');
	});

	it('assigns an added control an ID above every live one', () => {
		// The fixture carries NextAvailableID EQUAL to OkButton's ID (20), so
		// trusting the field alone reuses a live ID - and duplicate site IDs
		// kill the form at load. Found by real Excel refusing to compile.
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		applyFormMarkup(wb, 'EntryForm', markup.replace(
			'</Form>',
			`    <Label Name="AddedNote" Left="12" Top="270" Width="150" Height="14" />${CRLF}</Form>`,
		));
		resetWorkbookCacheForTests();
		const { frx } = readFormExport(wb, 'EntryForm');
		const streams = parseFormFrx(frx)!;
		const controls = parseFormDesignerStreams(streams.f, streams.o, (b, c) => b.toString(c ? 'latin1' : 'utf16le'));
		// The independent reader parses the mutated form and sees the addition.
		expect(controls?.map((c) => c.name)).toContain('AddedNote');
	});

	it('edits controls inside a Frame and captions of MultiPage pages', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		const edited = markup
			.replace('<OptionButton Name="PickAir" Left="8"', '<OptionButton Name="PickAir" Left="10"')
			.replace('<Page Name="Page1" Caption="Page1">', '<Page Name="Page1" Caption="Basics">');
		const result = applyFormMarkup(wb, 'EntryForm', edited);
		expect(result.applied).toContain('position of PickAir');
		expect(result.applied).toContain('page captions of Wizard');
		resetWorkbookCacheForTests();
		const after = readFormMarkup(wb, 'EntryForm').markup;
		expect(after).toContain('<OptionButton Name="PickAir" Left="10"');
		expect(after).toContain('Caption="Basics"');
	});
});

describe('what the apply refuses, whole-document', () => {
	it('refuses an unquoted value with its line', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup
			.replace('Left="250"', 'Left=250')
			.replace('Left="262"', 'Left=262');
		expect(() => applyFormMarkup(wb, 'EntryForm', markup)).toThrow(/QUOTED/);
	});

	it('refuses two controls with one name', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup
			.replace('<Label Name="ViewNote"', '<Label Name="NameLabel"');
		expect(() => applyFormMarkup(wb, 'EntryForm', markup)).toThrow(/named NameLabel/);
	});

	it('refuses adding what it cannot author yet, by name', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(() => applyFormMarkup(wb, 'EntryForm', markup.replace(
			'</Form>',
			`    <ActiveX Name="Web1" Left="0" Top="0" />${CRLF}</Form>`,
		))).toThrow(/class table/);
		expect(() => applyFormMarkup(wb, 'EntryForm', markup.replace(
			'<Page Name="Page2" Caption="Page2" />' + CRLF,
			'',
		))).toThrow(/adding or removing pages/);
	});

	it('applies nothing when any part of the document is broken', () => {
		const wb = workbook();
		const before = fs.readFileSync(wb);
		const markup = readFormMarkup(wb, 'EntryForm').markup
			.replace('Caption="Quarter Entry"', 'Caption="Half Entry"')  // a valid edit...
			.replace('</Form>', '');                                     // ...in a broken document
		expect(() => applyFormMarkup(wb, 'EntryForm', markup)).toThrow(/never closed/);
		expect(fs.readFileSync(wb).equals(before)).toBe(true);
	});
});

describe('a form from nothing', () => {
	it('creates a UserForm the reader stack fully recognises', () => {
		const wb = workbook();
		addFormModule(wb, 'FrmFresh', `Option Explicit${CRLF}`);
		resetWorkbookCacheForTests();
		const mods = readModules(wb, false);
		const fresh = mods.find((m) => m.name === 'FrmFresh');
		expect(fresh?.type).toBe('userform');
		// The attribute header this engine writes is the one #47's plumbing
		// reads: a fresh form has a default instance.
		expect(fresh?.predeclaredId).toBe(true);
		expect(fresh?.implicitMembers).toEqual([]);
		expect(readFormMarkup(wb, 'FrmFresh').markup)
			.toContain('<Form Name="FrmFresh" Caption="FrmFresh" Width="240" Height="180" />');
	});

	it('takes controls through markup, and the analyzer sees them', () => {
		const wb = workbook();
		addFormModule(wb, 'FrmFresh', `Option Explicit${CRLF}`);
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'FrmFresh').markup;
		const result = applyFormMarkup(wb, 'FrmFresh', markup.replace(
			/<Form([^>]*?)\s*\/>/,
			`<Form$1>${CRLF}    <CommandButton Name="GoBtn" Left="80" Top="120" Width="72" Height="24" Caption="Go" />${CRLF}</Form>`,
		));
		expect(result.applied).toContain('added CommandButton GoBtn');
		resetWorkbookCacheForTests();
		const fresh = readModules(wb, false).find((m) => m.name === 'FrmFresh');
		// The designer-declared control surface feeds completion/diagnostics.
		expect(fresh?.implicitMembers).toEqual([{ name: 'GoBtn', type: 'MSForms.CommandButton' }]);
	});
});

describe('the parser alone', () => {
	it('reads attributes across wrapped lines', () => {
		const root = parseFormMarkup([
			'<Form Name="F">',
			'    <CommandButton Name="Ok" Left="1" Top="2"',
			'                   Caption="Start" />',
			'</Form>',
		].join('\n'));
		expect(root.children[0].attrs.get('Caption')).toBe('Start');
	});

	it('skips comments and unescapes attribute text', () => {
		const root = parseFormMarkup([
			'<!-- a comment -->',
			'<Form Name="F" Caption="A &quot;B&quot; &amp; C" />',
		].join('\n'));
		expect(root.attrs.get('Caption')).toBe('A "B" & C');
	});
});

describe('spoken font and tab edits land', () => {
	it('applies Font.Size, Font.Bold, and a tab caption', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		const edited = markup
			.replace('<CommandButton Name="OkButton" Left="262" Top="250" Width="72" Height="24" Caption="Start" Font.Name="Tahoma" Font.Size="8.25"',
				'<CommandButton Name="OkButton" Left="262" Top="250" Width="72" Height="24" Caption="Start" Font.Name="Tahoma" Font.Size="12" Font.Bold="True"')
			.replace('<Tab Caption="Tab1" />', '<Tab Caption="Overview" />');
		const result = applyFormMarkup(wb, 'EntryForm', edited);
		expect(result.applied).toContain('Font.Size of OkButton');
		expect(result.applied).toContain('Font style of OkButton');
		expect(result.applied).toContain('tab captions of Views');
		resetWorkbookCacheForTests();
		const after = readFormMarkup(wb, 'EntryForm').markup;
		expect(after).toContain('Font.Size="12"');
		expect(after).toMatch(/OkButton[^\r\n]*Font\.Bold="True"/);
		expect(after).toContain('<Tab Caption="Overview" />');
	});

	it('refuses a font on a kind that carries none', () => {
		const wb = workbook();
		const markup = readFormMarkup(wb, 'EntryForm').markup
			.replace('<Image Name="Badge"', '<Image Name="Badge" Font.Size="10"');
		expect(() => applyFormMarkup(wb, 'EntryForm', markup)).toThrow(/carries no font/);
	});
});
