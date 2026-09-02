import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import {
	FrmHeaderError,
	formatFrmHeader,
	frmControls,
	frmFrxRefs,
	frmMembers,
	frmProperty,
	parseFrmHeader,
	printFrmHeader,
	type FrmHeader,
} from '../src/vba/vb6/frmHeader';
import { decodeFrxList, readFrxRecords } from '../src/vba/vb6/frx';
import { decodeCodePage } from '../src/vba/codePages';
import { readModules, readVb6FormHeaderForTests } from '../src/vba/workbookService';
import { resetVb6ProjectCacheForTests } from '../src/vba/vb6/vb6Project';

// A VB6 form's designer header and its .frx sidecar, measured on files Visual
// Basic 6 wrote (tests/fixtures/vb6/*/NOTICE.md). Two oracles run on every
// fixture: the header printed as read is the file's own bytes, and the header
// REGENERATED from the model in the designer's layout is those same bytes -
// so the layout rules are proven, not assumed.

const ROOT = path.join(__dirname, 'fixtures', 'vb6');
const decode = (b: Buffer): string => decodeCodePage(b, 1252);

function fixtureForms(): string[] {
	const out: string[] = [];
	for (const dir of fs.readdirSync(ROOT)) {
		const full = path.join(ROOT, dir);
		if (!fs.statSync(full).isDirectory()) { continue; }
		for (const file of fs.readdirSync(full)) {
			if (/\.(frm|ctl)$/i.test(file)) { out.push(path.join(full, file)); }
		}
	}
	return out.sort();
}

function load(file: string): { text: string; header: FrmHeader } {
	const text = fs.readFileSync(file, 'latin1');
	const header = parseFrmHeader(text);
	if (!header) { throw new Error(`${file}: no header`); }
	return { text, header };
}

const RUN_AS_TI = path.join(ROOT, 'RunAsTrustedInstaller', 'Form1.frm');
const MCD = path.join(ROOT, 'Diabetes-prediction-1.0', 'MCD.frm');
const OPEN_DIALOG = path.join(ROOT, 'audiostation', 'Form_OpenDialog.frm');
const SETTINGS_RECORD = path.join(ROOT, 'audiostation', 'Form_Settings_Record.frm');
const TRACK_PROPERTIES = path.join(ROOT, 'audiostation', 'Form_Track_Properties.frm');
const HYPERLINK = path.join(ROOT, 'audiostation', 'Hyperlink.ctl');
const MIX_SLIDER = path.join(ROOT, 'audiostation', 'MixSlider.ctl');
const FRM_INFO = path.join(ROOT, 'polyworks', 'frmInfo.frm');
const FRM_PREFERENCES = path.join(ROOT, 'polyworks', 'frmPreferences.frm');

describe('every fixture header, two ways back to bytes', () => {
	const forms = fixtureForms();

	it('has forms to measure', () => {
		expect(forms.length).toBeGreaterThanOrEqual(8);
	});

	for (const file of forms) {
		const name = path.relative(ROOT, file);

		it(`${name}: prints back as read, and ends where the module text begins`, () => {
			const { text, header } = load(file);
			expect(printFrmHeader(header)).toBe(text.slice(0, header.endOffset));
			expect(text.slice(header.endOffset)).toMatch(/^Attribute VB_Name = "/);
		});

		it(`${name}: regenerates from the model in the designer's own layout`, () => {
			const { text, header } = load(file);
			expect(formatFrmHeader(header)).toBe(text.slice(0, header.endOffset));
		});
	}
});

describe('what the header says', () => {
	it('reads the form, its controls, a Font group, and a long-string sidecar reference', () => {
		const { header } = load(RUN_AS_TI);
		expect(header.version).toBe('VERSION 5.00');
		expect(header.objects).toEqual([]);
		expect(header.form.progId).toBe('VB.Form');
		expect(header.form.name).toBe('Form1');
		expect(frmProperty(header.form, 'Caption')?.value).toBe('"Run as ACTUAL administrator"');
		expect(frmProperty(header.form, 'BorderStyle')).toMatchObject({ value: '1', comment: 'Fixed Single' });
		const font = header.form.members.find((m) => m.kind === 'group');
		expect(font).toMatchObject({ kind: 'group', name: 'Font' });
		expect(font?.kind === 'group' && font.members.find((m) => m.kind === 'property' && m.key === 'Name')).toMatchObject({ value: '"Tahoma"' });

		const controls = frmControls(header);
		expect(controls.map((c) => `${c.name}:${c.progId}`)).toEqual(
			expect.arrayContaining(['Text2:VB.TextBox', 'Command1:VB.CommandButton', 'Text1:VB.TextBox']),
		);
		expect(controls.every((c) => c.parent === 'Form1')).toBe(true);

		const refs = frmFrxRefs(header);
		expect(refs).toHaveLength(1);
		expect(refs[0].property.key).toBe('Caption');
		expect(refs[0].property.frx).toEqual({ file: 'Form1.frx', offset: 0, long: true });
	});

	it('reads OCX references, an ImageList of class-id groups, and an SSTab\'s dotted keys', () => {
		const { header } = load(OPEN_DIALOG);
		expect(header.objects.length).toBeGreaterThan(0);
		expect(header.objects[0]).toMatchObject({ guid: expect.stringMatching(/^\{[0-9A-F-]+\}$/i), file: expect.stringMatching(/\.ocx$/i) });

		const imageList = frmControls(header).find((c) => /ImageList$/i.test(c.progId));
		expect(imageList).toBeDefined();
		const images = imageList!.members.find((m) => m.kind === 'group' && m.name === 'Images');
		expect(images).toMatchObject({ kind: 'group', guid: expect.stringMatching(/^\{/) });
		expect(images?.kind === 'group' && images.members.filter((m) => m.kind === 'group' && /^ListImage\d+$/.test(m.name)).length).toBeGreaterThan(0);

		const tab = frmControls(header).find((c) => /SSTab$/i.test(c.progId));
		expect(tab).toBeDefined();
		expect(frmProperty(tab!, 'Tab(0).ControlCount')?.value).toBe('2');
		expect(frmProperty(tab!, 'TabPicture(0)')?.frx).toMatchObject({ long: false, offset: 0x0C });
	});

	it('reads menus as nested controls and folds a control array into one member', () => {
		const { header } = load(FRM_INFO);
		const menus = frmControls(header).filter((c) => c.progId === 'VB.Menu');
		expect(menus.length).toBeGreaterThan(0);
		expect(menus.some((m) => m.children.length > 0)).toBe(true);

		const members = frmMembers(header);
		const names = members.map((m) => m.name.toLowerCase());
		expect(new Set(names).size).toBe(names.length);
		expect(members.some((m) => m.array)).toBe(true);
		const indexed = frmControls(header).filter((c) => c.index !== undefined);
		expect(indexed.length).toBeGreaterThan(1);
		const arrayName = indexed[0].name;
		expect(members.find((m) => m.name === arrayName)).toMatchObject({ array: true, type: indexed[0].progId });
	});

	it('refuses a header that never closes, and answers undefined for no header', () => {
		expect(parseFrmHeader('Option Explicit\r\nSub A()\r\nEnd Sub\r\n')).toBeUndefined();
		expect(() => parseFrmHeader('VERSION 5.00\r\nBegin VB.Form F \r\n   Caption = "x"\r\n')).toThrow(FrmHeaderError);
		expect(() => parseFrmHeader('VERSION 5.00\r\nEnd\r\n')).toThrow(/no open control block/);
	});

	it('lays out a synthetic model the way the designer does', () => {
		const header = parseFrmHeader([
			'VERSION 5.00',
			'Begin VB.Form Form1 ',
			'   Caption         =   "F"',
			'   Tab(0).ControlEnabled=   -1  \'True',
			'   Object.Width           =   80',
			'   BeginProperty Font ',
			'      Name            =   "Tahoma"',
			'   EndProperty',
			'   Begin VB.Label Label1 ',
			'      Index           =   0',
			'   End',
			'End',
			'Attribute VB_Name = "Form1"',
			'',
		].join('\r\n'))!;
		expect(formatFrmHeader(header)).toBe([
			'VERSION 5.00',
			'Begin VB.Form Form1 ',
			'   Caption         =   "F"',
			'   Tab(0).ControlEnabled=   -1  \'True',
			'   Object.Width           =   80',
			'   BeginProperty Font ',
			'      Name            =   "Tahoma"',
			'   EndProperty',
			'   Begin VB.Label Label1 ',
			'      Index           =   0',
			'   End',
			'End',
			'',
		].join('\r\n'));
	});
});

describe('the .frx sidecar', () => {
	function records(file: string) {
		const { header } = load(file);
		const frx = frmFrxRefs(header)[0]?.property.frx?.file;
		const blob = fs.readFileSync(path.join(path.dirname(file), frx!));
		return readFrxRecords(header, blob, decode);
	}

	it('decodes a long string: 32-bit length, then the text', () => {
		const [caption] = records(RUN_AS_TI);
		expect(caption.property).toBe('Caption');
		expect(caption.bytes).toHaveLength(147);
		expect(caption.value).toMatchObject({ kind: 'longString', text: expect.stringMatching(/^Running as TrustedInstaller/) });
		expect((caption.value as { text: string }).text).toHaveLength(143);
	});

	it('decodes a picture and a short string from one sidecar', () => {
		const all = records(MCD);
		const picture = all.find((r) => r.property === 'Picture');
		expect(picture?.value).toMatchObject({ kind: 'picture' });
		expect((picture?.value as { bytes: Buffer }).bytes).toHaveLength(37596);
		const text = all.find((r) => r.property === 'Text');
		expect(text?.offset).toBe(0x92e0);
		expect(text?.value).toMatchObject({ kind: 'shortString', text: expect.stringMatching(/^159,82,187,/) });
		expect((text?.value as { text: string }).text).toHaveLength(157);
	});

	it('decodes ComboBox rows, an empty picture, and an indexed picture', () => {
		const rows = records(SETTINGS_RECORD);
		const list = rows.find((r) => r.property === 'List');
		expect(list?.value).toMatchObject({ kind: 'list' });
		expect((list?.value as { items: string[] }).items[0]).toBe('English');
		const itemData = rows.find((r) => r.property === 'ItemData');
		expect((itemData?.value as { items: string[] }).items).toEqual(['0', '0', '0']);
		const icon = rows.find((r) => r.property === 'Icon');
		expect(icon?.value).toMatchObject({ kind: 'picture' });
		expect((icon?.value as { bytes: Buffer }).bytes).toEqual(Buffer.from('6c74000000000000', 'hex'));

		const dialog = records(OPEN_DIALOG);
		expect(dialog.find((r) => r.property === 'TabPicture(0)')?.value).toMatchObject({ kind: 'picture' });

		const shortText = records(TRACK_PROPERTIES).find((r) => r.property === 'Text');
		expect(shortText?.value).toMatchObject({ kind: 'shortString', text: expect.stringMatching(/^Textbox/) });

		const control = records(HYPERLINK);
		expect(control.map((r) => r.property)).toEqual(['ToolboxBitmap', 'MouseIcon']);
		expect(control.every((r) => r.value.kind === 'picture')).toBe(true);

		// A UserControl hosting other controls carries their extender
		// properties (`Object.Width`), the one key family padded wider.
		const { header } = load(MIX_SLIDER);
		const hosted = frmControls(header).find((c) => c.members.some((m) => m.kind === 'property' && m.key.startsWith('Object.')));
		expect(hosted).toBeDefined();
	});

	it('reads an empty list as no rows and refuses a truncated one', () => {
		expect(decodeFrxList(Buffer.from('0000', 'hex'), decode)).toEqual([]);
		expect(decodeFrxList(Buffer.from('01000300', 'hex'), decode)).toBeUndefined();
	});

	it('covers every referenced byte: the records tile the sidecar', () => {
		for (const file of [MCD, SETTINGS_RECORD, FRM_PREFERENCES, OPEN_DIALOG]) {
			const all = records(file);
			const total = all.reduce((sum, r) => sum + r.bytes.length, 0);
			const frx = frmFrxRefs(load(file).header)[0].property.frx!.file;
			expect(total, file).toBe(fs.statSync(path.join(path.dirname(file), frx)).size);
			expect(all.every((r) => r.value.kind !== 'opaque'), file).toBe(true);
		}
	});
});

describe('the project reader hands the analyzer the designer\'s controls', () => {
	it('lists a real VB6 form\'s controls as implicit members, typed by prog id', () => {
		resetVb6ProjectCacheForTests();
		const vbp = path.join(ROOT, 'RunAsTrustedInstaller', 'Project1.vbp');
		const form = readModules(vbp, true).find((m) => m.name === 'Form1');
		expect(form?.implicitMembers).toEqual(expect.arrayContaining([
			{ name: 'Text2', type: 'VB.TextBox' },
			{ name: 'Command1', type: 'VB.CommandButton' },
			{ name: 'Text1', type: 'VB.TextBox' },
		]));
		expect(readVb6FormHeaderForTests(vbp, 'Form1')?.form.name).toBe('Form1');
		resetVb6ProjectCacheForTests();
	});
});
