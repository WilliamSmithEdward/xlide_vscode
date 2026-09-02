import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { frmControls, parseFrmHeader } from '../src/vba/vb6/frmHeader';
import type { FrmHeader } from '../src/vba/vb6/frmHeader';
import { readFrxRecords } from '../src/vba/vb6/frx';
import { decodeCodePage } from '../src/vba/codePages';
import {
	VB6_CONTROLS, VB6_DEFAULT_EVENTS, VB6_DESIGNER_ONLY_KEYS, VB6_DESIGN_PROPERTIES, VB6_TOOLBOX, listFrmProperties,
	pictureDataUriOf, sceneOfFrmHeader, twipsToPt, vb6CanvasKind, vb6ControlSpec, vb6MenuCaptions, vb6PaneVocabulary,
} from '../src/vba/vb6/frmScene';
import { applyFrmDesignerOp } from '../src/vba/vb6/frmDesignerOps';
import { getVb6ObjectModel } from '../src/analyzer/host/vb6ObjectModel';
import { renderFormSceneHtml } from '../src/vba/oforms/preview';
import type { SceneControl } from '../src/vba/oforms/preview';
import { formatOleColor, parseOleColor } from '../src/vba/oforms/markup';
import { frmDisplayValue } from '../src/vba/vb6/frmScene';

// The VB6 scene adapter, measured on every fixture form: the canvas draws
// each control the header holds, at the bounds the header states, in points.

const FIXTURES = path.join(__dirname, 'fixtures', 'vb6');
const decode = (b: Buffer): string => decodeCodePage(b, 1252);

function formFiles(dir = FIXTURES): string[] {
	const out: string[] = [];
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name);
		if (entry.isDirectory()) { out.push(...formFiles(full)); }
		else if (/[.](frm|ctl)$/i.test(entry.name)) { out.push(full); }
	}
	return out.sort();
}

function headerOf(file: string): FrmHeader {
	const header = parseFrmHeader(fs.readFileSync(file, 'latin1'));
	if (!header) { throw new Error(`not a designer file: ${file}`); }
	return header;
}

function flatten(controls: readonly SceneControl[]): SceneControl[] {
	return controls.flatMap((c) => [c, ...flatten(c.children)]);
}

function fixture(rel: string): string {
	return path.join(FIXTURES, ...rel.split('/'));
}

describe('twipsToPt', () => {
	it('prints a twentieth of a point, shortest', () => {
		expect(twipsToPt(240)).toBe('12');
		expect(twipsToPt(250)).toBe('12.5');
		expect(twipsToPt(1)).toBe('0.05');
		expect(twipsToPt(0)).toBe('0');
		expect(twipsToPt(-15)).toBe('-0.75');
	});
});

describe('vb6CanvasKind', () => {
	it('maps the intrinsic controls and nothing else', () => {
		expect(vb6CanvasKind('VB.CommandButton')).toBe('CommandButton');
		expect(vb6CanvasKind('VB.HScrollBar')).toBe('ScrollBar');
		expect(vb6CanvasKind('VB.VScrollBar')).toBe('ScrollBar');
		expect(vb6CanvasKind('VB.Menu')).toBeUndefined();
		expect(vb6CanvasKind('MSComctlLib.ListView')).toBeUndefined();
		expect(vb6CanvasKind('Audiostation.ButtonBig')).toBeUndefined();
	});
	it('offers every intrinsic control the toolbox has, including the file-system, data and OLE controls', () => {
		for (const kind of VB6_TOOLBOX) {
			expect(vb6CanvasKind(`VB.${kind}`), kind).toBeDefined();
		}
		expect(VB6_TOOLBOX).toEqual(['Label', 'TextBox', 'ComboBox', 'ListBox', 'CheckBox', 'OptionButton',
			'CommandButton', 'Frame', 'PictureBox', 'Image', 'HScrollBar', 'VScrollBar', 'Timer', 'Line', 'Shape',
			'DriveListBox', 'DirListBox', 'FileListBox', 'Data', 'OLE']);
		// Every intrinsic control type the model knows is a kind the designer draws.
		const model = getVb6ObjectModel();
		const controlTypes = Object.keys(model.types).filter((name) => {
			const members = model.types[name].members;
			return name.startsWith('VB.') && members.some((m) => m.kind === 'property' && m.name === 'TabIndex');
		});
		for (const progId of controlTypes) {
			expect(VB6_CONTROLS[progId], progId).toBeDefined();
		}
	});

	it('describes each kind once, and the designer classes as no control', () => {
		expect(vb6CanvasKind('VB.Form')).toBeUndefined();
		expect(vb6ControlSpec('VB.Frame')).toMatchObject({ kind: 'Frame', container: true, captioned: true, base: 'Frame' });
		expect(vb6ControlSpec('VB.Timer')?.size).toBeUndefined();
		expect(VB6_DESIGN_PROPERTIES['VB.Label']).toContain('Caption');
		for (const [progId, spec] of Object.entries(VB6_CONTROLS)) {
			expect(progId.startsWith('VB.'), progId).toBe(true);
			expect(spec.vocabularyFrom, progId).toMatch(/^(fixtures|model)$/);
			if (spec.text === 'Caption') { expect(spec.captioned, progId).toBe(true); }
			if (spec.scale) { expect(spec.size, progId).toBeDefined(); }
			if (spec.base && spec.kind !== 'Timer' && spec.kind !== 'Line') { expect(spec.size, progId).toBeDefined(); }
		}
	});

	it('names only properties the model declares, bar the keys only the designer writes', () => {
		const model = getVb6ObjectModel();
		const designerOnly = new Set(VB6_DESIGNER_ONLY_KEYS.map((k) => k.toLowerCase()));
		for (const [progId, spec] of Object.entries(VB6_CONTROLS)) {
			const type = model.types[progId];
			expect(type, progId).toBeDefined();
			const known = new Set(type.members.filter((m) => m.kind === 'property').map((m) => m.name.toLowerCase()));
			for (const key of spec.designProperties) {
				if (key === 'Font' || designerOnly.has(key.toLowerCase())) { continue; }
				expect(known.has(key.toLowerCase()), `${progId}.${key}`).toBe(true);
			}
		}
	});

	it('opens a double-click on an event the kind actually has, or refuses by name', () => {
		const model = getVb6ObjectModel();
		for (const [progId, spec] of Object.entries(VB6_CONTROLS)) {
			const events = new Set((model.types[progId]?.members ?? [])
				.filter((m) => m.kind === 'event').map((m) => m.name.toLowerCase()));
			// The canvas falls back to Click for a kind with no entry.
			const effective = spec.defaultEvent ?? 'Click';
			if (effective === '') {
				// Only for a kind the designer refuses outright.
				expect(VB6_DEFAULT_EVENTS[spec.kind], progId).toBe('');
				continue;
			}
			expect(events.has(effective.toLowerCase()), `${progId} has no ${effective} event`).toBe(true);
		}
		// The kinds VB6 gives no Click: they must not fall back to one.
		for (const progId of ['VB.HScrollBar', 'VB.VScrollBar', 'VB.Data', 'VB.DriveListBox']) {
			expect(VB6_CONTROLS[progId].defaultEvent, progId).toBeTruthy();
		}
		expect(VB6_DEFAULT_EVENTS.ScrollBar).toBe('Change');
		expect(VB6_DEFAULT_EVENTS.Data).toBe('Reposition');
	});

	it('draws the file-system, data and OLE controls rather than a foreign box', () => {
		const header = headerOf(fixture('RunAsTrustedInstaller/Form1.frm'));
		let text = fs.readFileSync(fixture('RunAsTrustedInstaller/Form1.frm'), 'latin1');
		for (const kind of ['DriveListBox', 'DirListBox', 'FileListBox', 'Data', 'OLE']) {
			text = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: kind, left: 5, top: 5 }).text;
		}
		const grown = parseFrmHeader(text)!;
		const scene = sceneOfFrmHeader(grown, { formName: 'Form1' });
		const drawn = new Map(flatten(scene.controls).map((c) => [c.name, c]));
		expect(drawn.get('Drive1')?.kind).toBe('DriveListBox');
		expect(drawn.get('Dir1')?.kind).toBe('DirListBox');
		expect(drawn.get('File1')?.kind).toBe('FileListBox');
		expect(drawn.get('Data1')?.kind).toBe('Data');
		expect(drawn.get('OLE1')?.kind).toBe('OLE');
		expect(flatten(scene.controls).some((c) => c.kind === 'Foreign')).toBe(false);
		const html = renderFormSceneHtml(scene, { formName: 'Form1' });
		expect(html).toContain('class="ctl edit combo" data-name="Drive1"');
		expect(html).toContain('class="ctl edit list" data-name="Dir1"');
		expect(html).toContain('class="ctl data" data-name="Data1"');
		expect(html).toContain('class="ctl ole" data-name="OLE1"');
		expect(html).toContain('.data .nav');
		// The Data control shows its caption between the navigator's arrows.
		expect(html).toContain('<span class="cap">Data1</span>');
		expect(header.form.children.length).toBeLessThan(grown.form.children.length);
	});

	it('lists a vocabulary for the kinds no fixture uses, and says where it came from', () => {
		for (const progId of ['VB.Timer', 'VB.ListBox', 'VB.Image', 'VB.HScrollBar', 'VB.Data', 'VB.FileListBox', 'VB.OLE']) {
			expect(VB6_CONTROLS[progId].vocabularyFrom, progId).toBe('model');
			expect(VB6_DESIGN_PROPERTIES[progId].length, progId).toBeGreaterThan(0);
		}
		expect(VB6_DESIGN_PROPERTIES['VB.Timer']).toEqual(['Enabled', 'Interval', 'Tag']);
		expect(VB6_DESIGN_PROPERTIES['VB.HScrollBar']).toContain('LargeChange');
		expect(VB6_CONTROLS['VB.Label'].vocabularyFrom).toBe('fixtures');
		// A scroll bar has no ToolTipText, and an OLE container none either: the model says so.
		expect(VB6_DESIGN_PROPERTIES['VB.VScrollBar']).not.toContain('ToolTipText');
		expect(VB6_DESIGN_PROPERTIES['VB.OLE']).not.toContain('ToolTipText');
	});
});

describe('sceneOfFrmHeader on every fixture form', () => {
	const files = formFiles();
	it('has fixtures to measure', () => {
		expect(files.length).toBeGreaterThan(5);
	});

	for (const file of files) {
		const rel = path.relative(FIXTURES, file).split(path.sep).join('/');
		it(`draws every non-menu control of ${rel} at its header bounds`, () => {
			const header = headerOf(file);
			const scene = sceneOfFrmHeader(header, { formName: header.form.name });
			const drawn = flatten(scene.controls);
			const expected = frmControls(header).filter((c) => c.progId !== 'VB.Menu');
			expect(drawn.length).toBe(expected.length);
			const byName = new Map(drawn.map((c) => [c.name, c]));
			for (const control of expected) {
				const name = control.index === undefined ? control.name : `${control.name}(${control.index})`;
				const drawnControl = byName.get(name);
				expect(drawnControl, name).toBeDefined();
				const kind = vb6CanvasKind(control.progId);
				expect(drawnControl!.kind).toBe(kind ?? 'Foreign');
				const prop = (key: string): number | undefined => {
					const m = control.members.find((x) => x.kind === 'property' && x.key.toLowerCase() === key.toLowerCase());
					return m && m.kind === 'property' && !m.frx ? Number(m.value.trim()) : undefined;
				};
				if (kind === 'Line') {
					const x1 = prop('X1') ?? 0;
					expect(drawnControl!.style).toContain(`left:${twipsToPt(Math.min(x1, prop('X2') ?? x1))}pt;`);
					continue;
				}
				for (const key of ['Left', 'Top']) {
					const twips = prop(key);
					if (twips !== undefined) {
						expect(drawnControl!.style, `${name}.${key}`).toContain(`${key.toLowerCase()}:${twipsToPt(twips)}pt;`);
					}
				}
				if (kind !== 'Timer') {
					for (const key of ['Width', 'Height']) {
						const twips = prop(key);
						if (twips !== undefined) {
							expect(drawnControl!.style, `${name}.${key}`).toContain(`${key.toLowerCase()}:${twipsToPt(twips)}pt;`);
						}
					}
				}
			}
			const html = renderFormSceneHtml(scene, { formName: header.form.name });
			for (const control of drawn) {
				expect(html).toContain(`data-name="${control.name.replace(/&/g, '&amp;')}"`);
			}
			// The script itself says isNaN and typeof undefined; a style must not.
			expect(html).not.toMatch(/:(NaN|undefined)(pt)?;/);
			expect(html).not.toMatch(/\b(NaN|undefined)pt\b/);
		});
	}
});

describe('sceneOfFrmHeader details', () => {
	it('draws the menu bar of frmInfo from its top-level menus, ampersands dropped', () => {
		const header = headerOf(fixture('polyworks/frmInfo.frm'));
		const menus = vb6MenuCaptions(header.form);
		expect(menus.length).toBeGreaterThan(0);
		for (const caption of menus) { expect(caption).not.toContain('&'); }
		const scene = sceneOfFrmHeader(header, { formName: 'frmInfo' });
		expect(scene.form.menus).toEqual(menus);
		const html = renderFormSceneHtml(scene, { formName: 'frmInfo' });
		expect(html).toContain('class="menubar"');
		expect(flatten(scene.controls).some((c) => c.name.startsWith('mnu'))).toBe(false);
	});

	it('renders no menu bar for a form without menus', () => {
		const header = headerOf(fixture('audiostation/Form_OpenDialog.frm'));
		const scene = sceneOfFrmHeader(header, { formName: 'Form_OpenDialog' });
		expect(scene.form.menus).toEqual([]);
		expect(renderFormSceneHtml(scene, { formName: 'Form_OpenDialog' })).not.toContain('class="menubar"');
	});

	it('takes the form size from ClientWidth and ClientHeight, the caption from Caption', () => {
		const header = headerOf(fixture('audiostation/Form_OpenDialog.frm'));
		const prop = (key: string): string => {
			const m = header.form.members.find((x) => x.kind === 'property' && x.key === key);
			return m && m.kind === 'property' ? m.value.trim() : '';
		};
		const scene = sceneOfFrmHeader(header, { formName: 'Form_OpenDialog' });
		expect(scene.form.widthPt).toBe(twipsToPt(Number(prop('ClientWidth'))));
		expect(scene.form.heightPt).toBe(twipsToPt(Number(prop('ClientHeight'))));
		expect(scene.form.caption).toBe(prop('Caption').replace(/^"|"$/g, ''));
		expect(scene.toolbox).toEqual(VB6_TOOLBOX);
	});

	it('draws a Line between its points and a Shape in its shape', () => {
		const lines = formFiles().filter((f) => fs.readFileSync(f, 'latin1').includes('Begin VB.Line'));
		const shapes = formFiles().filter((f) => fs.readFileSync(f, 'latin1').includes('Begin VB.Shape'));
		expect(lines.length).toBeGreaterThan(0);
		expect(shapes.length).toBeGreaterThan(0);
		const lineScene = sceneOfFrmHeader(headerOf(lines[0]), { formName: 'f' });
		const line = flatten(lineScene.controls).find((c) => c.kind === 'Line');
		expect(line).toBeDefined();
		const lineHtml = renderFormSceneHtml(lineScene, { formName: 'f' });
		expect(lineHtml).toContain('class="ctl line"');
		expect(lineHtml).toContain('<svg');
		const shapeScene = sceneOfFrmHeader(headerOf(shapes[0]), { formName: 'f' });
		const shape = flatten(shapeScene.controls).find((c) => c.kind === 'Shape');
		expect(shape).toBeDefined();
		expect(shape!.style).toContain('border-radius:');
	});

	it('nests the children of a PictureBox under its surface', () => {
		const files = formFiles().filter((f) => fs.readFileSync(f, 'latin1').includes('Begin VB.PictureBox'));
		expect(files.length).toBeGreaterThan(0);
		let nested = 0;
		for (const file of files) {
			const header = headerOf(file);
			const scene = sceneOfFrmHeader(header, { formName: header.form.name });
			for (const box of flatten(scene.controls).filter((c) => c.kind === 'PictureBox')) {
				nested += box.children.length;
			}
			const html = renderFormSceneHtml(scene, { formName: header.form.name });
			expect(html).toContain('class="ctl frame picture');
		}
		expect(nested).toBeGreaterThan(0);
	});

	it('shows a checked CheckBox and a true OptionButton as on', () => {
		let checked = 0;
		let on = 0;
		for (const file of formFiles()) {
			const header = headerOf(file);
			for (const c of flatten(sceneOfFrmHeader(header, { formName: header.form.name }).controls)) {
				if (c.kind === 'CheckBox' && c.on) { checked += 1; }
				if (c.kind === 'OptionButton' && c.on) { on += 1; }
			}
		}
		expect(checked).toBeGreaterThan(0);
		expect(on).toBeGreaterThan(0);
	});

	it('paints a sidecar icon the browser can show, and leaves the form bare without the sidecar', () => {
		const file = fixture('audiostation/Form_OpenDialog.frm');
		const header = headerOf(file);
		const blob = fs.readFileSync(file.replace(/[.]frm$/i, '.frx'));
		const records = readFrxRecords(header, blob, decode);
		const pictures = records.filter((r) => r.value.kind === 'picture');
		expect(pictures.length).toBeGreaterThan(0);
		let painted = 0;
		for (const record of pictures) {
			if (record.value.kind !== 'picture') { continue; }
			const uri = pictureDataUriOf(record.value.bytes);
			// An empty picture is `lt` and zeros, with or without a class id
			// in front (8 or 24 bytes); anything longer holds an image.
			if (record.value.bytes.length <= 24) {
				expect(uri, `${record.control}.${record.property}`).toBeUndefined();
			} else {
				expect(uri, `${record.control}.${record.property}`).toMatch(/^data:image\/(x-icon|bmp|png|gif|jpeg);base64,/);
				painted += 1;
			}
		}
		expect(painted).toBeGreaterThan(0);
		// A form picture without the class id in front: MCD's JPEG background.
		const mcd = fixture('Diabetes-prediction-1.0/MCD.frm');
		const mcdHeader = headerOf(mcd);
		const mcdRecords = readFrxRecords(mcdHeader, fs.readFileSync(mcd.replace(/[.]frm$/i, '.frx')), decode);
		const mcdByOffset = new Map(mcdRecords.map((r) => [r.offset, r.value]));
		const mcdScene = sceneOfFrmHeader(mcdHeader, { formName: 'MCD', frx: (ref) => mcdByOffset.get(ref.offset) });
		expect(mcdScene.form.pictureCss).toContain('data:image/jpeg;base64,');
		const byOffset = new Map(records.map((r) => [r.offset, r.value]));
		const withFrx = sceneOfFrmHeader(header, { formName: 'f', frx: (ref) => byOffset.get(ref.offset) });
		const bare = sceneOfFrmHeader(header, { formName: 'f' });
		const pictured = (controls: readonly SceneControl[]): number =>
			flatten(controls).filter((c) => c.style.includes('background-image:')).length
			+ (withFrx.form.pictureCss.includes('background-image:') ? 1 : 0);
		expect(pictured(withFrx.controls)).toBeGreaterThanOrEqual(pictured(bare.controls));
		expect(bare.form.pictureCss).toBe('');
	});
});

describe('listFrmProperties', () => {
	it('lists the form under the empty key and every control by its canvas name', () => {
		const header = headerOf(fixture('polyworks/frmInfo.frm'));
		const props = listFrmProperties(header, { formName: 'frmInfo' });
		expect(props['']).toBeDefined();
		expect(props[''].kind).toBe('Form');
		expect(props[''].rows.find((r) => r.prop === 'Caption')).toBeDefined();
		expect(props[''].rows.find((r) => r.prop === 'Name')?.value).toBe(header.form.name);
		const scene = sceneOfFrmHeader(header, { formName: 'frmInfo' });
		for (const control of flatten(scene.controls)) {
			expect(props[control.name], control.name).toBeDefined();
			const rows = props[control.name].rows;
			expect(rows[0]).toEqual({ prop: 'Name', value: control.name.replace(/[(].*$/, '') });
			if (control.kind !== 'Line' && control.kind !== 'Timer') {
				expect(rows.map((r) => r.prop)).toEqual(expect.arrayContaining(['Left', 'Top', 'Width', 'Height']));
			}
		}
		expect(Object.keys(props).some((k) => k.startsWith('mnu'))).toBe(false);
	});

	it('lists the design-time vocabulary of the kind beside what the header states, blank until set', () => {
		const header = headerOf(fixture('RunAsTrustedInstaller/Form1.frm'));
		const props = listFrmProperties(header, { formName: 'Form1' });
		const rows = props.Command1.rows;
		expect(rows.find((r) => r.prop === 'Caption')?.value).toBe('Run As TrustedInstaller');
		expect(rows.find((r) => r.prop === 'Enabled')?.value).toBe('False');
		expect(rows.find((r) => r.prop === 'Default')?.value).toBe('True');
		expect(rows.find((r) => r.prop === 'BackColor')).toEqual({ prop: 'BackColor', value: '' });
		expect(rows.find((r) => r.prop === 'Font.Name')).toEqual({ prop: 'Font.Name', value: '' });
		expect(rows.some((r) => r.prop === 'hWnd' || r.prop === 'Parent')).toBe(false);
		// A picture the designer will not write is never offered blank; a
		// header that states one still shows it, as its sidecar reference.
		const form = props[''].rows;
		expect(form.some((r) => r.prop === 'Picture' && r.value === '')).toBe(false);
		const withIcon = listFrmProperties(headerOf(fixture('audiostation/Form_OpenDialog.frm')), { formName: 'f' })[''].rows;
		expect(withIcon.find((r) => r.prop === 'Icon')?.value).toMatch(/Form_OpenDialog\.frx:0000/);
		for (const target of Object.values(props)) {
			const names = target.rows.map((r) => r.prop.toLowerCase());
			expect(new Set(names).size).toBe(names.length);
		}
		const width = header.form.members.find((m) => m.kind === 'property' && m.key === 'ClientWidth');
		expect(props[''].rows.find((r) => r.prop === 'Width')?.value).toBe(twipsToPt(Number(width?.kind === 'property' ? width.value : 0)));
		expect(props[''].rows.some((r) => r.prop === 'ClientWidth')).toBe(false);
	});

	it('offers the model\'s own constants for enum properties, True/False for Booleans, and a text field otherwise', () => {
		const { enums, bools } = vb6PaneVocabulary();
		expect(enums['Form.BorderStyle']).toEqual(expect.arrayContaining([['2', 'vbSizable'], ['3', 'vbFixedDialog']]));
		expect(enums['Form.BorderStyle'].map(([v]) => v)).toEqual(['0', '1', '2', '3', '4', '5']);
		expect(enums['CheckBox.Value']).toEqual(expect.arrayContaining([['2', 'vbGrayed']]));
		expect(enums['Form.StartUpPosition'].length).toBe(4);
		// The model declares the type but holds no constants for it: no dropdown.
		expect(enums['TextBox.BorderStyle']).toBeUndefined();
		expect(bools).toEqual(expect.arrayContaining(['Form.KeyPreview', 'TextBox.Locked', 'Label.AutoSize']));
		expect(bools).not.toContain('CheckBox.Value');
		const header = headerOf(fixture('RunAsTrustedInstaller/Form1.frm'));
		const html = renderFormSceneHtml(sceneOfFrmHeader(header, { formName: 'Form1' }), { formName: 'Form1' });
		expect(html).toContain('"Form.BorderStyle"');
		expect(html).toContain('"Form.KeyPreview"');
		// The pane's own MSForms tables never answer a VB6 row by bare name.
		expect(html).toContain('const ENUM_FALLBACK = false;');
		// A Line or Shape has no events; the table says so and the canvas passes it on.
		expect(html).toContain('"Line":""');
	});

	it('shows a color the way the pane\'s swatch and picker speak it', () => {
		const rgb = frmDisplayValue({ kind: 'property', key: 'BackColor', value: '&H004A3C31&' });
		expect(rgb).toBe(formatOleColor(parseOleColor('&H004A3C31&')!));
		expect(rgb).toMatch(/^#[0-9a-f]{6}$/i);
		expect(frmDisplayValue({ kind: 'property', key: 'BackColor', value: '&H8000000F&' })).toBe(formatOleColor(0x8000000f));
		expect(frmDisplayValue({ kind: 'property', key: 'Caption', value: '"&H00&"' })).toBe('&H00&');
		const header = headerOf(fixture('polyworks/frmInfo.frm'));
		const rows = listFrmProperties(header, { formName: 'frmInfo' })[''].rows;
		expect(rows.find((r) => r.prop === 'BackColor')?.value).toBe(rgb);
	});

	it('flattens a Font group into Font.* rows, unquoted', () => {
		const withFont = formFiles().find((f) => /BeginProperty Font/.test(fs.readFileSync(f, 'latin1')));
		expect(withFont).toBeDefined();
		const header = headerOf(withFont!);
		const props = listFrmProperties(header, { formName: header.form.name });
		const rows = Object.values(props).flatMap((p) => p.rows);
		const fontName = rows.find((r) => r.prop === 'Font.Name');
		expect(fontName).toBeDefined();
		expect(fontName!.value.startsWith('"')).toBe(false);
	});
});
