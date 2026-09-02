import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { FrmHeaderError, frmControls, frmProperty, parseFrmHeader } from '../src/vba/vb6/frmHeader';
import { readFrxRecords } from '../src/vba/vb6/frx';
import type { FrmControl, FrmHeader } from '../src/vba/vb6/frmHeader';
import {
	applyFrmDesignerOp, frmDesignerOpOfGesture, twipsOfPt, vb6FormHandlerPrefix, vb6HeaderEndOf, vb6PendingRecordsToWrite,
} from '../src/vba/vb6/frmDesignerOps';
import type { FrmDesignerOp } from '../src/vba/vb6/frmDesignerOps';
import { vb6CanvasKind, vb6ControlName } from '../src/vba/vb6/frmScene';
import { appendVb6Sidecar, applyVb6FormDesignerOp, readVb6FormPreview, vb6SidecarFileFor } from '../src/vba/projectService';

// Designer gestures as header rewrites, measured on the fixture forms: a
// gesture that changes nothing returns the file's own bytes, a real gesture
// is a header-only diff, and every result parses back.

const ROOT = path.join(__dirname, 'fixtures', 'vb6');
const FORM1 = path.join(ROOT, 'RunAsTrustedInstaller', 'Form1.frm');
const FRM_INFO = path.join(ROOT, 'polyworks', 'frmInfo.frm');
const FRM_PREFERENCES = path.join(ROOT, 'polyworks', 'frmPreferences.frm');
const MCD = path.join(ROOT, 'Diabetes-prediction-1.0', 'MCD.frm');

function fixtureForms(): string[] {
	const out: string[] = [];
	for (const dir of fs.readdirSync(ROOT)) {
		const full = path.join(ROOT, dir);
		if (!fs.statSync(full).isDirectory()) { continue; }
		for (const file of fs.readdirSync(full)) {
			if (/[.](frm|ctl)$/i.test(file)) { out.push(path.join(full, file)); }
		}
	}
	return out.sort();
}

function read(file: string): string {
	return fs.readFileSync(file, 'latin1');
}

function headerOf(text: string): FrmHeader {
	const header = parseFrmHeader(text);
	if (!header) { throw new Error('no header'); }
	return header;
}

function codeOf(text: string): string {
	return text.slice(headerOf(text).endOffset);
}

function control(text: string, canvasName: string): FrmControl {
	const found = frmControls(headerOf(text)).find((c) => vb6ControlName(c).toLowerCase() === canvasName.toLowerCase());
	if (!found) { throw new Error(`no control ${canvasName}`); }
	return found;
}

function value(text: string, canvasName: string, key: string): string | undefined {
	return frmProperty(control(text, canvasName), key)?.value;
}

function num(text: string, canvasName: string, key: string): number {
	return Number(value(text, canvasName, key));
}

function apply(text: string, op: FrmDesignerOp): string {
	return applyFrmDesignerOp(text, op).text;
}

/** The canvas names of every control, in file order. */
function names(text: string): string[] {
	return frmControls(headerOf(text)).map(vb6ControlName);
}

describe('a gesture that changes nothing', () => {
	for (const file of fixtureForms()) {
		const rel = path.relative(ROOT, file).split(path.sep).join('/');
		it(`returns the bytes of ${rel}`, () => {
			const text = read(file);
			const header = headerOf(text);
			const plain = frmControls(header).find((c) => {
				const kind = vb6CanvasKind(c.progId);
				return c.progId !== 'VB.Menu' && kind !== 'Line' && kind !== 'Timer'
					&& frmProperty(c, 'Left') && frmProperty(c, 'Top') && !frmProperty(c, 'Left')!.frx;
			});
			if (plain) {
				const name = vb6ControlName(plain);
				const left = Number(frmProperty(plain, 'Left')!.value) / 20;
				const top = Number(frmProperty(plain, 'Top')!.value) / 20;
				expect(apply(text, { kind: 'geometry', name, left, top })).toBe(text);
				// Sending a control to the end it already sits at moves nothing.
				const siblings = header.form.children.map(vb6ControlName);
				const at = siblings.indexOf(name);
				if (at === siblings.length - 1) {
					expect(apply(text, { kind: 'zOrder', name, toFront: true })).toBe(text);
				} else if (at === 0) {
					expect(apply(text, { kind: 'zOrder', name, toFront: false })).toBe(text);
				}
			}
			const width = frmProperty(header.form, 'ClientWidth');
			const height = frmProperty(header.form, 'ClientHeight');
			if (width && height) {
				expect(apply(text, { kind: 'formSize', width: Number(width.value) / 20, height: Number(height.value) / 20 })).toBe(text);
			}
		});
	}
});

describe('a real gesture', () => {
	for (const file of fixtureForms()) {
		const rel = path.relative(ROOT, file).split(path.sep).join('/');
		it(`rewrites only the header of ${rel}`, () => {
			const text = read(file);
			const header = headerOf(text);
			const plain = frmControls(header).find((c) => c.progId !== 'VB.Menu' && vb6CanvasKind(c.progId) !== 'Line'
				&& frmProperty(c, 'Left') && !frmProperty(c, 'Left')!.frx);
			if (!plain) { return; }
			const name = vb6ControlName(plain);
			const before = Number(frmProperty(plain, 'Left')!.value);
			const after = apply(text, { kind: 'geometry', name, left: before / 20 + 1 });
			expect(after).not.toBe(text);
			expect(codeOf(after)).toBe(codeOf(text));
			expect(num(after, name, 'Left')).toBe(before + 20);
			// Back where it was: the file's own bytes again.
			expect(apply(after, { kind: 'geometry', name, left: before / 20 })).toBe(text);
		});
	}
});

describe('geometry', () => {
	it('moves and sizes a control in twips, twenty to the point', () => {
		const text = read(FORM1);
		const after = apply(text, { kind: 'geometry', name: 'Command1', left: 10.5, top: 20, width: 100, height: 25 });
		expect(value(after, 'Command1', 'Left')).toBe('210');
		expect(value(after, 'Command1', 'Top')).toBe('400');
		expect(value(after, 'Command1', 'Width')).toBe('2000');
		expect(value(after, 'Command1', 'Height')).toBe('500');
		expect(twipsOfPt(0.05)).toBe(1);
	});

	it('moves a batch, and a Timer takes only a position', () => {
		const text = read(FORM1);
		const withTimer = apply(text, { kind: 'add', container: '', controlKind: 'Timer', left: 5, top: 5 });
		const after = apply(withTimer, {
			kind: 'geometryBatch',
			items: [{ name: 'Text1', left: 1, top: 2 }, { name: 'Timer1', left: 3, top: 4, width: 50, height: 50 }],
		});
		expect(value(after, 'Text1', 'Left')).toBe('20');
		expect(value(after, 'Text1', 'Top')).toBe('40');
		expect(value(after, 'Timer1', 'Left')).toBe('60');
		expect(value(after, 'Timer1', 'Width')).toBeUndefined();
	});

	it('moves a Line by its two points and keeps which end is which', () => {
		const text = read(MCD);
		const line = frmControls(headerOf(text)).find((c) => c.progId === 'VB.Line')!;
		const name = vb6ControlName(line);
		const x1 = num(text, name, 'X1');
		const x2 = num(text, name, 'X2');
		const y1 = num(text, name, 'Y1');
		const y2 = num(text, name, 'Y2');
		const left = Math.min(x1, x2);
		const top = Math.min(y1, y2);
		const after = apply(text, { kind: 'geometry', name, left: left / 20 + 10, top: top / 20 + 5 });
		expect(num(after, name, 'X1') - x1).toBe(200);
		expect(num(after, name, 'X2') - x2).toBe(200);
		expect(num(after, name, 'Y1') - y1).toBe(100);
		expect(num(after, name, 'Y2') - y2).toBe(100);
		// The canvas box of a flat line is one point tall; that is not a length.
		if (y1 === y2) {
			const resized = apply(text, { kind: 'geometry', name, width: Math.abs(x2 - x1) / 20 + 10, height: 1 });
			expect(num(resized, name, 'Y2')).toBe(y2);
			expect(Math.abs(num(resized, name, 'X2') - num(resized, name, 'X1'))).toBe(Math.abs(x2 - x1) + 200);
		}
	});

	it('sizes the form through its client size and its twip scale', () => {
		const text = read(FORM1);
		const after = apply(text, { kind: 'formSize', width: 300, height: 200 });
		const form = headerOf(after).form;
		expect(frmProperty(form, 'ClientWidth')!.value).toBe('6000');
		expect(frmProperty(form, 'ClientHeight')!.value).toBe('4000');
		if (frmProperty(form, 'ScaleWidth') && (frmProperty(form, 'ScaleMode') === undefined)) {
			expect(frmProperty(form, 'ScaleWidth')!.value).toBe('6000');
			expect(frmProperty(form, 'ScaleHeight')!.value).toBe('4000');
		}
		expect(codeOf(after)).toBe(codeOf(text));
	});
});

describe('add', () => {
	it('writes a CommandButton the way the designer does: fresh name, sorted members, next TabIndex', () => {
		const text = read(FORM1);
		const result = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'CommandButton', left: 10, top: 20 });
		expect(result.newName).toBe('Command2');
		const added = control(result.text, 'Command2');
		expect(added.progId).toBe('VB.CommandButton');
		const keys = added.members.map((m) => (m.kind === 'property' ? m.key : m.name));
		expect(keys).toEqual(['Caption', 'Height', 'Left', 'TabIndex', 'Top', 'Width']);
		expect(frmProperty(added, 'Caption')!.value).toBe('"Command2"');
		expect(frmProperty(added, 'Left')!.value).toBe('200');
		expect(frmProperty(added, 'Top')!.value).toBe('400');
		const tabs = frmControls(headerOf(text)).map((c) => Number(frmProperty(c, 'TabIndex')?.value ?? -1));
		expect(Number(frmProperty(added, 'TabIndex')!.value)).toBe(Math.max(...tabs) + 1);
		expect(result.text).toContain('   Begin VB.CommandButton Command2 \r\n      Caption         =   "Command2"\r\n');
		expect(codeOf(result.text)).toBe(codeOf(text));
		// The next one counts on.
		expect(applyFrmDesignerOp(result.text, { kind: 'add', container: '', controlKind: 'CommandButton', left: 0, top: 0 }).newName).toBe('Command3');
	});

	it('names each kind by its base and writes its own properties', () => {
		const text = read(FORM1);
		const textBox = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'TextBox', left: 0, top: 0 });
		expect(textBox.newName).toBe('Text3');
		expect(value(textBox.text, 'Text3', 'Text')).toBe('"Text3"');
		const line = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'Line', left: 10, top: 10 });
		expect(line.newName).toBe('Line1');
		expect(control(line.text, 'Line1').members.map((m) => (m.kind === 'property' ? m.key : m.name))).toEqual(['X1', 'X2', 'Y1', 'Y2']);
		expect(value(line.text, 'Line1', 'X2')).toBe(String(200 + 1215));
		const picture = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'PictureBox', left: 0, top: 0 });
		expect(value(picture.text, 'Picture1', 'ScaleWidth')).toBe('1155');
		const timer = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'Timer', left: 0, top: 0 });
		expect(control(timer.text, 'Timer1').members.map((m) => (m.kind === 'property' ? m.key : m.name))).toEqual(['Left', 'Top']);
		const scroll = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'VScrollBar', left: 0, top: 0 });
		expect(control(scroll.text, 'VScroll1').progId).toBe('VB.VScrollBar');
	});

	it('nests a control added inside a PictureBox, and refuses a container that cannot hold one', () => {
		const text = read(FRM_INFO);
		const result = applyFrmDesignerOp(text, { kind: 'add', container: 'picTitle', controlKind: 'Label', left: 1, top: 1 });
		const box = control(result.text, 'picTitle');
		expect(box.children.some((c) => c.name === result.newName)).toBe(true);
		const form1 = read(FORM1);
		expect(() => apply(form1, { kind: 'add', container: 'Text1', controlKind: 'Label', left: 0, top: 0 })).toThrow(/cannot hold/);
		expect(() => apply(form1, { kind: 'add', container: '', controlKind: 'Bogus', left: 0, top: 0 })).toThrow(/toolbox/);
		expect(() => apply(form1, { kind: 'add', container: 'Nope', controlKind: 'Label', left: 0, top: 0 })).toThrow(/No control named Nope/);
	});
});

describe('setProp', () => {
	it('spells strings, booleans, colors, and numbers as the designer does', () => {
		const text = read(FORM1);
		let after = apply(text, { kind: 'setProp', name: 'Command1', prop: 'Caption', value: 'Say "Hi"' });
		expect(after).toContain('      Caption         =   "Say ""Hi"""\r\n');
		after = apply(after, { kind: 'setProp', name: 'Command1', prop: 'Enabled', value: 'False' });
		expect(after).toContain("      Enabled         =   0   'False\r\n");
		after = apply(after, { kind: 'setProp', name: 'Command1', prop: 'Enabled', value: 'True' });
		expect(after).toContain("      Enabled         =   -1  'True\r\n");
		after = apply(after, { kind: 'setProp', name: 'Command1', prop: 'BackColor', value: '#0000ff' });
		expect(value(after, 'Command1', 'BackColor')).toBe('&H00FF0000&');
		after = apply(after, { kind: 'setProp', name: 'Command1', prop: 'TabIndex', value: '7' });
		expect(value(after, 'Command1', 'TabIndex')).toBe('7');
		// Members stay in name order.
		const keys = control(after, 'Command1').members.map((m) => (m.kind === 'property' ? m.key : m.name).toLowerCase());
		expect([...keys].sort()).toEqual(keys);
		expect(codeOf(after)).toBe(codeOf(text));
	});

	it('writes the designer\'s gloss after a measured enum value, and a bare number otherwise', () => {
		const text = read(FORM1);
		const dialog = apply(text, { kind: 'setProp', name: '', prop: 'BorderStyle', value: '3' });
		expect(dialog).toContain("   BorderStyle     =   3  'Fixed Dialog\r\n");
		const sizable = apply(text, { kind: 'setProp', name: '', prop: 'BorderStyle', value: '2' });
		expect(sizable).toContain('   BorderStyle     =   2\r\n');
		const centered = apply(text, { kind: 'setProp', name: 'Text1', prop: 'Alignment', value: '2' });
		expect(centered).toContain("      Alignment       =   2  'Center\r\n");
		// The value a line already holds is left as written, gloss and all.
		const again = apply(dialog, { kind: 'setProp', name: '', prop: 'BorderStyle', value: '3' });
		expect(again).toBe(dialog);
	});

	it('sets geometry in points, the form size through Width and Height', () => {
		const text = read(FORM1);
		const after = apply(text, { kind: 'setProp', name: 'Text1', prop: 'Width', value: '100' });
		expect(value(after, 'Text1', 'Width')).toBe('2000');
		const form = apply(text, { kind: 'setProp', name: '', prop: 'Width', value: '300' });
		expect(frmProperty(headerOf(form).form, 'ClientWidth')!.value).toBe('6000');
	});

	it('writes a Font group in the designer\'s fixed order, creating one from the defaults', () => {
		const text = read(FORM1);
		const after = apply(text, { kind: 'setProp', name: 'Command1', prop: 'Font.Bold', value: 'True' });
		const group = control(after, 'Command1').members.find((m) => m.kind === 'group' && m.name === 'Font');
		expect(group).toBeDefined();
		if (group?.kind !== 'group') { throw new Error('no group'); }
		expect(group.members.map((m) => (m.kind === 'property' ? m.key : m.name))).toEqual(
			['Name', 'Size', 'Charset', 'Weight', 'Underline', 'Italic', 'Strikethrough']);
		expect(group.members.find((m) => m.kind === 'property' && m.key === 'Weight')).toMatchObject({ value: '700' });
		const italic = apply(after, { kind: 'setProp', name: 'Command1', prop: 'Font.Italic', value: 'True' });
		expect(italic).toContain("         Italic          =   -1  'True\r\n");
		expect(() => apply(text, { kind: 'setProp', name: 'Command1', prop: 'Font.Color', value: '1' })).toThrow(/not a font property/);
	});

	it('renames a control, and every element of a control array with it', () => {
		const text = read(FORM1);
		const result = applyFrmDesignerOp(text, { kind: 'setProp', name: 'Command1', prop: 'Name', value: 'btnGo' });
		expect(result.newName).toBe('btnGo');
		expect(result.text).toContain('   Begin VB.CommandButton btnGo \r\n');
		expect(names(result.text)).not.toContain('Command1');
		expect(() => apply(text, { kind: 'setProp', name: 'Command1', prop: 'Name', value: 'Text1' })).toThrow(/already exists/);
		expect(() => apply(text, { kind: 'setProp', name: 'Command1', prop: 'Name', value: '1bad' })).toThrow(/not a valid/);
		expect(() => apply(text, { kind: 'setProp', name: '', prop: 'Name', value: 'frmOther' })).toThrow(/module/);

		const prefs = read(FRM_PREFERENCES);
		const elements = names(prefs).filter((n) => /^fraPref\(\d+\)$/.test(n));
		expect(elements.length).toBeGreaterThan(1);
		const renamed = applyFrmDesignerOp(prefs, { kind: 'setProp', name: elements[0], prop: 'Name', value: 'shpPref' });
		expect(renamed.newName).toBe(elements[0].replace('fraPref', 'shpPref'));
		expect(names(renamed.text).filter((n) => /^shpPref\(\d+\)$/.test(n)).length).toBe(elements.length);
		expect(names(renamed.text).some((n) => n.startsWith('fraPref'))).toBe(false);
	});

	it('makes a control an array element through Index, and a plain control again', () => {
		const text = read(FORM1);
		const indexed = applyFrmDesignerOp(text, { kind: 'setProp', name: 'Label1', prop: 'Index', value: '3' });
		expect(indexed.newName).toBe('Label1(3)');
		expect(indexed.text).toContain('      Index           =   3\r\n');
		const plain = applyFrmDesignerOp(indexed.text, { kind: 'setProp', name: 'Label1(3)', prop: 'Index', value: '' });
		expect(plain.newName).toBe('Label1');
		expect(plain.text).toBe(text);
		const prefs = read(FRM_PREFERENCES);
		const first = names(prefs).find((n) => /^fraPref\(\d+\)$/.test(n))!;
		expect(() => apply(prefs, { kind: 'setProp', name: first, prop: 'Index', value: '' })).toThrow(/control array/);
	});

	it('spells a value by the type the model declares, not by how the value reads', () => {
		const text = read(FORM1);
		expect(apply(text, { kind: 'setProp', name: 'Label1', prop: 'Caption', value: 'True' })).toContain('      Caption         =   "True"\r\n');
		expect(apply(text, { kind: 'setProp', name: 'Label1', prop: 'Caption', value: '42' })).toContain('      Caption         =   "42"\r\n');
		expect(apply(text, { kind: 'setProp', name: 'Text1', prop: 'Tag', value: '0' })).toContain('      Tag             =   "0"\r\n');
		expect(apply(text, { kind: 'setProp', name: 'Command1', prop: 'Enabled', value: '1' })).toContain("      Enabled         =   -1  'True\r\n");
		expect(() => apply(text, { kind: 'setProp', name: 'Command1', prop: 'Enabled', value: 'maybe' })).toThrow(/True or False/);
		expect(apply(text, { kind: 'setProp', name: 'Label1', prop: 'ForeColor', value: '#ff0000' })).toContain('      ForeColor       =   &H000000FF&\r\n');
	});

	it('refuses the sidecar kinds it does not write: pictures and lists', () => {
		const text = read(FORM1);
		expect(() => apply(text, { kind: 'setProp', name: 'Command1', prop: 'Picture', value: 'x' })).toThrow(/picture.*sidecar/);
		expect(() => apply(text, { kind: 'setProp', name: '', prop: 'Icon', value: 'x' })).toThrow(/picture.*sidecar/);
		expect(() => apply(text, { kind: 'setProp', name: 'Text1', prop: 'List', value: 'x' })).toThrow(/sidecar/);
		expect(() => apply(text, { kind: 'setProp', name: 'Label1', prop: 'MouseIcon', value: '' })).toThrow(/picture/);
	});

	it('keeps a multi-line string in the sidecar, through the store the caller supplies', () => {
		const text = read(FORM1);
		expect(() => apply(text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' })).toThrow(/\.frx/);
		const stored: string[] = [];
		const result = applyFrmDesignerOp(text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' }, {
			storeString: (v, header) => { stored.push(v); expect(header.form.name).toBe('Form1'); return { file: 'Form1.frx', offset: 0x2a, long: false }; },
		});
		expect(stored).toEqual(['one\r\ntwo']);
		expect(value(result.text, 'Text1', 'Text')).toBe('"Form1.frx":002A');
		const long = applyFrmDesignerOp(text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'a\nb' }, {
			storeString: () => ({ file: 'Form1.frx', offset: 0x1000, long: true }),
		});
		expect(value(long.text, 'Text1', 'Text')).toBe('$"Form1.frx":1000');
	});
});

describe('remove, reparent, order, duplicate', () => {
	it('removes a container with its children, outermost only, all or nothing', () => {
		const text = read(FRM_INFO);
		const box = control(text, 'picTitle');
		expect(box.children.length).toBeGreaterThan(0);
		const child = vb6ControlName(box.children[0]);
		const result = applyFrmDesignerOp(text, { kind: 'removeMany', names: [child, 'picTitle'] });
		expect(result.removed).toEqual(['picTitle']);
		expect(names(result.text)).not.toContain('picTitle');
		expect(names(result.text)).not.toContain(child);
		expect(codeOf(result.text)).toBe(codeOf(text));
		expect(() => apply(text, { kind: 'removeMany', names: ['picTitle', 'Nope'] })).toThrow(/No control named Nope/);
		const one = applyFrmDesignerOp(text, { kind: 'remove', name: child });
		expect(one.removed).toEqual([child]);
		expect(names(one.text)).toContain('picTitle');
	});

	it('reparents into a PictureBox and refuses a control inside itself', () => {
		const text = read(FRM_INFO);
		const header = headerOf(text);
		const top = header.form.children.find((c) => c.progId !== 'VB.Menu' && c.name !== 'picTitle' && c.children.length === 0)!;
		const name = vb6ControlName(top);
		const after = apply(text, { kind: 'reparent', name, container: 'picTitle', left: 2, top: 3 });
		const box = control(after, 'picTitle');
		expect(box.children.some((c) => vb6ControlName(c) === name)).toBe(true);
		expect(headerOf(after).form.children.some((c) => vb6ControlName(c) === name)).toBe(false);
		expect(value(after, name, 'Left')).toBe('40');
		expect(value(after, name, 'Top')).toBe('60');
		expect(() => apply(text, { kind: 'reparent', name: 'picTitle', container: 'picTitle', left: 0, top: 0 })).toThrow(/inside itself/);
	});

	it('orders siblings: front is last in the file, back is first', () => {
		const text = read(FORM1);
		const front = headerOf(apply(text, { kind: 'zOrder', name: 'Text2', toFront: true })).form.children;
		expect(front[front.length - 1].name).toBe('Text2');
		const back = headerOf(apply(text, { kind: 'zOrder', name: 'Label1', toFront: false })).form.children;
		expect(back[0].name).toBe('Label1');
	});

	it('reorders tab stops within the indices the controls already hold', () => {
		const text = read(FORM1);
		const before = ['Text1', 'Text2', 'Command1'].map((n) => num(text, n, 'TabIndex')).sort((a, b) => a - b);
		const after = apply(text, { kind: 'tabOrder', names: ['Command1', 'Text2', 'Text1'] });
		expect(num(after, 'Command1', 'TabIndex')).toBe(before[0]);
		expect(num(after, 'Text2', 'TabIndex')).toBe(before[1]);
		expect(num(after, 'Text1', 'TabIndex')).toBe(before[2]);
		expect(num(after, 'Label1', 'TabIndex')).toBe(num(text, 'Label1', 'TabIndex'));
	});

	it('duplicates under fresh names, nudged, with fresh tab stops and no Index', () => {
		const text = read(FORM1);
		const result = applyFrmDesignerOp(text, { kind: 'duplicate', names: ['Command1', 'Text1'] });
		expect(result.newNames).toEqual(['Command2', 'Text3']);
		expect(num(result.text, 'Command2', 'Left')).toBe(num(text, 'Command1', 'Left') + 120);
		expect(num(result.text, 'Command2', 'Top')).toBe(num(text, 'Command1', 'Top') + 120);
		expect(value(result.text, 'Command2', 'Caption')).toBe(value(text, 'Command1', 'Caption'));
		const tabs = frmControls(headerOf(result.text)).map((c) => frmProperty(c, 'TabIndex')?.value).filter((v) => v !== undefined);
		expect(new Set(tabs).size).toBe(tabs.length);

		const info = read(FRM_INFO);
		const box = control(info, 'picTitle');
		const copy = applyFrmDesignerOp(info, { kind: 'duplicate', names: ['picTitle'] });
		const cloned = control(copy.text, copy.newNames![0]);
		expect(cloned.children.length).toBe(box.children.length);
		expect(new Set(names(copy.text)).size).toBe(names(copy.text).length);

		const prefs = read(FRM_PREFERENCES);
		const element = names(prefs).find((n) => /^fraPref\(\d+\)$/.test(n))!;
		const shape = applyFrmDesignerOp(prefs, { kind: 'duplicate', names: [element] });
		expect(shape.newNames![0]).not.toMatch(/\(/);
		expect(value(shape.text, shape.newNames![0], 'Index')).toBeUndefined();
	});
});

describe('the document around the header', () => {
	it('finds the header end where the parser does, and the handler prefix by designer class', () => {
		const text = read(FORM1);
		expect(vb6HeaderEndOf(text)).toBe(headerOf(text).endOffset);
		for (const file of fixtureForms()) {
			const t = read(file);
			expect(vb6HeaderEndOf(t), file).toBe(headerOf(t).endOffset);
		}
		expect(vb6FormHandlerPrefix(text)).toBe('Form');
		expect(vb6FormHandlerPrefix(read(path.join(ROOT, 'audiostation', 'Hyperlink.ctl')))).toBe('UserControl');
		expect(vb6FormHandlerPrefix('VERSION 5.00\r\nBegin VB.MDIForm MDIForm1 \r\nEnd\r\n')).toBe('MDIForm');
	});

	it('never puts code inside the header span', () => {
		// A nested End written without its indent: the parser tracks depth, and so does the boundary.
		const nested = 'VERSION 5.00\r\nBegin VB.Form Form1 \r\n   Begin VB.TextBox Text1 \r\n      Left = 0\r\nEnd\r\nEnd\r\nAttribute VB_Name = "Form1"\r\n';
		expect(vb6HeaderEndOf(nested)).toBe(headerOf(nested).endOffset);
		expect(nested.slice(vb6HeaderEndOf(nested)!)).toBe('Attribute VB_Name = "Form1"\r\n');
		// The header's own End is gone: `End Sub` in the code is not a boundary.
		const broken = 'VERSION 5.00\r\nBegin VB.Form Form1 \r\n   Caption = "x"\r\nAttribute VB_Name = "Form1"\r\nPrivate Sub Form_Load()\r\nEnd Sub\r\n';
		expect(vb6HeaderEndOf(broken)).toBeUndefined();
		// ...unless the document still opens with the text a pane placed last, whose length is the boundary.
		const placed = 'VERSION 5.00\r\nBegin VB.Form Form1 \r\n   Caption = "x"\r\n';
		expect(vb6HeaderEndOf(broken, placed)).toBe(placed.length);
		expect(vb6HeaderEndOf(broken, 'VERSION 5.00\r\nBegin VB.Form Other \r\n')).toBeUndefined();
		// An unindented End line with nothing after it on the line is the fallback boundary.
		const bare = 'VERSION 5.00\r\nBegin VB.Form Form1 \r\n   Begin VB.TextBox T \r\nEnd\r\nOption Explicit\r\n';
		expect(vb6HeaderEndOf(bare)).toBe(bare.indexOf('Option'));
		expect(vb6HeaderEndOf('no header here')).toBeUndefined();
	});

	it('bounds a header exactly whatever line endings it and the code use', () => {
		const mixed = 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption = "x"\nEnd\n' + 'Attribute VB_Name = "Form1"\r\nOption Explicit\r\n';
		const header = headerOf(mixed);
		expect(mixed.slice(header.endOffset)).toBe('Attribute VB_Name = "Form1"\r\nOption Explicit\r\n');
		const moved = applyFrmDesignerOp(mixed, { kind: 'formSize', width: 100, height: 100 });
		expect(moved.text.slice(moved.headerEnd)).toBe('Attribute VB_Name = "Form1"\r\nOption Explicit\r\n');
		expect(moved.oldHeaderEnd).toBe(header.endOffset);
	});

	it('reports where the header ends before and after a gesture', () => {
		const text = read(FORM1);
		const same = applyFrmDesignerOp(text, { kind: 'formSize', width: Number(value(text, 'Text1', 'Width')) / 20 + 0, height: 1 });
		expect(same.oldHeaderEnd).toBe(headerOf(text).endOffset);
		const added = applyFrmDesignerOp(text, { kind: 'add', container: '', controlKind: 'Label', left: 0, top: 0 });
		expect(added.oldHeaderEnd).toBe(headerOf(text).endOffset);
		expect(added.text.slice(added.headerEnd)).toBe(codeOf(text));
		expect(added.text.slice(0, added.headerEnd)).toBe(codeOf(added.text) === codeOf(text) ? added.text.slice(0, headerOf(added.text).endOffset) : '');
		const untouched = applyFrmDesignerOp(text, { kind: 'zOrder', name: 'Label1', toFront: true });
		expect(untouched.text).toBe(text);
		expect(untouched.headerEnd).toBe(headerOf(text).endOffset);
	});

	it('refuses text that is not a form', () => {
		expect(() => applyFrmDesignerOp('Option Explicit\r\n', { kind: 'formSize', width: 1, height: 1 })).toThrow(FrmHeaderError);
	});
});

describe('the engine: gestures over a form file with its sidecar', () => {
	let dir: string | undefined;
	afterEach(() => {
		if (dir) { fs.rmSync(dir, { recursive: true, force: true }); }
		dir = undefined;
	});

	it('places a multi-line Text as a pending sidecar record, renders it pending, and writes it at save', () => {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vb6-designer-'));
		const frm = path.join(dir, 'Form1.frm');
		fs.copyFileSync(FORM1, frm);
		const frx = path.join(dir, 'Form1.frx');
		fs.copyFileSync(FORM1.replace(/frm$/, 'frx'), frx);
		const sizeBefore = fs.statSync(frx).size;
		const text = read(frm);
		// The gesture: a record placed, nothing written.
		const first = applyVb6FormDesignerOp(frm, text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' });
		expect(fs.statSync(frx).size).toBe(sizeBefore);
		expect(first.sidecar).toMatchObject({ file: 'Form1.frx', base: sizeBefore, offset: sizeBefore });
		const record = Buffer.from(first.sidecar!.record, 'base64');
		expect(record.length).toBe(1 + 'one\r\ntwo'.length);
		expect(record[0]).toBe('one\r\ntwo'.length);
		expect(frmProperty(control(first.text, 'Text1'), 'Text')!.frx).toEqual({ file: 'Form1.frx', offset: sizeBefore, long: false });
		// A second record lands after the first, by the pending bytes.
		const second = applyVb6FormDesignerOp(frm, first.text, { kind: 'setProp', name: 'Text2', prop: 'Text', value: 'a\r\nb' }, record.length);
		expect(second.sidecar).toMatchObject({ base: sizeBefore, offset: sizeBefore + record.length });
		const pending = { file: 'Form1.frx', base: sizeBefore, records: [first.sidecar!.record, second.sidecar!.record] };
		// Rendered from disk plus the pending records.
		const preview = readVb6FormPreview(frm, second.text, 'Text1', undefined, pending);
		expect(preview.html).toContain('one\r\ntwo');
		expect(preview.html).toContain('a\r\nb');
		expect(preview.headerEnd).toBe(headerOf(second.text).endOffset);
		// Without them, the references point past the file: blank, not wrong.
		expect(readVb6FormPreview(frm, second.text).html).not.toContain('one\r\ntwo');
		// The save: appended in order; a sidecar that moved on is refused.
		expect(() => appendVb6Sidecar(frm, 'Form1.frx', sizeBefore + 1, pending.records)).toThrow(/changed on disk/);
		expect(fs.statSync(frx).size).toBe(sizeBefore);
		const written = appendVb6Sidecar(frm, 'Form1.frx', sizeBefore, pending.records);
		expect(written.length).toBe(sizeBefore + record.length + Buffer.from(second.sidecar!.record, 'base64').length);
		expect(fs.statSync(frx).size).toBe(written.length);
		fs.writeFileSync(frm, second.text, 'latin1');
		const saved = readVb6FormPreview(frm, second.text, 'Text1');
		expect(saved.html).toContain('one\r\ntwo');
		expect(saved.html).toContain('a\r\nb');
	});

	it('reads every record even when unreferenced bytes follow it in the sidecar', () => {
		for (const rel of ['RunAsTrustedInstaller/Form1', 'audiostation/Form_OpenDialog', 'polyworks/frmPreferences']) {
			const frm = path.join(ROOT, ...`${rel}.frm`.split('/'));
			const header = headerOf(read(frm));
			const blob = fs.readFileSync(frm.replace(/frm$/, 'frx'));
			const decode = (b: Buffer): string => b.toString('latin1');
			const clean = readFrxRecords(header, blob, decode);
			// An undone value's record, left at the end by an earlier save.
			const orphan = Buffer.concat([blob, Buffer.from([5]), Buffer.from('stale')]);
			const withOrphan = readFrxRecords(header, orphan, decode);
			expect(withOrphan.map((r) => r.value.kind), rel).toEqual(clean.map((r) => r.value.kind));
			for (let i = 0; i < clean.length; i++) {
				const a = clean[i].value;
				const b = withOrphan[i].value;
				if (a.kind === 'longString' || a.kind === 'shortString') { expect(b).toEqual(a); }
				if (a.kind === 'picture' && b.kind === 'picture') { expect(b.bytes.equals(a.bytes)).toBe(true); }
				if (a.kind === 'list') { expect(b).toEqual(a); }
			}
		}
	});

	it('answers every canvas gesture with an op and a selection', () => {
		expect(frmDesignerOpOfGesture({ type: 'formResize', width: 10, height: 20 }).op).toEqual({ kind: 'formSize', width: 10, height: 20 });
		expect(frmDesignerOpOfGesture({ type: 'paste', names: ['A'] }).op).toEqual({ kind: 'duplicate', names: ['A'] });
		expect(frmDesignerOpOfGesture({ type: 'tabOrder', container: 'Frame1', names: ['A', 'B'] }).op).toEqual({ kind: 'tabOrder', names: ['A', 'B'] });
		expect(frmDesignerOpOfGesture({ type: 'geometry', name: 'A', left: 1 }).selectAfter({})).toBe('A');
		expect(frmDesignerOpOfGesture({ type: 'add', container: '', controlKind: 'Label', left: 0, top: 0 }).selectAfter({ newName: 'Label4' })).toBe('Label4');
		expect(frmDesignerOpOfGesture({ type: 'setProp', name: 'A', prop: 'Name', value: 'B' }).selectAfter({ newName: 'B' })).toBe('B');
		expect(frmDesignerOpOfGesture({ type: 'setProp', name: 'A', prop: 'Caption', value: 'x' }).selectAfter({})).toBe('A');
		expect(frmDesignerOpOfGesture({ type: 'paste', names: ['A'] }).selectAfter({ newNames: ['A1', 'B1'] })).toBe('A1');
		expect(frmDesignerOpOfGesture({ type: 'formResize', width: 1, height: 1 }).selectAfter({})).toBe('');
		expect(frmDesignerOpOfGesture({ type: 'remove', name: 'A' }).selectAfter({})).toBeUndefined();
		expect(frmDesignerOpOfGesture({ type: 'geometryBatch', anchor: 'B', items: [] }).selectAfter({})).toBe('B');
	});

	it('writes pending records up to the last one the document still references', () => {
		const records = ['AA==', 'AQ==', 'Ag=='];
		const offsets = [0x93, 0x9a, 0xa1];
		const all = 'Text = "Form1.frx":0093\r\nText = "Form1.frx":009A\r\nCaption = $"Form1.frx":00A1\r\n';
		expect(vb6PendingRecordsToWrite('Form1.frx', records, offsets, all)).toEqual(records);
		// The last one undone: not written. The middle one undone: still written, the last one's offset counts on it.
		expect(vb6PendingRecordsToWrite('Form1.frx', records, offsets, 'Text = "Form1.frx":0093\r\nText = "Form1.frx":009A\r\n')).toEqual(['AA==', 'AQ==']);
		expect(vb6PendingRecordsToWrite('Form1.frx', records, offsets, 'Text = "Form1.frx":0093\r\nCaption = $"Form1.frx":00A1\r\n')).toEqual(records);
		expect(vb6PendingRecordsToWrite('Form1.frx', records, offsets, 'Option Explicit\r\n')).toEqual([]);
	});

	it('names the sidecar VB6 pairs with the module: .frx, .ctx, .pgx, .dsx', () => {
		expect(vb6SidecarFileFor('C:\\work\\Form1.frm', undefined)).toBe('Form1.frx');
		expect(vb6SidecarFileFor('/work/Mix.ctl', undefined)).toBe('Mix.ctx');
		expect(vb6SidecarFileFor('/work/Page.pag', undefined)).toBe('Page.pgx');
		expect(vb6SidecarFileFor('/work/Thing.dsr', undefined)).toBe('Thing.dsx');
		const named = headerOf(read(path.join(ROOT, 'audiostation', 'MixSlider.ctl')));
		expect(vb6SidecarFileFor('/elsewhere/Other.ctl', named)).toBe('MixSlider.ctx');
	});

	it('renders a form without a sidecar and refuses text without a header', () => {
		const text = read(FORM1);
		const { html } = readVb6FormPreview(FORM1, text);
		for (const name of names(text)) { expect(html).toContain(`data-name="${name}"`); }
		expect(() => readVb6FormPreview(FORM1, 'Option Explicit')).toThrow(/designer header/);
	});
});
