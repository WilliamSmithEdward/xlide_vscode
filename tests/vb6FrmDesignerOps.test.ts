import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { FrmHeaderError, frmControls, frmProperty, parseFrmHeader } from '../src/vba/vb6/frmHeader';
import type { FrmControl, FrmHeader } from '../src/vba/vb6/frmHeader';
import {
	applyFrmDesignerOp, twipsOfPt, vb6FormHandlerPrefix, vb6HeaderEndOf,
} from '../src/vba/vb6/frmDesignerOps';
import type { FrmDesignerOp } from '../src/vba/vb6/frmDesignerOps';
import { vb6CanvasKind, vb6ControlName } from '../src/vba/vb6/frmScene';
import { applyVb6FormDesignerOp, readVb6FormPreview } from '../src/vba/projectService';

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

	it('keeps a multi-line string in the sidecar, through the store the caller supplies', () => {
		const text = read(FORM1);
		expect(() => apply(text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' })).toThrow(/\.frx/);
		const stored: string[] = [];
		const result = applyFrmDesignerOp(text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' }, {
			storeString: (v) => { stored.push(v); return { file: 'Form1.frx', offset: 0x2a, long: false }; },
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
		const after = apply(text, { kind: 'tabOrder', container: '', names: ['Command1', 'Text2', 'Text1'] });
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
	it('finds the header end at the form block\'s own End, and the handler prefix by designer class', () => {
		const text = read(FORM1);
		expect(vb6HeaderEndOf(text)).toBe(headerOf(text).endOffset);
		for (const file of fixtureForms()) {
			const t = read(file);
			expect(vb6HeaderEndOf(t), file).toBe(headerOf(t).endOffset);
		}
		expect(vb6FormHandlerPrefix(text)).toBe('Form');
		expect(vb6FormHandlerPrefix(read(path.join(ROOT, 'audiostation', 'Hyperlink.ctl')))).toBe('UserControl');
		expect(vb6FormHandlerPrefix('VERSION 5.00\r\nBegin VB.MDIForm MDIForm1 \r\nEnd\r\n')).toBe('MDIForm');
		expect(vb6HeaderEndOf('no header here')).toBe('no header here'.length);
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

	it('appends a multi-line Text to the .frx and points the header at it, then renders it', () => {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vb6-designer-'));
		const frm = path.join(dir, 'Form1.frm');
		fs.copyFileSync(FORM1, frm);
		const frx = path.join(dir, 'Form1.frx');
		if (fs.existsSync(FORM1.replace(/frm$/, 'frx'))) { fs.copyFileSync(FORM1.replace(/frm$/, 'frx'), frx); }
		const sizeBefore = fs.existsSync(frx) ? fs.statSync(frx).size : 0;
		const text = read(frm);
		const result = applyVb6FormDesignerOp(frm, text, { kind: 'setProp', name: 'Text1', prop: 'Text', value: 'one\r\ntwo' });
		expect(fs.statSync(frx).size).toBe(sizeBefore + 1 + 'one\r\ntwo'.length);
		const ref = frmProperty(control(result.text, 'Text1'), 'Text')!;
		expect(ref.frx).toEqual({ file: 'Form1.frx', offset: sizeBefore, long: false });
		fs.writeFileSync(frm, result.text, 'latin1');
		const { html } = readVb6FormPreview(frm, result.text, 'Text1');
		expect(html).toContain('data-name="Text1"');
		expect(html).toContain('one\ntwo'.replace('\n', '\r\n').replace(/\r\n/, '\r\n'));
	});

	it('renders a form without a sidecar and refuses text without a header', () => {
		const text = read(FORM1);
		const { html } = readVb6FormPreview(FORM1, text);
		for (const name of names(text)) { expect(html).toContain(`data-name="${name}"`); }
		expect(() => readVb6FormPreview(FORM1, 'Option Explicit')).toThrow(/designer header/);
	});
});
