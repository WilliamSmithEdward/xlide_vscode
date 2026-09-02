// Designer gestures on a VB6 form, as rewrites of its header text.
//
// The form's file is the document: the header at its top is the design and
// the code below it is the code. Every gesture parses the header
// (frmHeader.ts), changes the tree, and prints it back in the designer's
// own layout - which the fixtures prove is byte-identical when nothing
// changed, so a gesture that changes nothing returns the file's own bytes,
// and a real gesture is a header-only diff. The code after the header is
// never touched, and a header the designer's layout does not reproduce
// (hand-edited spacing) is normalized by its first real gesture.
//
// What the header cannot hold inline - a string with line breaks - lives in
// the `.frx` sidecar; the caller supplies the store for it (the engine
// appends to the sidecar) or the gesture is refused with the reason.
//
// Units: the canvas speaks points; the header holds twips, twenty to the point.

import {
	FrmHeaderError, formatFrmHeader, frmProperty, parseFrmHeader,
	type FrmControl, type FrmHeader, type FrmProperty, type FrmPropertyGroup, type FrmPropertyNode, type FrxRef,
} from './frmHeader';
import { parseOleColor } from '../oforms/markup';
import { VB6_ENUM_GLOSSES, frmNumberOf, vb6CanvasKind, vb6ControlName } from './frmScene';

export interface FrmGeometry {
	name: string;
	left?: number;
	top?: number;
	width?: number;
	height?: number;
}

export type FrmDesignerOp =
	| ({ kind: 'geometry' } & FrmGeometry)
	| { kind: 'geometryBatch'; items: FrmGeometry[] }
	| { kind: 'add'; container: string; controlKind: string; left: number; top: number }
	| { kind: 'remove'; name: string }
	| { kind: 'removeMany'; names: string[] }
	| { kind: 'reparent'; name: string; container: string; left: number; top: number }
	| { kind: 'setProp'; name: string; prop: string; value: string }
	| { kind: 'formSize'; width: number; height: number }
	| { kind: 'zOrder'; name: string; toFront: boolean }
	| { kind: 'tabOrder'; container: string; names: string[] }
	| { kind: 'duplicate'; names: string[]; offsetPt?: number };

export interface FrmDesignerOpResult {
	/** The whole document, header rewritten, code as it was. */
	text: string;
	/** What the canvas should select next: a renamed or added control. */
	newName?: string;
	/** The names a paste created, outermost controls only. */
	newNames?: string[];
	/** The names a delete removed, outermost controls only. */
	removed?: string[];
}

export interface FrmDesignerOpOptions {
	/**
	 * Stores a string the header cannot spell inline in the sidecar and
	 * returns the reference to write. Absent, such a value is refused.
	 */
	storeString?: (value: string) => FrxRef;
}

export function twipsOfPt(pt: number): number {
	return Math.round(pt * 20);
}

/** Applies one gesture to a form's text: the header rewritten, the code kept. */
export function applyFrmDesignerOp(text: string, op: FrmDesignerOp, options: FrmDesignerOpOptions = {}): FrmDesignerOpResult {
	const header = parseFrmHeader(text);
	if (!header) {
		throw new FrmHeaderError('The document does not open with a designer header (VERSION 5.00 / Begin VB.Form ... End).');
	}
	const before = formatFrmHeader(header);
	const result = applyToHeader(header, op, options);
	const after = formatFrmHeader(header);
	return { text: after === before ? text : after + text.slice(header.endOffset), ...result };
}

function applyToHeader(header: FrmHeader, op: FrmDesignerOp, options: FrmDesignerOpOptions): Omit<FrmDesignerOpResult, 'text'> {
	switch (op.kind) {
		case 'geometry':
			setGeometry(header, op);
			return {};
		case 'geometryBatch':
			for (const item of op.items) { setGeometry(header, item); }
			return {};
		case 'add':
			return { newName: addControl(header, op.container, op.controlKind, twipsOfPt(op.left), twipsOfPt(op.top)) };
		case 'remove':
			return { removed: removeControls(header, [op.name]) };
		case 'removeMany':
			return { removed: removeControls(header, op.names) };
		case 'reparent':
			reparent(header, op.name, op.container, op.left, op.top);
			return {};
		case 'setProp':
			return { newName: setProp(header, op.name, op.prop, op.value, options) };
		case 'formSize':
			setFormSize(header, twipsOfPt(op.width), twipsOfPt(op.height));
			return {};
		case 'zOrder':
			setZOrder(header, op.name, op.toFront);
			return {};
		case 'tabOrder':
			setTabOrder(header, op.names);
			return {};
		case 'duplicate':
			return { newNames: duplicate(header, op.names, twipsOfPt(op.offsetPt ?? 6)) };
		default:
			throw new Error(`Unknown designer operation: ${(op as { kind: string }).kind}`);
	}
}

// --- the tree ---------------------------------------------------------------

interface Located {
	control: FrmControl;
	parent: FrmControl;
}

function* walk(control: FrmControl): Generator<FrmControl> {
	for (const child of control.children) {
		yield child;
		yield* walk(child);
	}
}

function locate(header: FrmHeader, canvasName: string): Located {
	const wanted = canvasName.toLowerCase();
	const find = (parent: FrmControl): Located | undefined => {
		for (const control of parent.children) {
			if (vb6ControlName(control).toLowerCase() === wanted) { return { control, parent }; }
			const deeper = find(control);
			if (deeper) { return deeper; }
		}
		return undefined;
	};
	const found = find(header.form);
	if (!found) { throw new Error(`No control named ${canvasName} on ${header.form.name}.`); }
	return found;
}

/** The form for '' or its own name, else the named control, which must be one that holds controls. */
function containerOf(header: FrmHeader, name: string): FrmControl {
	if (name === '' || name.toLowerCase() === header.form.name.toLowerCase()) { return header.form; }
	const { control } = locate(header, name);
	const kind = vb6CanvasKind(control.progId);
	// An OCX or UserControl may be a container (an SSTab is); only the
	// intrinsic kinds that are not containers are refused.
	if (kind !== undefined && kind !== 'Frame' && kind !== 'PictureBox') {
		throw new Error(`${name} is a ${kind}; it cannot hold controls.`);
	}
	return control;
}

function contains(ancestor: FrmControl, control: FrmControl): boolean {
	return ancestor.children.some((c) => c === control || contains(c, control));
}

function allNames(header: FrmHeader): Set<string> {
	const names = new Set<string>([header.form.name.toLowerCase()]);
	for (const control of walk(header.form)) { names.add(control.name.toLowerCase()); }
	return names;
}

function freshName(header: FrmHeader, base: string): string {
	const taken = allNames(header);
	let n = 1;
	while (taken.has(`${base}${n}`.toLowerCase())) { n += 1; }
	return `${base}${n}`;
}

function numberProp(control: FrmControl, key: string): number | undefined {
	const p = frmProperty(control, key);
	return p && !p.frx ? frmNumberOf(p.value) : undefined;
}

// --- members, in the designer's order --------------------------------------

function memberName(m: FrmPropertyNode): string {
	return m.kind === 'property' ? m.key : m.name;
}

function collate(a: string, b: string): number {
	const x = a.toLowerCase();
	const y = b.toLowerCase();
	return x < y ? -1 : x > y ? 1 : 0;
}

/**
 * Replaces a member in place, or inserts it where the designer would: the
 * designer writes a control's members in name order, case-insensitively,
 * groups included (Font sits between ClientWidth and Icon). A header not
 * in that order takes the new member at its end.
 */
function setMember(members: FrmPropertyNode[], node: FrmPropertyNode): void {
	const name = memberName(node).toLowerCase();
	const at = members.findIndex((m) => m.kind === node.kind && memberName(m).toLowerCase() === name);
	if (at >= 0) {
		members[at] = node;
		return;
	}
	const sorted = members.every((m, i) => i === 0 || collate(memberName(members[i - 1]), m.kind === 'property' ? m.key : m.name) <= 0);
	const after = sorted ? members.findIndex((m) => collate(memberName(m), memberName(node)) > 0) : -1;
	if (after < 0) { members.push(node); } else { members.splice(after, 0, node); }
}

function removeMember(control: FrmControl, key: string): void {
	const lower = key.toLowerCase();
	control.members = control.members.filter((m) => !(m.kind === 'property' && m.key.toLowerCase() === lower));
}

function setNumber(control: FrmControl, key: string, value: number): void {
	const existing = frmProperty(control, key);
	setMember(control.members, { kind: 'property', key: existing?.key ?? key, value: String(value) });
}

function quoted(value: string): string {
	return `"${value.replace(/"/g, '""')}"`;
}

/** The designer's own boolean spellings: `-1  'True`, and `0   'False` with the zero padded to the width of -1. */
function boolMember(key: string, on: boolean): FrmProperty {
	return on ? { kind: 'property', key, value: '-1', comment: 'True' } : { kind: 'property', key, value: "0   'False" };
}

function stringMember(key: string, value: string, options: FrmDesignerOpOptions): FrmProperty {
	// eslint-disable-next-line no-control-regex
	if (/[\u0000-\u001f]/.test(value)) {
		if (!options.storeString) {
			throw new Error(`${key}: a value with line breaks is stored in the form's .frx sidecar, which this operation cannot write.`);
		}
		const ref = options.storeString(value);
		const offset = ref.offset.toString(16).toUpperCase().padStart(4, '0');
		return { kind: 'property', key, value: `${ref.long ? '$' : ''}"${ref.file}":${offset}`, frx: ref };
	}
	return { kind: 'property', key, value: quoted(value) };
}

function colorMember(key: string, value: string): FrmProperty {
	const color = parseOleColor(value);
	if (color === undefined) { throw new Error(`${key}: ${value} is not a color (#rrggbb, &H00BBGGRR&, or a system color name).`); }
	return { kind: 'property', key, value: `&H${(color >>> 0).toString(16).toUpperCase().padStart(8, '0')}&` };
}

function isTrue(value: string): boolean {
	const v = value.trim().toLowerCase();
	return v === 'true' || v === '-1' || v === 'yes' || v === 'on';
}

function isBooleanText(value: string): boolean {
	const v = value.trim().toLowerCase();
	return v === 'true' || v === 'false';
}

// --- geometry ----------------------------------------------------------------

function setGeometry(header: FrmHeader, g: FrmGeometry): void {
	const { control } = locate(header, g.name);
	const kind = vb6CanvasKind(control.progId);
	if (kind === 'Line') {
		setLineGeometry(control, g);
		return;
	}
	if (g.left !== undefined) { setNumber(control, 'Left', twipsOfPt(g.left)); }
	if (g.top !== undefined) { setNumber(control, 'Top', twipsOfPt(g.top)); }
	if (kind === 'Timer') { return; }
	if (g.width !== undefined) { setNumber(control, 'Width', twipsOfPt(g.width)); }
	if (g.height !== undefined) { setNumber(control, 'Height', twipsOfPt(g.height)); }
}

/** A Line is its two points; the canvas moves and resizes their bounding box, and the ends keep their sides. */
function setLineGeometry(control: FrmControl, g: FrmGeometry): void {
	const x1 = numberProp(control, 'X1') ?? 0;
	const y1 = numberProp(control, 'Y1') ?? 0;
	const x2 = numberProp(control, 'X2') ?? x1;
	const y2 = numberProp(control, 'Y2') ?? y1;
	const left = Math.min(x1, x2);
	const top = Math.min(y1, y2);
	const width = Math.abs(x2 - x1);
	const height = Math.abs(y2 - y1);
	const newLeft = g.left !== undefined ? twipsOfPt(g.left) : left;
	const newTop = g.top !== undefined ? twipsOfPt(g.top) : top;
	// The canvas draws a flat line one point thick; that thickness is not a length.
	const newWidth = g.width !== undefined && !(width === 0 && twipsOfPt(g.width) <= 20) ? twipsOfPt(g.width) : width;
	const newHeight = g.height !== undefined && !(height === 0 && twipsOfPt(g.height) <= 20) ? twipsOfPt(g.height) : height;
	setNumber(control, 'X1', x1 <= x2 ? newLeft : newLeft + newWidth);
	setNumber(control, 'X2', x1 <= x2 ? newLeft + newWidth : newLeft);
	setNumber(control, 'Y1', y1 <= y2 ? newTop : newTop + newHeight);
	setNumber(control, 'Y2', y1 <= y2 ? newTop + newHeight : newTop);
}

function setFormSize(header: FrmHeader, width: number, height: number): void {
	const form = header.form;
	setNumber(form, 'ClientWidth', width);
	setNumber(form, 'ClientHeight', height);
	// The scale mirrors the client size only when the form measures in twips.
	if ((numberProp(form, 'ScaleMode') ?? 1) === 1) {
		if (frmProperty(form, 'ScaleWidth')) { setNumber(form, 'ScaleWidth', width); }
		if (frmProperty(form, 'ScaleHeight')) { setNumber(form, 'ScaleHeight', height); }
	}
}

// --- add ---------------------------------------------------------------------

interface NewControlSpec {
	/** The designer's base name: Command1, Text1, Picture1. */
	base: string;
	width?: number;
	height?: number;
	/** The property that carries the control's own name as its first value. */
	text?: 'Caption' | 'Text';
	/** Whether the kind takes a TabIndex. */
	tab: boolean;
	/** A PictureBox writes its scale, the client area inside its border. */
	scale?: boolean;
}

/**
 * What the designer writes for a control dropped from the toolbox: its
 * base name, its default size in twips, the caption or text that repeats
 * the name, a TabIndex where the kind has one. The sizes are the toolbox
 * defaults as `.frm` files commonly show them; any size is a valid header.
 */
const NEW_CONTROL: Readonly<Record<string, NewControlSpec>> = {
	Label: { base: 'Label', width: 1215, height: 255, text: 'Caption', tab: true },
	TextBox: { base: 'Text', width: 1215, height: 285, text: 'Text', tab: true },
	ComboBox: { base: 'Combo', width: 1215, height: 315, text: 'Text', tab: true },
	ListBox: { base: 'List', width: 1215, height: 1035, tab: true },
	CheckBox: { base: 'Check', width: 1215, height: 255, text: 'Caption', tab: true },
	OptionButton: { base: 'Option', width: 1215, height: 255, text: 'Caption', tab: true },
	CommandButton: { base: 'Command', width: 1215, height: 495, text: 'Caption', tab: true },
	Frame: { base: 'Frame', width: 1215, height: 1215, text: 'Caption', tab: true },
	PictureBox: { base: 'Picture', width: 1215, height: 1215, tab: true, scale: true },
	Image: { base: 'Image', width: 1215, height: 1215, tab: false },
	HScrollBar: { base: 'HScroll', width: 1215, height: 255, tab: true },
	VScrollBar: { base: 'VScroll', width: 255, height: 1215, tab: true },
	Timer: { base: 'Timer', tab: false },
	Line: { base: 'Line', tab: false },
	Shape: { base: 'Shape', width: 1215, height: 1215, tab: false },
};

function nextTabIndex(header: FrmHeader): number {
	let max = -1;
	for (const control of walk(header.form)) {
		const tab = numberProp(control, 'TabIndex');
		if (tab !== undefined && tab > max) { max = tab; }
	}
	return max + 1;
}

function addControl(header: FrmHeader, containerName: string, controlKind: string, left: number, top: number): string {
	const spec = NEW_CONTROL[controlKind];
	if (!spec) { throw new Error(`${controlKind} is not a control the VB6 toolbox offers.`); }
	const container = containerOf(header, containerName);
	const name = freshName(header, spec.base);
	const members: FrmPropertyNode[] = [];
	const num = (key: string, value: number): void => { members.push({ kind: 'property', key, value: String(value) }); };
	if (spec.text === 'Caption') { members.push({ kind: 'property', key: 'Caption', value: quoted(name) }); }
	if (spec.text === 'Text') { members.push({ kind: 'property', key: 'Text', value: quoted(name) }); }
	if (controlKind === 'Line') {
		num('X1', left);
		num('X2', left + 1215);
		num('Y1', top);
		num('Y2', top);
	} else {
		num('Left', left);
		num('Top', top);
		if (spec.width !== undefined) { num('Width', spec.width); }
		if (spec.height !== undefined) { num('Height', spec.height); }
		if (spec.scale && spec.width !== undefined && spec.height !== undefined) {
			num('ScaleWidth', spec.width - 60);
			num('ScaleHeight', spec.height - 60);
		}
	}
	if (spec.tab) { num('TabIndex', nextTabIndex(header)); }
	members.sort((a, b) => collate(memberName(a), memberName(b)));
	container.children.push({ progId: `VB.${controlKind}`, name, members, children: [] });
	return name;
}

// --- remove, reparent, order -------------------------------------------------

/** Removes the named controls, outermost only: a child of a removed container goes with it. All-or-nothing on unknown names. */
function removeControls(header: FrmHeader, names: readonly string[]): string[] {
	const located = names.map((name) => ({ name, ...locate(header, name) }));
	const outermost = located.filter((one, i) => located.findIndex((o) => o.control === one.control) === i
		&& !located.some((other) => other.control !== one.control && contains(other.control, one.control)));
	for (const one of outermost) {
		one.parent.children.splice(one.parent.children.indexOf(one.control), 1);
	}
	return outermost.map((one) => one.name);
}

function reparent(header: FrmHeader, name: string, containerName: string, left: number, top: number): void {
	const { control, parent } = locate(header, name);
	const target = containerOf(header, containerName);
	if (target === control || contains(control, target)) { throw new Error(`${name} cannot be placed inside itself.`); }
	parent.children.splice(parent.children.indexOf(control), 1);
	target.children.push(control);
	setGeometry(header, { name, left, top });
}

/** The canvas draws later siblings on top; front is last. */
function setZOrder(header: FrmHeader, name: string, toFront: boolean): void {
	const { control, parent } = locate(header, name);
	parent.children.splice(parent.children.indexOf(control), 1);
	if (toFront) { parent.children.push(control); } else { parent.children.unshift(control); }
}

/**
 * Reorders the tab stops the canvas listed: TabIndex is one sequence across
 * the whole form, so the named controls exchange the indices they already
 * hold, in the new order, and every other control keeps its own.
 */
function setTabOrder(header: FrmHeader, names: readonly string[]): void {
	const controls = names.map((name) => locate(header, name).control).filter((c) => numberProp(c, 'TabIndex') !== undefined);
	const indices = controls.map((c) => numberProp(c, 'TabIndex') as number).sort((a, b) => a - b);
	controls.forEach((control, i) => setNumber(control, 'TabIndex', indices[i]));
}

// --- setProp -----------------------------------------------------------------

const COLOR_KEYS = new Set(['backcolor', 'forecolor', 'fillcolor', 'bordercolor', 'maskcolor']);
const GEOMETRY_KEYS = new Set(['left', 'top', 'width', 'height']);
const IDENTIFIER = /^[A-Za-z][A-Za-z0-9_]{0,39}$/;
const FONT_ORDER = ['Name', 'Size', 'Charset', 'Weight', 'Underline', 'Italic', 'Strikethrough'];

/** Sets one property as the pane spells it; returns the canvas name to select when the name changed. */
function setProp(header: FrmHeader, name: string, prop: string, value: string, options: FrmDesignerOpOptions): string | undefined {
	const isForm = name === '' || name.toLowerCase() === header.form.name.toLowerCase();
	const target = isForm ? header.form : locate(header, name).control;
	const lower = prop.toLowerCase();
	if (lower === 'name') { return rename(header, target, value); }
	if (lower === 'index') { return setIndex(header, target, value); }
	if (isForm && (lower === 'width' || lower === 'height')) {
		const twips = twipsOfPt(Number(value));
		if (!Number.isFinite(twips)) { throw new Error(`${prop}: ${value} is not a size.`); }
		setFormSize(header, lower === 'width' ? twips : numberProp(target, 'ClientWidth') ?? 0, lower === 'height' ? twips : numberProp(target, 'ClientHeight') ?? 0);
		return undefined;
	}
	if (!isForm && GEOMETRY_KEYS.has(lower)) {
		const n = Number(value);
		if (!Number.isFinite(n)) { throw new Error(`${prop}: ${value} is not a length in points.`); }
		setGeometry(header, { name, [lower]: n });
		return undefined;
	}
	if (lower.startsWith('font.')) {
		setFontProp(target, prop.slice(5), value);
		return undefined;
	}
	const existing = frmProperty(target, prop);
	const key = existing?.key ?? prop;
	const trimmed = value.trim();
	if (COLOR_KEYS.has(lower)) {
		setMember(target.members, colorMember(key, trimmed));
	} else if (isBooleanText(trimmed)) {
		setMember(target.members, boolMember(key, isTrue(trimmed)));
	} else if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
		// The same value again keeps the line, gloss included.
		if (existing && !existing.frx && existing.value.trim() === trimmed) { return undefined; }
		const gloss = VB6_ENUM_GLOSSES[`${target.progId}.${key}`]?.[trimmed];
		setMember(target.members, gloss ? { kind: 'property', key, value: trimmed, comment: gloss } : { kind: 'property', key, value: trimmed });
	} else if (trimmed === '' && existing && !/^"/.test(existing.value) && !existing.frx) {
		// Clearing a number or an enum returns it to its default: the line goes.
		removeMember(target, key);
	} else {
		setMember(target.members, stringMember(key, value, options));
	}
	return undefined;
}

function rename(header: FrmHeader, target: FrmControl, newName: string): string {
	const name = newName.trim();
	if (!IDENTIFIER.test(name)) { throw new Error(`${newName} is not a valid control name (a letter, then letters, digits, or underscores).`); }
	if (target === header.form) {
		throw new Error('The form is named by its module: rename the module, whose name is in the project file and the Attribute VB_Name line.');
	}
	const old = target.name.toLowerCase();
	if (name.toLowerCase() !== old) {
		const taken = allNames(header);
		taken.delete(old);
		if (taken.has(name.toLowerCase())) { throw new Error(`A control named ${name} already exists on ${header.form.name}.`); }
	}
	// A control array is one member: every element takes the new name.
	for (const control of walk(header.form)) {
		if (control.name.toLowerCase() === old) { control.name = name; }
	}
	return vb6ControlName(target);
}

/** Makes a control an array element (a number), or a plain control again (empty). */
function setIndex(header: FrmHeader, target: FrmControl, value: string): string {
	if (target === header.form) { throw new Error('A form has no Index.'); }
	const trimmed = value.trim();
	const siblings = [...walk(header.form)].filter((c) => c !== target && c.name.toLowerCase() === target.name.toLowerCase());
	if (trimmed === '') {
		if (siblings.length > 0) { throw new Error(`${target.name} is a control array with ${siblings.length + 1} elements; remove the others before clearing Index.`); }
		removeMember(target, 'Index');
		return target.name;
	}
	if (!/^\d+$/.test(trimmed)) { throw new Error(`Index: ${value} is not a whole number.`); }
	if (siblings.some((c) => (numberProp(c, 'Index') ?? -1) === Number(trimmed))) {
		throw new Error(`${target.name}(${trimmed}) already exists.`);
	}
	setNumber(target, 'Index', Number(trimmed));
	return vb6ControlName(target);
}

function defaultFontMembers(): FrmPropertyNode[] {
	return [
		{ kind: 'property', key: 'Name', value: '"MS Sans Serif"' },
		{ kind: 'property', key: 'Size', value: '8.25' },
		{ kind: 'property', key: 'Charset', value: '0' },
		{ kind: 'property', key: 'Weight', value: '400' },
		boolMember('Underline', false),
		boolMember('Italic', false),
		boolMember('Strikethrough', false),
	];
}

/** Font.Name, Font.Size, Font.Bold (Weight 700 or 400), Font.Italic, Font.Underline, Font.Strikethrough; the group in the designer's fixed order. */
function setFontProp(control: FrmControl, field: string, value: string): void {
	let group = control.members.find((m): m is FrmPropertyGroup => m.kind === 'group' && m.name.toLowerCase() === 'font');
	if (!group) {
		group = { kind: 'group', name: 'Font', members: defaultFontMembers() };
		setMember(control.members, group);
	}
	const f = field.toLowerCase();
	let node: FrmProperty;
	switch (f) {
		case 'name':
			node = { kind: 'property', key: 'Name', value: quoted(value.trim()) };
			break;
		case 'size': {
			const size = Number(value);
			if (!Number.isFinite(size) || size <= 0) { throw new Error(`Font.Size: ${value} is not a size in points.`); }
			node = { kind: 'property', key: 'Size', value: String(size) };
			break;
		}
		case 'bold':
			node = { kind: 'property', key: 'Weight', value: isTrue(value) ? '700' : '400' };
			break;
		case 'weight':
		case 'charset':
			if (!/^\d+$/.test(value.trim())) { throw new Error(`Font.${field}: ${value} is not a whole number.`); }
			node = { kind: 'property', key: f === 'weight' ? 'Weight' : 'Charset', value: value.trim() };
			break;
		case 'italic':
		case 'underline':
		case 'strikethrough':
			node = boolMember(f === 'italic' ? 'Italic' : f === 'underline' ? 'Underline' : 'Strikethrough', isTrue(value));
			break;
		default:
			throw new Error(`Font.${field} is not a font property.`);
	}
	const at = group.members.findIndex((m) => m.kind === 'property' && m.key.toLowerCase() === node.key.toLowerCase());
	if (at >= 0) {
		group.members[at] = node;
		return;
	}
	const rank = FONT_ORDER.indexOf(node.key);
	const before = group.members.findIndex((m) => FONT_ORDER.indexOf(memberName(m)) > rank);
	if (before < 0) { group.members.push(node); } else { group.members.splice(before, 0, node); }
}

// --- duplicate ---------------------------------------------------------------

/** Clones the named controls beside their originals, nudged, under fresh names; a container's children come along and are renamed with it. */
function duplicate(header: FrmHeader, names: readonly string[], nudge: number): string[] {
	const located = names.map((name) => locate(header, name));
	const outermost = located.filter((one, i) => located.findIndex((o) => o.control === one.control) === i
		&& !located.some((other) => other.control !== one.control && contains(other.control, one.control)));
	const taken = allNames(header);
	let nextTab = nextTabIndex(header);
	const newNames: string[] = [];
	for (const one of outermost) {
		const clone = structuredClone(one.control) as FrmControl;
		for (const control of [clone, ...walk(clone)]) {
			// A copy is a new control, never another element of the original's array.
			removeMember(control, 'Index');
			const base = control.name.replace(/\d+$/, '') || 'Copy';
			let n = 1;
			while (taken.has(`${base}${n}`.toLowerCase())) { n += 1; }
			control.name = `${base}${n}`;
			taken.add(control.name.toLowerCase());
			if (frmProperty(control, 'TabIndex')) {
				setNumber(control, 'TabIndex', nextTab);
				nextTab += 1;
			}
		}
		if (vb6CanvasKind(clone.progId) === 'Line') {
			for (const key of ['X1', 'X2', 'Y1', 'Y2']) { setNumber(clone, key, (numberProp(clone, key) ?? 0) + nudge); }
		} else {
			setNumber(clone, 'Left', (numberProp(clone, 'Left') ?? 0) + nudge);
			setNumber(clone, 'Top', (numberProp(clone, 'Top') ?? 0) + nudge);
		}
		one.parent.children.push(clone);
		newNames.push(clone.name);
	}
	return newNames;
}

// --- the document around the header ------------------------------------------

/**
 * Where the header ends in a form's text: after the form block's own `End`,
 * the one header line with no indent (every nested block's End is indented).
 * The whole text when no such line exists; a parse then names the problem.
 */
export function vb6HeaderEndOf(text: string): number {
	const match = /^End[ \t]*\r?\n?/m.exec(text);
	return match ? match.index + match[0].length : text.length;
}

/** The handler prefix a designer's own events use: Form_Load, MDIForm_Load, UserControl_Initialize. */
export function vb6FormHandlerPrefix(text: string): string {
	const match = /^Begin VB\.(\w+)\s/m.exec(text);
	switch (match?.[1]) {
		case 'MDIForm': return 'MDIForm';
		case 'UserControl': return 'UserControl';
		case 'PropertyPage': return 'PropertyPage';
		default: return 'Form';
	}
}
