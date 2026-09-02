// The designer header a VB6 form, UserControl, PropertyPage or Designer file
// opens with: `VERSION 5.00`, the OCX references the form uses, then one
// `Begin <ProgId> <Name>` ... `End` block per control, nested inside their
// containers, holding `Key = Value` property lines and `BeginProperty` groups
// (a Font, an ImageList's images). There is no Microsoft specification of the
// layout; every rule here was measured on forms Visual Basic 6 wrote
// (tests/fixtures/vb6), and the printer reproducing those files byte for byte
// is the proof the rules hold.
//
// Two ways back to text: `printFrmHeader` re-emits the lines as read, so a
// header nothing touched comes back identical; `formatFrmHeader` regenerates
// every line from the model in the designer's own layout, which is what an
// edit uses - and which, on every fixture, also matches the original.

/** A `"Form1.frx":0000` value: a property whose bytes live in the sidecar. */
export interface FrxRef {
	file: string;
	offset: number;
	/** The `$"..."` form: a long string, stored with a 32-bit length. */
	long: boolean;
}

export interface FrmProperty {
	kind: 'property';
	/** As written, including any index or dotted path: `Tab(0).Control(0)`. */
	key: string;
	/** The value text with any trailing comment removed. */
	value: string;
	/** The `'True` / `'Fixed Single` gloss the designer writes after enums and booleans. */
	comment?: string;
	frx?: FrxRef;
}

export interface FrmPropertyGroup {
	kind: 'group';
	/** `Font`, `ListImage1`, `ColumnHeader(1)`. */
	name: string;
	/** The `{...}` class id some groups carry (ImageList images, column headers). */
	guid?: string;
	members: FrmPropertyNode[];
}

export type FrmPropertyNode = FrmProperty | FrmPropertyGroup;

export interface FrmControl {
	/** `VB.CommandButton`, `MSComctlLib.ListView`, `Audiostation.ButtonBig`. */
	progId: string;
	name: string;
	members: FrmPropertyNode[];
	children: FrmControl[];
}

/** `Object = "{GUID}#2.0#0"; "mscomctl.ocx"`: an OCX the form uses. */
export interface FrmObjectRef {
	guid: string;
	version: string;
	file: string;
}

export interface FrmHeader {
	/** The `VERSION 5.00` line as written. */
	version: string;
	objects: FrmObjectRef[];
	/** The form (or UserControl, PropertyPage, Designer) itself. */
	form: FrmControl;
	eol: string;
	/** Every header line as read, `VERSION` through the closing `End`. */
	lines: string[];
	/** The offset in the file text where the module text begins. */
	endOffset: number;
}

export class FrmHeaderError extends Error {}

const BEGIN_RE = /^(\s*)Begin\s+(\S+)\s+(\S+)\s*$/;
const END_RE = /^\s*End\s*$/;
const BEGIN_PROPERTY_RE = /^\s*BeginProperty\s+(\S+)(?:\s+(\{[0-9A-Fa-f-]+\}))?\s*$/;
const END_PROPERTY_RE = /^\s*EndProperty\s*$/;
const OBJECT_RE = /^\s*Object\s*=\s*"\{([0-9A-Fa-f-]+)\}#([^#]+)#[^"]*";\s*"([^"]*)"\s*$/;
const PROPERTY_RE = /^\s*([\w.()]+)\s*=\s*(.*)$/;
const FRX_RE = /^(\$?)"([^"]*)":([0-9A-Fa-f]+)$/;
const COMMENT_RE = /^(-?[\d.,Ee+-]+|&H[0-9A-Fa-f]+&)\s{2}'(.*)$/;

/** Splits a value into its text and the designer's trailing comment, if any. */
function splitValue(raw: string): { value: string; comment?: string; frx?: FrxRef } {
	const trimmed = raw.replace(/\s+$/, '');
	const frx = FRX_RE.exec(trimmed);
	if (frx) {
		return { value: trimmed, frx: { long: frx[1] === '$', file: frx[2], offset: parseInt(frx[3], 16) } };
	}
	const comment = COMMENT_RE.exec(trimmed);
	if (comment) {
		return { value: comment[1], comment: comment[2] };
	}
	return { value: trimmed };
}

/**
 * Parses the header at the top of a file's text. Undefined when the text
 * does not open with a designer header; an error when it opens with one that
 * is malformed (a block that never closes, a stray `End`).
 */
export function parseFrmHeader(text: string): FrmHeader | undefined {
	const eol = text.includes('\r\n') ? '\r\n' : '\n';
	const all = text.split(/\r?\n/);
	if (all.length === 0 || !/^\s*VERSION\s+\d/.test(all[0])) {
		return undefined;
	}
	const header: FrmHeader = { version: all[0], objects: [], form: undefined as unknown as FrmControl, eol, lines: [], endOffset: 0 };
	const controls: FrmControl[] = [];
	const groups: FrmPropertyGroup[] = [];
	let consumed = 0;
	let closed = false;
	for (let i = 0; i < all.length; i++) {
		const line = all[i];
		header.lines.push(line);
		// The offset advances by the line's own newline, whichever it is: a
		// header spliced in with one convention above code with the other
		// must still end exactly where its End line does.
		consumed += line.length;
		if (i < all.length - 1) { consumed += text.startsWith('\r\n', consumed) ? 2 : 1; }
		if (i === 0) {
			continue;
		}
		const begin = BEGIN_RE.exec(line);
		if (begin) {
			const control: FrmControl = { progId: begin[2], name: begin[3], members: [], children: [] };
			if (controls.length === 0) {
				header.form = control;
			} else {
				controls[controls.length - 1].children.push(control);
			}
			controls.push(control);
			continue;
		}
		if (END_RE.test(line)) {
			if (groups.length > 0 || controls.length === 0) {
				throw new FrmHeaderError(`Line ${i + 1}: End with no open control block.`);
			}
			controls.pop();
			if (controls.length === 0) {
				closed = true;
				break;
			}
			continue;
		}
		if (controls.length === 0) {
			const object = OBJECT_RE.exec(line);
			if (object) {
				header.objects.push({ guid: `{${object[1]}}`, version: object[2], file: object[3] });
				continue;
			}
			if (line.trim() === '') {
				continue;
			}
			throw new FrmHeaderError(`Line ${i + 1}: expected Object= or Begin before the form block.`);
		}
		const beginProperty = BEGIN_PROPERTY_RE.exec(line);
		if (beginProperty) {
			const group: FrmPropertyGroup = { kind: 'group', name: beginProperty[1], members: [] };
			if (beginProperty[2]) {
				group.guid = beginProperty[2];
			}
			currentMembers(controls, groups).push(group);
			groups.push(group);
			continue;
		}
		if (END_PROPERTY_RE.test(line)) {
			if (groups.length === 0) {
				throw new FrmHeaderError(`Line ${i + 1}: EndProperty with no open property group.`);
			}
			groups.pop();
			continue;
		}
		const property = PROPERTY_RE.exec(line);
		if (property) {
			currentMembers(controls, groups).push({ kind: 'property', key: property[1], ...splitValue(property[2]) });
			continue;
		}
		if (line.trim() === '') {
			continue;
		}
		throw new FrmHeaderError(`Line ${i + 1}: not a designer header line: ${line.trim().slice(0, 40)}`);
	}
	if (!closed) {
		throw new FrmHeaderError('The form block never closes.');
	}
	header.endOffset = consumed;
	return header;
}

function currentMembers(controls: FrmControl[], groups: FrmPropertyGroup[]): FrmPropertyNode[] {
	return groups.length > 0 ? groups[groups.length - 1].members : controls[controls.length - 1].members;
}

/** The header text exactly as it was read. */
export function printFrmHeader(header: FrmHeader): string {
	return header.lines.join(header.eol) + header.eol;
}

const INDENT = '   ';
/** The designer pads a key to this width before the `=`; a longer key gets no padding. */
const KEY_WIDTH = 16;
/**
 * The extender properties of a control hosted on a UserControl or a custom
 * control (`Object.Width`, `Object.Visible`) pad wider. Measured on five of
 * thirty-four forms in the census corpus, the only files that carry them.
 */
const EXTENDER_KEY_WIDTH = 23;

function formatProperty(depth: number, property: FrmProperty): string {
	const width = /^Object\./.test(property.key) ? EXTENDER_KEY_WIDTH : KEY_WIDTH;
	const key = property.key.length >= width ? `${property.key}=` : `${property.key.padEnd(width)}=`;
	const value = property.comment !== undefined ? `${property.value}  '${property.comment}` : property.value;
	return `${INDENT.repeat(depth)}${key}   ${value}`;
}

function formatMembers(depth: number, members: readonly FrmPropertyNode[], out: string[]): void {
	for (const member of members) {
		if (member.kind === 'property') {
			out.push(formatProperty(depth, member));
			continue;
		}
		out.push(`${INDENT.repeat(depth)}BeginProperty ${member.name}${member.guid ? ` ${member.guid}` : ''} `);
		formatMembers(depth + 1, member.members, out);
		out.push(`${INDENT.repeat(depth)}EndProperty`);
	}
}

function formatControl(depth: number, control: FrmControl, out: string[]): void {
	out.push(`${INDENT.repeat(depth)}Begin ${control.progId} ${control.name} `);
	formatMembers(depth + 1, control.members, out);
	for (const child of control.children) {
		formatControl(depth + 1, child, out);
	}
	out.push(`${INDENT.repeat(depth)}End`);
}

/**
 * The header regenerated from the model in the designer's own layout: three
 * spaces per level, keys padded to sixteen columns before the `=`, three
 * spaces after it, a trailing space after every `Begin` and `BeginProperty`
 * name, two spaces before a value's comment. Measured, and pinned by the
 * fixtures printing back identical.
 */
export function formatFrmHeader(header: FrmHeader): string {
	const out: string[] = [header.version];
	for (const object of header.objects) {
		out.push(`Object = "${object.guid}#${object.version}#0"; "${object.file}"`);
	}
	formatControl(0, header.form, out);
	return out.join(header.eol) + header.eol;
}

/** A control as the code-behind sees it: one member per name, arrays folded. */
export interface FrmMember {
	name: string;
	/** The prog id, which is the type the `vb6` model keys on: `VB.TextBox`. */
	type: string;
	/** True when the designer holds several controls of this name with an `Index`. */
	array: boolean;
	/** The container's name, or the form's for a top-level control. */
	parent: string;
}

/** Every control on the form in document order, the form itself excluded. */
export function frmControls(header: FrmHeader): Array<FrmControl & { parent: string; index?: number }> {
	const out: Array<FrmControl & { parent: string; index?: number }> = [];
	const walk = (control: FrmControl): void => {
		for (const child of control.children) {
			const indexProperty = child.members.find(
				(m): m is FrmProperty => m.kind === 'property' && m.key.toLowerCase() === 'index',
			);
			const index = indexProperty ? Number(indexProperty.value) : undefined;
			out.push({ ...child, parent: control.name, ...(index !== undefined && !Number.isNaN(index) ? { index } : {}) });
			walk(child);
		}
	};
	walk(header.form);
	return out;
}

/** The members the form declares through its designer, one per name. */
export function frmMembers(header: FrmHeader): FrmMember[] {
	const byName = new Map<string, FrmMember>();
	for (const control of frmControls(header)) {
		const key = control.name.toLowerCase();
		const existing = byName.get(key);
		if (existing) {
			existing.array = true;
			continue;
		}
		byName.set(key, {
			name: control.name,
			type: control.progId,
			array: control.index !== undefined,
			parent: control.parent,
		});
	}
	return [...byName.values()];
}

/** The property, wherever it sits, when the control has it (case-insensitive). */
export function frmProperty(control: FrmControl, key: string): FrmProperty | undefined {
	const wanted = key.toLowerCase();
	return control.members.find((m): m is FrmProperty => m.kind === 'property' && m.key.toLowerCase() === wanted);
}

/** Every `.frx` reference in the header, in document order, with the control that owns it. */
export function frmFrxRefs(header: FrmHeader): Array<{ control: FrmControl; property: FrmProperty; group?: string }> {
	const out: Array<{ control: FrmControl; property: FrmProperty; group?: string }> = [];
	const walkMembers = (control: FrmControl, members: readonly FrmPropertyNode[], group?: string): void => {
		for (const member of members) {
			if (member.kind === 'property') {
				if (member.frx) {
					out.push({ control, property: member, ...(group ? { group } : {}) });
				}
			} else {
				walkMembers(control, member.members, group ? `${group}.${member.name}` : member.name);
			}
		}
	};
	const walk = (control: FrmControl): void => {
		walkMembers(control, control.members);
		for (const child of control.children) {
			walk(child);
		}
	};
	walk(header.form);
	return out;
}
