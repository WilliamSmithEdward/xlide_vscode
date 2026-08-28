// The form-as-text projection: XLIDE form markup.
//
// The dialect is xlide_vbide's, adopted here so both products speak the same
// document: XAML-shaped elements, every value quoted, one level of dotting
// for fonts, colors spelled `#rrggbb` for literals and by NAME for system
// colors (a system color is a question about the machine, and freezing it to
// today's answer would stop the control following the theme).
//
// The rules that make the document honest, from the vbide design:
//   - projection, not source: generated FROM the binary, applied back TO it;
//   - the control list is total, the property list is not - an unspoken
//     property is never touched on apply;
//   - apply is a diff keyed by control name: only-in-markup adds,
//     only-in-model removes, matched controls set what changed;
//   - a parse error applies nothing.

import { himetricToPoints, pointsToHimetric, OformsReader, OformsWriter } from './bytes';
import {
	recordHas,
	setRecordString,
	setRecordValue,
	RECORD_SPECS_BY_CACHE_INDEX,
	TEXT_PROPS_SPEC as TEXT_PROPS_REF,
	type ParsedRecord,
} from './records';
import {
	siteName,
	siteId,
	siteCacheIndex,
	siteIsContainer,
	type SiteModel,
} from './formStream';
import { containerStorageName, controlKindOfSite, type FormPackage } from './formPackage';

// ------------------------------------------------------------------ colors

/** System palette names for OLE_COLOR 0x80xxxxxx values, by index. */
const SYSTEM_COLOR_NAMES: readonly string[] = [
	'ScrollBars', 'Desktop', 'ActiveTitleBar', 'InactiveTitleBar', 'MenuBar',
	'WindowBackground', 'WindowFrame', 'MenuText', 'WindowText', 'TitleBarText',
	'ActiveBorder', 'InactiveBorder', 'ApplicationWorkspace', 'Highlight',
	'HighlightText', 'ButtonFace', 'ButtonShadow', 'GrayText', 'ButtonText',
	'InactiveTitleBarText', 'ButtonHighlight', 'ButtonDarkShadow', 'ButtonLight',
	'InfoText', 'InfoBackground', 'HotTracking', 'GradientActiveTitleBar',
	'GradientInactiveTitleBar', 'MenuHighlight', 'MenuBackground',
];

const SYSTEM_COLOR_INDEX: ReadonlyMap<string, number> = new Map(
	SYSTEM_COLOR_NAMES.map((name, index) => [name.toLowerCase(), index]),
);

/** OLE_COLOR -> markup spelling. Raw layout is 0xTTBBGGRR little-endian. */
export function formatOleColor(value: number): string {
	const type = (value >>> 24) & 0xff;
	if (type === 0x80) {
		const name = SYSTEM_COLOR_NAMES[value & 0xffffff];
		if (name) { return name; }
	}
	const r = value & 0xff;
	const g = (value >>> 8) & 0xff;
	const b = (value >>> 16) & 0xff;
	const hex = (n: number): string => n.toString(16).padStart(2, '0');
	return `#${hex(r)}${hex(g)}${hex(b)}`;
}

/** Markup spelling -> OLE_COLOR, accepting #hex, system names, and &H...&. */
export function parseOleColor(text: string): number | undefined {
	const t = text.trim();
	const sys = SYSTEM_COLOR_INDEX.get(t.toLowerCase());
	if (sys !== undefined) { return (0x80000000 | sys) >>> 0; }
	const hex = /^#([0-9a-f]{6})$/i.exec(t);
	if (hex) {
		const rgb = parseInt(hex[1], 16);
		const r = (rgb >>> 16) & 0xff;
		const g = (rgb >>> 8) & 0xff;
		const b = rgb & 0xff;
		return (r | (g << 8) | (b << 16)) >>> 0;
	}
	const vb = /^&h([0-9a-f]{1,8})&?$/i.exec(t);
	if (vb) { return parseInt(vb[1], 16) >>> 0; }
	return undefined;
}

// ------------------------------------------------------------------ numbers

/**
 * Points, printed as the SHORTEST decimal that converts back to the same
 * HIMETRIC value. Excel stores 12 pt as round(12 * 2540/72) = 423, and the
 * naive back-conversion prints 11.99; the designer says 12, and so should
 * the document.
 */
function formatPoints(himetric: number): string {
	const pt = himetricToPoints(himetric);
	for (const decimals of [0, 1, 2]) {
		const factor = 10 ** decimals;
		const candidate = Math.round(pt * factor) / factor;
		if (pointsToHimetric(candidate) === himetric) {
			return String(candidate);
		}
	}
	return String(Math.round(pt * 100) / 100);
}

function escapeAttr(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;');
}

function unescapeAttr(text: string): string {
	return text
		.replace(/&quot;/g, '"')
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&');
}

// -------------------------------------------------------------- tab captions

/** Decodes a TabStrip Items array (ArrayStrings) into its captions. */
export function decodeArrayStrings(raw: Buffer): string[] {
	const r = new OformsReader(raw);
	const out: string[] = [];
	while (r.pos < raw.length) {
		const start = r.pos;
		const cc = r.u32();
		const cch = cc & 0x7fffffff;
		const compressed = (cc & 0x80000000) !== 0;
		const bytes = r.bytes(compressed ? cch : cch * 2);
		out.push(bytes.toString(compressed ? 'latin1' : 'utf16le'));
		const over = (r.pos - start) % 4;
		if (over !== 0) { r.bytes(4 - over); }
	}
	return out;
}

/** Encodes captions back to an ArrayString run (compressed when they fit). */
export function encodeArrayStrings(captions: readonly string[]): Buffer {
	const w = new OformsWriter();
	for (const caption of captions) {
		const start = w.position;
		const compressed = [...caption].every((c) => c.charCodeAt(0) <= 0xff);
		w.u32((caption.length & 0x7fffffff) | (compressed ? 0x80000000 : 0));
		w.bytes(Buffer.from(caption, compressed ? 'latin1' : 'utf16le'));
		const over = (w.position - start) % 4;
		if (over !== 0) { w.bytes(Buffer.alloc(4 - over)); }
	}
	return w.toBuffer();
}

// ------------------------------------------------------------------- print

/** Numeric record fields printed per kind, in this order, when stored. */
const PRINTED_FIELDS: Readonly<Record<string, readonly string[]>> = {
	TextBox: ['MaxLength', 'BorderStyle', 'SpecialEffect', 'ScrollBars'],
	ComboBox: ['BorderStyle', 'SpecialEffect', 'ListRows'],
	ListBox: ['BorderStyle', 'SpecialEffect'],
	CheckBox: ['SpecialEffect'],
	OptionButton: ['SpecialEffect'],
	ToggleButton: [],
	Label: ['BorderStyle', 'SpecialEffect'],
	CommandButton: [],
	Image: ['BorderStyle', 'SpecialEffect', 'PictureSizeMode', 'PictureAlignment'],
	SpinButton: ['Min', 'Max', 'Position', 'SmallChange', 'Orientation', 'Delay'],
	ScrollBar: ['Min', 'Max', 'Position', 'SmallChange', 'LargeChange', 'Orientation', 'ProportionalThumb', 'Delay'],
	TabStrip: ['TabOrientation', 'TabStyle'],
	MultiPage: [],
	Frame: [],
	Page: [],
};

const COLOR_FIELDS = new Set(['BackColor', 'ForeColor', 'BorderColor']);

export function printFormMarkup(
	pkg: FormPackage,
	formName: string,
	options: { captionFallback?: string } = {},
): string {
	const lines: string[] = [];
	lines.push(`<!-- ${formName}, as xlide projects it. Edits here apply back to the workbook. -->`);
	printPackage(pkg, formName, lines, 0, 'Form', undefined, options.captionFallback);
	return lines.join('\r\n') + '\r\n';
}

function printPackage(
	pkg: FormPackage,
	name: string,
	lines: string[],
	depth: number,
	tag: 'Form' | 'Frame' | 'Page',
	site?: SiteModel,
	captionFallback?: string,
): void {
	const indent = '    '.repeat(depth);
	const attrs: string[] = [`Name="${escapeAttr(name)}"`];
	const record = pkg.form.record;
	// A top-level form's caption is persisted in the VBFrame text, not in the
	// f record - the service passes it in; a Frame's caption IS in its record.
	const caption = record.strings.get('Caption');
	if (caption) { attrs.push(`Caption="${escapeAttr(caption.text)}"`); }
	else if (tag === 'Form' && captionFallback !== undefined) {
		attrs.push(`Caption="${escapeAttr(captionFallback)}"`);
	}
	if (tag === 'Form') {
		const size = record.sizes.get('DisplayedSize');
		if (size) {
			attrs.push(`Width="${formatPoints(size.width)}"`, `Height="${formatPoints(size.height)}"`);
		}
	} else if (site) {
		pushGeometry(attrs, site, pkg);
	}
	for (const field of ['BackColor', 'ForeColor', 'BorderColor'] as const) {
		const v = record.values.get(field);
		if (v !== undefined && recordHas(record, field)) {
			attrs.push(`${field}="${formatOleColor(v)}"`);
		}
	}
	const se = record.values.get('SpecialEffect');
	if (se !== undefined && recordHas(record, 'SpecialEffect')) {
		attrs.push(`SpecialEffect="${se}"`);
	}
	pushSiteExtras(attrs, site);
	const children = printableChildren(pkg);
	if (children.length === 0) {
		lines.push(`${indent}<${tag} ${attrs.join(' ')} />`);
		return;
	}
	lines.push(`${indent}<${tag} ${attrs.join(' ')}>`);
	for (const child of children) { printChild(pkg, child, lines, depth + 1); }
	lines.push(`${indent}</${tag}>`);
}

interface PrintableChild {
	site: SiteModel;
	record?: ParsedRecord;
	raw?: boolean;
}

function printableChildren(pkg: FormPackage): PrintableChild[] {
	const out: PrintableChild[] = [];
	for (const entry of pkg.entries) {
		if (entry.kind === 'record') { out.push({ site: entry.site, record: entry.record }); }
		else if (entry.kind === 'raw') { out.push({ site: entry.site, raw: true }); }
		else { out.push({ site: entry.site }); }
	}
	return out;
}

function printChild(pkg: FormPackage, child: PrintableChild, lines: string[], depth: number): void {
	const indent = '    '.repeat(depth);
	const site = child.site;
	const kind = controlKindOfSite(site, child.record);
	const name = siteName(site);

	if (siteIsContainer(site)) {
		const inner = pkg.containers.get(siteId(site));
		if (!inner) {
			lines.push(`${indent}<!-- container ${escapeAttr(name)} has no storage -->`);
			return;
		}
		if (kind === 'Frame') {
			printPackage(inner, name, lines, depth, 'Frame', site);
			return;
		}
		if (kind === 'MultiPage') {
			printMultiPage(pkg, inner, site, lines, depth);
			return;
		}
		printPackage(inner, name, lines, depth, 'Page', site);
		return;
	}

	const attrs: string[] = [`Name="${escapeAttr(name)}"`];
	pushGeometry(attrs, site, undefined, child.record);
	const record = child.record;
	if (record) {
		const caption = record.strings.get('Caption');
		if (caption) { attrs.push(`Caption="${escapeAttr(caption.text)}"`); }
		const value = record.strings.get('Value');
		if (value) { attrs.push(`Value="${escapeAttr(value.text)}"`); }
		const group = record.strings.get('GroupName');
		if (group && group.text.length) { attrs.push(`GroupName="${escapeAttr(group.text)}"`); }
		for (const field of COLOR_FIELDS) {
			const v = record.values.get(field);
			if (v !== undefined && recordHas(record, field)) {
				attrs.push(`${field}="${formatOleColor(v)}"`);
			}
		}
		for (const field of PRINTED_FIELDS[kind] ?? []) {
			if (COLOR_FIELDS.has(field)) { continue; }
			const v = record.values.get(field);
			if (v !== undefined && recordHas(record, field)) {
				attrs.push(`${field}="${v}"`);
			}
		}
		const pw = record.values.get('PasswordChar');
		if (pw !== undefined && recordHas(record, 'PasswordChar') && pw !== 0) {
			attrs.push(`PasswordChar="${escapeAttr(String.fromCharCode(pw))}"`);
		}
		const accel = record.values.get('Accelerator');
		if (accel !== undefined && recordHas(record, 'Accelerator') && accel !== 0) {
			attrs.push(`Accelerator="${escapeAttr(String.fromCharCode(accel))}"`);
		}
		pushFont(attrs, record);
	}
	pushSiteExtras(attrs, site);
	if (child.raw) {
		attrs.push('ProgId=""');
		lines.push(`${indent}<ActiveX ${attrs.join(' ')} />`);
		return;
	}

	if (kind === 'TabStrip' && record) {
		const items = record.arrays.get('Items');
		const captions = items ? decodeArrayStrings(items) : [];
		if (captions.length) {
			lines.push(`${indent}<TabStrip ${attrs.join(' ')}>`);
			for (const c of captions) {
				lines.push(`${indent}    <Tab Caption="${escapeAttr(c)}" />`);
			}
			lines.push(`${indent}</TabStrip>`);
			return;
		}
	}
	lines.push(`${indent}<${kind} ${attrs.join(' ')} />`);
}

function printMultiPage(
	parent: FormPackage,
	mp: FormPackage,
	site: SiteModel,
	lines: string[],
	depth: number,
): void {
	const indent = '    '.repeat(depth);
	const attrs: string[] = [`Name="${escapeAttr(siteName(site))}"`];
	pushGeometry(attrs, site, mp);
	pushSiteExtras(attrs, site);
	// Page captions live on the MultiPage's own TabStrip record (its o).
	const tabStrip = mp.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
	const captions = tabStrip && tabStrip.kind === 'record'
		? decodeArrayStrings(tabStrip.record.arrays.get('Items') ?? Buffer.alloc(0))
		: [];
	lines.push(`${indent}<MultiPage ${attrs.join(' ')}>`);
	const pageSites = mp.form.sites.filter((s) => siteCacheIndex(s) === 7);
	pageSites.forEach((pageSite, index) => {
		const inner = mp.containers.get(siteId(pageSite));
		const pageIndent = '    '.repeat(depth + 1);
		const pageAttrs: string[] = [`Name="${escapeAttr(siteName(pageSite))}"`];
		const caption = captions[index];
		if (caption !== undefined) { pageAttrs.push(`Caption="${escapeAttr(caption)}"`); }
		if (!inner) {
			lines.push(`${pageIndent}<Page ${pageAttrs.join(' ')} />`);
			return;
		}
		const children = printableChildren(inner);
		if (children.length === 0) {
			lines.push(`${pageIndent}<Page ${pageAttrs.join(' ')} />`);
			return;
		}
		lines.push(`${pageIndent}<Page ${pageAttrs.join(' ')}>`);
		for (const child of children) { printChild(inner, child, lines, depth + 2); }
		lines.push(`${pageIndent}</Page>`);
	});
	lines.push(`${indent}</MultiPage>`);
}

function pushGeometry(
	attrs: string[],
	site: SiteModel,
	containerPkg?: FormPackage,
	record?: ParsedRecord,
): void {
	const pos = site.position;
	if (pos) {
		attrs.push(`Left="${formatPoints(pos.left)}"`, `Top="${formatPoints(pos.top)}"`);
	}
	const size = record?.sizes.get('Size')
		?? containerPkg?.form.record.sizes.get('DisplayedSize');
	if (size) {
		attrs.push(`Width="${formatPoints(size.width)}"`, `Height="${formatPoints(size.height)}"`);
	}
}

function pushSiteExtras(attrs: string[], site: SiteModel | undefined): void {
	if (!site) { return; }
	const tab = site.values.get('TabIndex');
	if (tab !== undefined && tab >= 0) { attrs.push(`TabIndex="${tab}"`); }
	const tip = site.strings.get('ControlTipText');
	if (tip && tip.text.length) { attrs.push(`ControlTipText="${escapeAttr(tip.text)}"`); }
	const tag = site.strings.get('Tag');
	if (tag && tag.text.length) { attrs.push(`Tag="${escapeAttr(tag.text)}"`); }
}

function pushFont(attrs: string[], record: ParsedRecord): void {
	const tp = record.textProps;
	if (!tp) { return; }
	const name = tp.strings.get('FontName');
	if (name && recordHas(tp, 'FontName')) { attrs.push(`Font.Name="${escapeAttr(name.text)}"`); }
	const height = tp.values.get('FontHeight');
	if (height !== undefined && recordHas(tp, 'FontHeight')) {
		// TextProps FontHeight is stored in twips: Tahoma 8.25pt is 165.
		attrs.push(`Font.Size="${Math.round((height / 20) * 100) / 100}"`);
	}
	const effects = tp.values.get('FontEffects');
	if (effects !== undefined && recordHas(tp, 'FontEffects')) {
		if (effects & 0x1) { attrs.push('Font.Bold="True"'); }
		if (effects & 0x2) { attrs.push('Font.Italic="True"'); }
	}
	const weight = tp.values.get('FontWeight');
	if (weight !== undefined && recordHas(tp, 'FontWeight') && weight >= 600
		&& !(effects !== undefined && (effects & 0x1))) {
		attrs.push('Font.Bold="True"');
	}
}

// ------------------------------------------------------------------- parse

export interface MarkupElement {
	tag: string;
	attrs: Map<string, string>;
	children: MarkupElement[];
	line: number;
}

export class FormMarkupError extends Error {
	constructor(readonly line: number, message: string) {
		super(`line ${line}: ${message}`);
	}
}

/** Parses the whole document; any error throws and nothing applies. */
export function parseFormMarkup(text: string): MarkupElement {
	const tokens = tokenize(text);
	let index = 0;
	const parseElement = (): MarkupElement => {
		const open = tokens[index];
		if (!open || open.kind !== 'open') {
			throw new FormMarkupError(open?.line ?? 1, 'expected an element');
		}
		index++;
		const el: MarkupElement = { tag: open.tag, attrs: open.attrs, children: [], line: open.line };
		if (open.selfClosed) { return el; }
		for (;;) {
			const next = tokens[index];
			if (!next) {
				throw new FormMarkupError(open.line, `<${open.tag}> is never closed`);
			}
			if (next.kind === 'close') {
				if (next.tag.toLowerCase() !== open.tag.toLowerCase()) {
					throw new FormMarkupError(next.line, `</${next.tag}> closes <${open.tag}>`);
				}
				index++;
				return el;
			}
			el.children.push(parseElement());
		}
	};
	const root = parseElement();
	if (index !== tokens.length) {
		throw new FormMarkupError(tokens[index].line, 'content after the closing </Form>');
	}
	if (root.tag.toLowerCase() !== 'form') {
		throw new FormMarkupError(root.line, `the root element must be <Form>, not <${root.tag}>`);
	}
	return root;
}

type Token =
	| { kind: 'open'; tag: string; attrs: Map<string, string>; selfClosed: boolean; line: number }
	| { kind: 'close'; tag: string; line: number };

function tokenize(text: string): Token[] {
	const tokens: Token[] = [];
	let pos = 0;
	let line = 1;
	const bump = (upTo: number): void => {
		for (let i = pos; i < upTo; i++) {
			if (text[i] === '\n') { line++; }
		}
		pos = upTo;
	};
	for (;;) {
		const lt = text.indexOf('<', pos);
		if (lt < 0) {
			const rest = text.slice(pos).trim();
			if (rest.length) { throw new FormMarkupError(line, `unexpected text: ${rest.slice(0, 40)}`); }
			return tokens;
		}
		const between = text.slice(pos, lt).trim();
		if (between.length) { throw new FormMarkupError(line, `unexpected text: ${between.slice(0, 40)}`); }
		bump(lt);
		if (text.startsWith('<!--', pos)) {
			const end = text.indexOf('-->', pos);
			if (end < 0) { throw new FormMarkupError(line, 'unterminated comment'); }
			bump(end + 3);
			continue;
		}
		if (text[pos + 1] === '/') {
			const gt = text.indexOf('>', pos);
			if (gt < 0) { throw new FormMarkupError(line, 'unterminated closing tag'); }
			const tag = text.slice(pos + 2, gt).trim();
			tokens.push({ kind: 'close', tag, line });
			bump(gt + 1);
			continue;
		}
		const startLine = line;
		const { tag, attrs, selfClosed, end } = readOpenTag(text, pos, startLine, bump);
		tokens.push({ kind: 'open', tag, attrs, selfClosed, line: startLine });
		bump(end);
	}
}

function readOpenTag(
	text: string,
	start: number,
	line: number,
	_bump: (n: number) => void,
): { tag: string; attrs: Map<string, string>; selfClosed: boolean; end: number } {
	let i = start + 1;
	const nameMatch = /^[A-Za-z][\w.]*/.exec(text.slice(i));
	if (!nameMatch) { throw new FormMarkupError(line, 'expected an element name after <'); }
	const tag = nameMatch[0];
	i += tag.length;
	const attrs = new Map<string, string>();
	for (;;) {
		while (i < text.length && /\s/.test(text[i])) { i++; }
		if (i >= text.length) { throw new FormMarkupError(line, `<${tag}> is never closed`); }
		if (text[i] === '>') { return { tag, attrs, selfClosed: false, end: i + 1 }; }
		if (text[i] === '/' && text[i + 1] === '>') { return { tag, attrs, selfClosed: true, end: i + 2 }; }
		const attrMatch = /^[A-Za-z][\w.]*/.exec(text.slice(i));
		if (!attrMatch) { throw new FormMarkupError(line, `cannot read an attribute in <${tag}>`); }
		const name = attrMatch[0];
		i += name.length;
		while (i < text.length && /\s/.test(text[i])) { i++; }
		if (text[i] !== '=') {
			throw new FormMarkupError(line, `${name} needs ="value" - every value is quoted`);
		}
		i++;
		while (i < text.length && /\s/.test(text[i])) { i++; }
		if (text[i] !== '"') {
			throw new FormMarkupError(line, `${name} needs a QUOTED value`);
		}
		const close = text.indexOf('"', i + 1);
		if (close < 0) { throw new FormMarkupError(line, `${name}'s value is never closed`); }
		if (attrs.has(name)) { throw new FormMarkupError(line, `${name} appears twice in <${tag}>`); }
		attrs.set(name, unescapeAttr(text.slice(i + 1, close)));
		i = close + 1;
	}
}

// ------------------------------------------------------------------- apply

export interface ApplyOutcome {
	/** Human-readable operations, in the order they landed. */
	applied: string[];
}

const KIND_TO_CACHE_INDEX: Readonly<Record<string, number>> = {
	Image: 12, Frame: 14, SpinButton: 16, CommandButton: 17, TabStrip: 18,
	Label: 21, TextBox: 23, ListBox: 24, ComboBox: 25, CheckBox: 26,
	OptionButton: 27, ToggleButton: 28, ScrollBar: 47, MultiPage: 57,
};

/**
 * Applies the parsed document to the package, keyed by control name. The
 * document must have parsed whole; a refused operation throws with its line
 * and nothing after it lands (the caller reparses the workbook on failure).
 */
export function applyFormMarkup(pkg: FormPackage, root: MarkupElement): ApplyOutcome {
	const outcome: ApplyOutcome = { applied: [] };
	applyToPackage(pkg, root, outcome, true);
	return outcome;
}

function applyToPackage(
	pkg: FormPackage,
	element: MarkupElement,
	outcome: ApplyOutcome,
	isForm: boolean,
): void {
	// The container's own attributes. A Page's Caption is NOT its FormControl
	// caption - it lives in the parent MultiPage's tab items, where
	// applyToMultiPage routes it - and the FORM's lives in the VBFrame the
	// service owns; only a Frame's caption belongs to this record.
	const record = pkg.form.record;
	const caption = element.tag.toLowerCase() === 'page' || isForm
		? undefined
		: element.attrs.get('Caption');
	if (caption !== undefined && caption !== (record.strings.get('Caption')?.text ?? '')) {
		setRecordString(record, 'Caption', caption);
		outcome.applied.push(`Caption of ${element.attrs.get('Name') ?? element.tag}`);
	}
	if (isForm) {
		const width = element.attrs.get('Width');
		const height = element.attrs.get('Height');
		if (width !== undefined || height !== undefined) {
			const size = record.sizes.get('DisplayedSize');
			if (size) {
				const nextW = width !== undefined ? pointsToHimetric(Number(width)) : size.width;
				const nextH = height !== undefined ? pointsToHimetric(Number(height)) : size.height;
				if (nextW !== size.width || nextH !== size.height) {
					record.sizes.set('DisplayedSize', { width: nextW, height: nextH });
					const logical = record.sizes.get('LogicalSize');
					if (logical) {
						record.sizes.set('LogicalSize', { width: nextW, height: nextH });
					}
					outcome.applied.push('Form size');
				}
			}
		}
	}
	for (const field of ['BackColor', 'ForeColor', 'BorderColor']) {
		const text = element.attrs.get(field);
		if (text === undefined) { continue; }
		const value = parseOleColor(text);
		if (value === undefined) {
			throw new FormMarkupError(element.line, `${field}="${text}" is not a color this dialect knows`);
		}
		if (record.values.get(field) !== value || !recordHas(record, field)) {
			setRecordValue(record, field, value);
			outcome.applied.push(`${field} of ${element.attrs.get('Name') ?? element.tag}`);
		}
	}

	// The children, diffed by name.
	const childElements = element.children.filter((c) => c.tag.toLowerCase() !== 'tab');
	const byName = new Map<string, MarkupElement>();
	for (const child of childElements) {
		const name = child.attrs.get('Name');
		if (!name) {
			throw new FormMarkupError(child.line, `<${child.tag}> has no Name; the name is the key the diff matches by`);
		}
		if (byName.has(name.toLowerCase())) {
			throw new FormMarkupError(child.line, `two controls are named ${name}`);
		}
		byName.set(name.toLowerCase(), child);
	}

	// Removals: in the model, not in the document.
	const keep: typeof pkg.entries = [];
	const keptSites: SiteModel[] = [];
	for (const entry of pkg.entries) {
		const name = siteName(entry.site).toLowerCase();
		if (byName.has(name)) {
			keep.push(entry);
			keptSites.push(entry.site);
			continue;
		}
		if (entry.kind === 'container') { pkg.containers.delete(siteId(entry.site)); }
		outcome.applied.push(`removed ${siteName(entry.site)}`);
		pkg.form.sitesStructurallyChanged = true;
	}
	pkg.entries = keep;
	pkg.form.sites = pkg.form.sites.filter((s) => keptSites.includes(s));

	// Matches and additions, in document order.
	for (const child of childElements) {
		const name = child.attrs.get('Name')!;
		const existing = pkg.entries.find((e) => siteName(e.site).toLowerCase() === name.toLowerCase());
		if (existing) {
			applyToExisting(pkg, existing, child, outcome);
			continue;
		}
		addControl(pkg, child, outcome);
	}
}

function applyToExisting(
	pkg: FormPackage,
	entry: FormPackage['entries'][number],
	element: MarkupElement,
	outcome: ApplyOutcome,
): void {
	const site = entry.site;
	const declaredKind = element.tag;
	const actualKind = controlKindOfSite(site, entry.kind === 'record' ? entry.record : undefined);
	if (entry.kind !== 'raw' && declaredKind.toLowerCase() !== actualKind.toLowerCase()
		&& !(declaredKind === 'ActiveX')) {
		throw new FormMarkupError(
			element.line,
			`${siteName(site)} is a ${actualKind}; changing its type means removing it and adding a new control`,
		);
	}
	applySiteAttrs(site, element, outcome);
	if (entry.kind === 'record') {
		applyRecordAttrs(entry.record, element, actualKind, outcome, siteName(site));
		if (actualKind === 'TabStrip') {
			applyTabCaptions(entry.record, element, outcome, siteName(site));
		}
	}
	if (entry.kind === 'container') {
		const inner = pkg.containers.get(siteId(site));
		if (inner) {
			if (actualKind === 'Frame') {
				applyToPackage(inner, element, outcome, false);
				syncContainerSize(inner, element, site, outcome);
			} else if (actualKind === 'MultiPage') {
				applyToMultiPage(inner, element, outcome, siteName(site));
				syncContainerSize(inner, element, site, outcome);
			}
		}
	}
}

function syncContainerSize(
	inner: FormPackage,
	element: MarkupElement,
	site: SiteModel,
	outcome: ApplyOutcome,
): void {
	const width = element.attrs.get('Width');
	const height = element.attrs.get('Height');
	if (width === undefined && height === undefined) { return; }
	const size = inner.form.record.sizes.get('DisplayedSize');
	if (!size) { return; }
	const nextW = width !== undefined ? pointsToHimetric(Number(width)) : size.width;
	const nextH = height !== undefined ? pointsToHimetric(Number(height)) : size.height;
	if (nextW === size.width && nextH === size.height) { return; }
	inner.form.record.sizes.set('DisplayedSize', { width: nextW, height: nextH });
	const logical = inner.form.record.sizes.get('LogicalSize');
	if (logical) { inner.form.record.sizes.set('LogicalSize', { width: nextW, height: nextH }); }
	outcome.applied.push(`size of ${siteName(site)}`);
}

function applyToMultiPage(
	mp: FormPackage,
	element: MarkupElement,
	outcome: ApplyOutcome,
	mpName: string,
): void {
	const pages = element.children.filter((c) => c.tag.toLowerCase() === 'page');
	const pageSites = mp.form.sites.filter((s) => siteCacheIndex(s) === 7);
	if (pages.length !== pageSites.length) {
		throw new FormMarkupError(
			element.line,
			`${mpName} has ${pageSites.length} pages; adding or removing pages through markup is not supported yet`,
		);
	}
	// Page captions live on the MultiPage's TabStrip Items array.
	const tabStrip = mp.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
	if (tabStrip && tabStrip.kind === 'record') {
		const items = tabStrip.record.arrays.get('Items');
		const captions = items ? decodeArrayStrings(items) : [];
		let changed = false;
		pages.forEach((page, index) => {
			const caption = page.attrs.get('Caption');
			if (caption !== undefined && captions[index] !== undefined && caption !== captions[index]) {
				captions[index] = caption;
				changed = true;
			}
		});
		if (changed) {
			const encoded = encodeArrayStrings(captions);
			tabStrip.record.arrays.set('Items', encoded);
			setRecordValue(tabStrip.record, 'ItemsSize', encoded.length);
			outcome.applied.push(`page captions of ${mpName}`);
		}
	}
	pages.forEach((page, index) => {
		const pageSite = pageSites[index];
		const inner = mp.containers.get(siteId(pageSite));
		if (inner) { applyToPackage(inner, page, outcome, false); }
	});
}

function applySiteAttrs(site: SiteModel, element: MarkupElement, outcome: ApplyOutcome): void {
	const left = element.attrs.get('Left');
	const top = element.attrs.get('Top');
	if (left !== undefined || top !== undefined) {
		const current = site.position ?? { left: 0, top: 0 };
		const next = {
			left: left !== undefined ? pointsToHimetric(Number(left)) : current.left,
			top: top !== undefined ? pointsToHimetric(Number(top)) : current.top,
		};
		if (next.left !== current.left || next.top !== current.top) {
			site.position = next;
			site.mask = (site.mask | (1 << 8)) >>> 0;
			outcome.applied.push(`position of ${siteName(site)}`);
		}
	}
	const tab = element.attrs.get('TabIndex');
	if (tab !== undefined && Number(tab) !== site.values.get('TabIndex')) {
		site.values.set('TabIndex', Number(tab));
		site.mask = (site.mask | (1 << 6)) >>> 0;
		outcome.applied.push(`TabIndex of ${siteName(site)}`);
	}
	const applyString = (attr: string, lenField: string, name: string, bit: number): void => {
		const text = element.attrs.get(attr);
		if (text === undefined) { return; }
		const current = site.strings.get(name)?.text ?? '';
		if (text === current) { return; }
		const compressed = [...text].every((c) => c.charCodeAt(0) <= 0xff);
		site.strings.set(name, { text, compressed, raw: Buffer.alloc(0), edited: true });
		site.values.set(lenField, 0); // refreshed at serialize
		site.mask = (site.mask | (1 << bit)) >>> 0;
		outcome.applied.push(`${attr} of ${siteName(site)}`);
	};
	applyString('ControlTipText', 'ControlTipTextData', 'ControlTipText', 11);
	applyString('Tag', 'TagData', 'Tag', 1);
}

function applyRecordAttrs(
	record: ParsedRecord,
	element: MarkupElement,
	kind: string,
	outcome: ApplyOutcome,
	name: string,
): void {
	const width = element.attrs.get('Width');
	const height = element.attrs.get('Height');
	if (width !== undefined || height !== undefined) {
		const size = record.sizes.get('Size') ?? { width: 0, height: 0 };
		const next = {
			width: width !== undefined ? pointsToHimetric(Number(width)) : size.width,
			height: height !== undefined ? pointsToHimetric(Number(height)) : size.height,
		};
		if (next.width !== size.width || next.height !== size.height) {
			record.sizes.set('Size', next);
			outcome.applied.push(`size of ${name}`);
		}
	}
	for (const [attr, field] of [['Caption', 'Caption'], ['Value', 'Value'], ['GroupName', 'GroupName']] as const) {
		const text = element.attrs.get(attr);
		if (text === undefined) { continue; }
		if (!record.spec.extra.some((f) => f.name === field && f.kind === 'str')) {
			throw new FormMarkupError(element.line, `a ${kind} has no ${attr}`);
		}
		if ((record.strings.get(field)?.text ?? '') === text) { continue; }
		setRecordString(record, field, text);
		outcome.applied.push(`${attr} of ${name}`);
	}
	for (const field of ['BackColor', 'ForeColor', 'BorderColor']) {
		const text = element.attrs.get(field);
		if (text === undefined) { continue; }
		const value = parseOleColor(text);
		if (value === undefined) {
			throw new FormMarkupError(element.line, `${field}="${text}" is not a color this dialect knows`);
		}
		if (!record.spec.data.some((f) => f.name === field)) {
			throw new FormMarkupError(element.line, `a ${kind} has no ${field}`);
		}
		if (record.values.get(field) !== value || !recordHas(record, field)) {
			setRecordValue(record, field, value);
			outcome.applied.push(`${field} of ${name}`);
		}
	}
	for (const field of PRINTED_FIELDS[kind] ?? []) {
		const text = element.attrs.get(field);
		if (text === undefined) { continue; }
		const value = Number(text);
		if (!Number.isFinite(value)) {
			throw new FormMarkupError(element.line, `${field}="${text}" is not a number`);
		}
		if (record.values.get(field) !== value || !recordHas(record, field)) {
			setRecordValue(record, field, value);
			outcome.applied.push(`${field} of ${name}`);
		}
	}
	applyFontAttrs(record, element, outcome, name);
	for (const [attr, field] of [['PasswordChar', 'PasswordChar'], ['Accelerator', 'Accelerator']] as const) {
		const text = element.attrs.get(attr);
		if (text === undefined) { continue; }
		if (!record.spec.data.some((f) => f.name === field)) {
			throw new FormMarkupError(element.line, `a ${kind} has no ${attr}`);
		}
		const value = text.length ? text.charCodeAt(0) : 0;
		if (record.values.get(field) !== value || !recordHas(record, field)) {
			setRecordValue(record, field, value);
			outcome.applied.push(`${attr} of ${name}`);
		}
	}
}

/**
 * Font.* edits land on the control's TextProps: name as its string, size in
 * twips, bold and italic as FontEffects bits with the weight kept in step.
 */
function applyFontAttrs(
	record: ParsedRecord,
	element: MarkupElement,
	outcome: ApplyOutcome,
	name: string,
): void {
	const wantsFont = [...element.attrs.keys()].some((k) => k.startsWith('Font.'));
	if (!wantsFont) { return; }
	const tp = record.textProps;
	if (!tp) {
		throw new FormMarkupError(element.line, `${name}: this control kind carries no font`);
	}
	const fontName = element.attrs.get('Font.Name');
	if (fontName !== undefined && fontName !== (tp.strings.get('FontName')?.text ?? '')) {
		setRecordString(tp, 'FontName', fontName);
		outcome.applied.push(`Font.Name of ${name}`);
	}
	const fontSize = element.attrs.get('Font.Size');
	if (fontSize !== undefined) {
		const twips = Math.round(Number(fontSize) * 20);
		if (!Number.isFinite(twips) || twips <= 0) {
			throw new FormMarkupError(element.line, `Font.Size="${fontSize}" is not a size`);
		}
		if (tp.values.get('FontHeight') !== twips || !recordHas(tp, 'FontHeight')) {
			setRecordValue(tp, 'FontHeight', twips);
			outcome.applied.push(`Font.Size of ${name}`);
		}
	}
	const boldAttr = element.attrs.get('Font.Bold');
	const italicAttr = element.attrs.get('Font.Italic');
	if (boldAttr !== undefined || italicAttr !== undefined) {
		const current = recordHas(tp, 'FontEffects') ? (tp.values.get('FontEffects') ?? 0) : 0;
		const bold = boldAttr !== undefined ? /^true$/i.test(boldAttr) : (current & 0x1) !== 0;
		const italic = italicAttr !== undefined ? /^true$/i.test(italicAttr) : (current & 0x2) !== 0;
		const next = (current & ~0x3) | (bold ? 0x1 : 0) | (italic ? 0x2 : 0);
		if (next !== current || !recordHas(tp, 'FontEffects')) {
			setRecordValue(tp, 'FontEffects', next >>> 0);
			setRecordValue(tp, 'FontWeight', bold ? 700 : 400);
			outcome.applied.push(`Font style of ${name}`);
		}
	}
}

/** Caption edits on a standalone TabStrip's <Tab> children. */
function applyTabCaptions(
	record: ParsedRecord,
	element: MarkupElement,
	outcome: ApplyOutcome,
	name: string,
): void {
	const tabs = element.children.filter((c) => c.tag.toLowerCase() === 'tab');
	if (tabs.length === 0) { return; }
	const items = record.arrays.get('Items');
	const captions = items ? decodeArrayStrings(items) : [];
	if (tabs.length !== captions.length) {
		throw new FormMarkupError(
			element.line,
			`${name} has ${captions.length} tabs; adding or removing tabs through markup is not supported yet`,
		);
	}
	let changed = false;
	tabs.forEach((tab, index) => {
		const caption = tab.attrs.get('Caption');
		if (caption !== undefined && caption !== captions[index]) {
			captions[index] = caption;
			changed = true;
		}
	});
	if (!changed) { return; }
	const encoded = encodeArrayStrings(captions);
	record.arrays.set('Items', encoded);
	setRecordValue(record, 'ItemsSize', encoded.length);
	outcome.applied.push(`tab captions of ${name}`);
}

// ------------------------------------------------------------- additions

function addControl(pkg: FormPackage, element: MarkupElement, outcome: ApplyOutcome): void {
	const kind = element.tag;
	const name = element.attrs.get('Name')!;
	if (kind === 'ActiveX') {
		throw new FormMarkupError(
			element.line,
			`${name}: creating an ActiveX control needs a class table entry this engine does not author yet`,
		);
	}
	if (kind === 'MultiPage' || kind === 'Page' || kind === 'Tab') {
		throw new FormMarkupError(
			element.line,
			`${name}: adding a ${kind} through markup is not supported yet; leaf controls and Frames are`,
		);
	}
	const cacheIndex = KIND_TO_CACHE_INDEX[kind];
	if (cacheIndex === undefined) {
		throw new FormMarkupError(element.line, `<${kind}> is not a control kind this dialect knows`);
	}

	// Excel's own fixture carries NextAvailableID EQUAL to the highest live ID
	// (OkButton holds 20, NextAvailableID says 20), so trusting the field
	// alone reuses a live ID - and duplicate site IDs kill the form at load.
	// The safe rule: above both the field and every existing ID, then record
	// the assignment the way Excel does, as the last ID handed out.
	const formRecord = pkg.form.record;
	let id = formRecord.values.get('NextAvailableID') ?? 1;
	for (const existing of pkg.form.sites) {
		const existingId = existing.values.get('ID');
		if (existingId !== undefined && existingId >= id) { id = existingId + 1; }
	}
	setRecordValue(formRecord, 'NextAvailableID', id);

	const site: SiteModel = {
		mask: 0,
		values: new Map(),
		strings: new Map(),
		pads: new Map(),
	};
	const setSite = (bit: number, field: string, value: number): void => {
		site.mask = (site.mask | (1 << bit)) >>> 0;
		site.values.set(field, value);
	};
	site.mask = (site.mask | (1 << 0)) >>> 0; // Name
	site.strings.set('Name', {
		text: name,
		compressed: [...name].every((c) => c.charCodeAt(0) <= 0xff),
		raw: Buffer.alloc(0),
		edited: true,
	});
	site.values.set('NameData', 0);
	setSite(2, 'ID', id);
	setSite(5, 'ObjectStreamSize', 0);
	setSite(6, 'TabIndex', nextTabIndex(pkg));
	setSite(7, 'ClsidCacheIndex', cacheIndex);
	site.mask = (site.mask | (1 << 8)) >>> 0;
	site.position = {
		left: pointsToHimetric(Number(element.attrs.get('Left') ?? '0')),
		top: pointsToHimetric(Number(element.attrs.get('Top') ?? '0')),
	};

	pkg.form.sites.push(site);
	pkg.form.sitesStructurallyChanged = true;

	if (kind === 'Frame') {
		const inner = newEmptyContainerPackage(pkg, element);
		pkg.containers.set(id, inner);
		pkg.entries.push({ kind: 'container', site });
		applyToPackage(inner, element, outcome, false);
		outcome.applied.push(`added Frame ${name}`);
		return;
	}

	const record = newRecordForKind(kind);
	record.sizes.set('Size', {
		width: pointsToHimetric(Number(element.attrs.get('Width') ?? '72')),
		height: pointsToHimetric(Number(element.attrs.get('Height') ?? '18')),
	});
	pkg.entries.push({ kind: 'record', site, record });
	applyRecordAttrs(record, element, kind, outcome, name);
	outcome.applied.push(`added ${kind} ${name}`);
}

function nextTabIndex(pkg: FormPackage): number {
	let max = -1;
	for (const site of pkg.form.sites) {
		const t = site.values.get('TabIndex');
		if (t !== undefined && t > max) { max = t; }
	}
	return max + 1;
}

function newRecordForKind(kind: string): ParsedRecord {
	const cacheIndex = KIND_TO_CACHE_INDEX[kind];
	const specIndex = cacheIndex === 14 || cacheIndex === 57 ? undefined : cacheIndex;
	const spec = specIndex !== undefined
		? requireSpec(specIndex)
		: undefined;
	if (!spec) { throw new RangeError(`${kind} has no record spec`); }
	const record: ParsedRecord = {
		spec,
		maskLo: 0,
		maskHi: 0,
		values: new Map(),
		strings: new Map(),
		sizes: new Map(),
		arrays: new Map(),
		pads: new Map(),
		streamData: new Map(),
	};
	// fSize MUST be set on every control record.
	const sizeExtra = spec.extra.find((f) => f.kind === 'size8');
	if (sizeExtra) {
		if (sizeExtra.bit < 32) { record.maskLo = (record.maskLo | (1 << sizeExtra.bit)) >>> 0; }
		else { record.maskHi = (record.maskHi | (1 << (sizeExtra.bit - 32))) >>> 0; }
	}
	// The MorphData family carries its concrete type in DisplayStyle.
	if (spec.mask64) {
		const style = ({ TextBox: 1, ListBox: 2, ComboBox: 3, CheckBox: 4, OptionButton: 5, ToggleButton: 6 } as Record<string, number>)[kind];
		if (style !== undefined && style !== 1) {
			setRecordValue(record, 'DisplayStyle', style);
		}
	}
	if (spec.textProps) {
		record.textProps = {
			spec: TEXT_PROPS_REF,
			maskLo: 0, maskHi: 0,
			values: new Map(), strings: new Map(), sizes: new Map(), arrays: new Map(),
			pads: new Map(), streamData: new Map(),
		};
	}
	return record;
}

function requireSpec(cacheIndex: number) {
	const spec = RECORD_SPECS_BY_CACHE_INDEX.get(cacheIndex);
	if (!spec) { throw new RangeError(`no record spec for cache index ${cacheIndex}`); }
	return spec;
}

/** A fresh, empty Frame package: a minimal FormControl and no sites. */
function newEmptyContainerPackage(parent: FormPackage, element: MarkupElement): FormPackage {
	const width = pointsToHimetric(Number(element.attrs.get('Width') ?? '100'));
	const height = pointsToHimetric(Number(element.attrs.get('Height') ?? '80'));
	const record: ParsedRecord = {
		spec: parent.form.record.spec,
		maskLo: 0, maskHi: 0,
		values: new Map(), strings: new Map(), sizes: new Map(), arrays: new Map(),
		pads: new Map(), streamData: new Map(),
	};
	const setBit = (bit: number): void => { record.maskLo = (record.maskLo | (1 << bit)) >>> 0; };
	setBit(3); record.values.set('NextAvailableID', 1);
	setBit(10); record.sizes.set('DisplayedSize', { width, height });
	setBit(11); record.sizes.set('LogicalSize', { width, height });
	setBit(27); record.values.set('DrawBuffer', 16000);
	return {
		form: {
			record,
			classTableRaw: Buffer.alloc(0),
			classTablePresent: false,
			sites: [],
			depthsRaw: Buffer.alloc(0),
			sitesStructurallyChanged: true,
			trailingRaw: Buffer.alloc(0),
		},
		entries: [],
		containers: new Map(),
		compObjRaw: parent.compObjRaw,
	};
}
