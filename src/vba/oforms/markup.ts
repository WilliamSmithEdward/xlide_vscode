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
//   - the control list is total, and so is each spoken control's DIALECT
//     vocabulary: an attribute quiet at its default means the default, so a
//     document apply restores what an edit set back (that is what makes the
//     designer's text undo honest). Fields OUTSIDE the dialect - pictures,
//     unknown records, foreign controls' payloads - are never touched;
//   - apply is a diff keyed by control name: only-in-markup adds,
//     only-in-model removes, matched controls set what changed;
//   - a parse error applies nothing.

import { himetricToPoints, pointsToHimetric, formatPointsShortest as formatPoints, OformsReader, OformsWriter } from './bytes';
import {
	recordHas,
	setRecordString,
	setRecordValue,
	RECORD_SPECS_BY_CACHE_INDEX,
	TEXT_PROPS_SPEC as TEXT_PROPS_REF,
	type ParsedRecord,
} from './records';
import {
	composeStdFont,
	parseStdFont,
	siteName,
	siteId,
	siteCacheIndex,
	siteIsContainer,
	type SiteModel,
} from './formStream';
import { containerStorageName, controlKindOfSite, type FormPackage } from './formPackage';
import { parsePageBookkeeping, serializePageBookkeeping, emptyPageProperties } from './pageBookkeeping';
import { PAGE_COMPOBJ, FRAME_COMPOBJ } from './newForm';

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
export const PRINTED_FIELDS: Readonly<Record<string, readonly string[]>> = {
	TextBox: ['MaxLength', 'BorderStyle', 'SpecialEffect', 'ScrollBars', 'MousePointer'],
	ComboBox: ['BorderStyle', 'SpecialEffect', 'ListRows', 'ColumnCount', 'ListStyle', 'MousePointer'],
	ListBox: ['BorderStyle', 'SpecialEffect', 'MultiSelect', 'ColumnCount', 'ListStyle', 'MousePointer'],
	CheckBox: ['SpecialEffect', 'MousePointer'],
	OptionButton: ['SpecialEffect', 'MousePointer'],
	ToggleButton: ['MousePointer'],
	Label: ['BorderStyle', 'SpecialEffect', 'MousePointer'],
	CommandButton: ['MousePointer'],
	Image: ['BorderStyle', 'SpecialEffect', 'PictureSizeMode', 'PictureAlignment', 'MousePointer'],
	SpinButton: ['Min', 'Max', 'Position', 'SmallChange', 'Orientation', 'Delay', 'MousePointer'],
	ScrollBar: ['Min', 'Max', 'Position', 'SmallChange', 'LargeChange', 'Orientation', 'ProportionalThumb', 'Delay', 'MousePointer'],
	TabStrip: ['TabOrientation', 'TabStyle', 'MousePointer'],
	MultiPage: [],
	Frame: [],
	Page: [],
};

// Alignment ([MS-OFORMS] VariousPropertiesBitfield bit 13) says which side a
// CheckBox or OptionButton caption sits on: the bit set puts the caption to
// the LEFT, which VBA spells Alignment 0; clear is the default 1, right.
const ALIGNMENT_BIT = 0x2000;
export const ALIGNMENT_KINDS: ReadonlySet<string> = new Set(['CheckBox', 'OptionButton']);

const COLOR_FIELDS = new Set(['BackColor', 'ForeColor', 'BorderColor']);

// VariousPropertyBits file-format defaults, per kind ([MS-OFORMS] 2.5.96) -
// what an ABSENT field means, so flag edits compose without touching any
// other property. TextBox and ComboBox carry Editable (bit 14) on top of the
// table's MorphData default: the spec makes it MUST-be-1 for a TextBox, and
// live Excel stores 0x2C80481B.
const VPB_DEFAULTS: Readonly<Record<string, number>> = {
	TextBox: 0x2C80481B, ComboBox: 0x2C80481B, ListBox: 0x2C80081B, CheckBox: 0x2C80081B,
	OptionButton: 0x2C80081B, ToggleButton: 0x2C80081B, Label: 0x0080001B,
	CommandButton: 0x1B, Image: 0x1B, TabStrip: 0x1B, ScrollBar: 0x1B, SpinButton: 0x1B,
};

// The named booleans the dialect speaks from that bitfield, with the kinds
// each applies to ([MS-OFORMS] 2.5.96.1). A flag prints only when it differs
// from the kind's default bit, so quiet forms stay quiet.
export const VPB_FLAGS: ReadonlyArray<readonly [string, number, readonly string[]]> = [
	['Enabled', 0x2, ['CheckBox', 'ComboBox', 'CommandButton', 'Image', 'Label', 'ListBox', 'OptionButton', 'ScrollBar', 'SpinButton', 'TabStrip', 'TextBox', 'ToggleButton']],
	['Locked', 0x4, ['CheckBox', 'ComboBox', 'CommandButton', 'ListBox', 'OptionButton', 'TextBox', 'ToggleButton']],
	['WordWrap', 0x800000, ['CheckBox', 'ComboBox', 'CommandButton', 'Label', 'ListBox', 'OptionButton', 'TextBox', 'ToggleButton']],
	['AutoSize', 0x10000000, ['CheckBox', 'ComboBox', 'CommandButton', 'Label', 'OptionButton', 'TextBox', 'ToggleButton']],
	['MultiLine', 0x80000000, ['TextBox']],
];

// PARAFORMAT_Alignment ([MS-OFORMS] 2.5.60.1). VBA's TextAlign enum numbers
// the same three positions DIFFERENTLY (center and right swap), so the
// dialect speaks words and never a bare number.
const TEXT_ALIGN_WORDS: Readonly<Record<number, string>> = { 1: 'Left', 2: 'Right', 3: 'Center' };
const TEXT_ALIGN_VALUES: Readonly<Record<string, number>> = { left: 1, right: 2, center: 3 };

// The kinds whose VBA surface exposes TextAlign. A CommandButton STORES a
// ParagraphAlign (Excel writes 3, centered) but has no TextAlign property,
// so the dialect stays silent about it there.
export const TEXT_ALIGN_KINDS: ReadonlySet<string> = new Set([
	'Label', 'TextBox', 'ComboBox', 'ListBox', 'CheckBox', 'OptionButton', 'ToggleButton',
]);

/** The stored-or-default VariousPropertyBits a control answers with. */
export function effectiveVariousPropertyBits(record: ParsedRecord, kind: string): number | undefined {
	const fallback = VPB_DEFAULTS[kind];
	if (fallback === undefined) { return undefined; }
	return recordHas(record, 'VariousPropertyBits')
		? (record.values.get('VariousPropertyBits') ?? fallback) >>> 0
		: fallback;
}

// Site BitFlags ([MS-OFORMS] 2.5.4): default 0x33 is fTabStop + fVisible +
// fStreamed + fAutoSize. TabStop is not a Label or Image property, Default
// and Cancel belong to CommandButton, and a Page's Visible bit is the
// CURRENT-page marker - never a property to spell.
export const SITE_BITFLAGS_DEFAULT = 0x33;
export const SITE_FLAGS: ReadonlyArray<readonly [string, number]> = [
	['TabStop', 0x1],
	['Visible', 0x2],
	['Default', 0x4],
	['Cancel', 0x8],
];

function siteFlagApplies(attr: string, kind: string | undefined): boolean {
	if (kind === undefined || kind === 'Form' || kind === 'Page') { return false; }
	if (attr === 'TabStop') { return kind !== 'Label' && kind !== 'Image'; }
	if (attr === 'Default' || attr === 'Cancel') { return kind === 'CommandButton'; }
	return true;
}

/** The VBFrame-persisted form properties the dialect spells on <Form>. */
export interface VbFrameMarkupProps {
	showModal?: string;
	startUpPosition?: string;
	whatsThisButton?: string;
}

/** The form record's own extra numerics the dialect prints and applies. */
export const FORM_EXTRA_FIELDS = ['BorderStyle', 'ScrollBars', 'Cycle', 'Zoom', 'MousePointer'] as const;

/**
 * What VBA accepts as a control name: a LETTER first (any script - Japanese
 * control names are legal VBA), then letters, digits, and underscores, at
 * most 255 characters. Enforced when a control is CREATED (an addition or a
 * rename); a control an existing workbook already carries is matched as-is.
 * A name outside this can never have an event handler, so accepting one
 * authors a form the user cannot wire up.
 */
export const LEGAL_CONTROL_NAME = /^\p{L}[\p{L}\p{Nd}_]{0,254}$/u;

export function printFormMarkup(
	pkg: FormPackage,
	formName: string,
	options: { captionFallback?: string; vbFrame?: VbFrameMarkupProps } = {},
): string {
	const lines: string[] = [];
	printPackage(pkg, formName, lines, 0, 'Form', undefined, options.captionFallback, options.vbFrame);
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
	vbFrame?: VbFrameMarkupProps,
): void {
	const indent = '    '.repeat(depth);
	const attrs: string[] = [`Name="${escapeAttr(name)}"`];
	const record = pkg.form.record;
	// A top-level form's caption is persisted in the VBFrame text, not in the
	// f record - the service passes it in; a Frame's caption IS in its record.
	// A FRAME'S empty caption stays unspoken (apply reads absence as empty);
	// the FORM'S prints even empty - it diffs against the VBFrame's line.
	const caption = record.strings.get('Caption');
	if (caption && caption.text.length) { attrs.push(`Caption="${escapeAttr(caption.text)}"`); }
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
	if (tag === 'Form') {
		// The form's own extras and its StdFont, quiet at their defaults -
		// the document must carry EVERYTHING a save is expected to keep.
		for (const field of FORM_EXTRA_FIELDS) {
			const v = record.values.get(field);
			if (v !== undefined && recordHas(record, field)) {
				attrs.push(`${field}="${v}"`);
			}
		}
		const font = pkg.form.fontRaw ? parseStdFont(pkg.form.fontRaw) : undefined;
		if (font) {
			attrs.push(`Font.Name="${escapeAttr(font.face)}"`);
			attrs.push(`Font.Size="${font.heightTenThousandthsPt / 10000}"`);
			if ((font.flags & 0x1) !== 0 || font.weight >= 600) { attrs.push('Font.Bold="True"'); }
			if ((font.flags & 0x2) !== 0) { attrs.push('Font.Italic="True"'); }
			if ((font.flags & 0x4) !== 0) { attrs.push('Font.Underline="True"'); }
			if ((font.flags & 0x8) !== 0) { attrs.push('Font.Strikethrough="True"'); }
		}
		if (vbFrame) {
			if (vbFrame.startUpPosition !== undefined && vbFrame.startUpPosition !== '1') {
				attrs.push(`StartUpPosition="${escapeAttr(vbFrame.startUpPosition)}"`);
			}
			if (vbFrame.showModal === 'False') { attrs.push('ShowModal="False"'); }
			if (vbFrame.whatsThisButton === 'True') { attrs.push('WhatsThisButton="True"'); }
		}
	}
	pushSiteExtras(attrs, site, tag);
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
		// Empty strings stay UNSPOKEN: the dialect cannot re-create a stored
		// empty on apply, so printing one would make the document say more
		// than a save could keep (the fuzz oracle's find, 2026-08-29).
		const caption = record.strings.get('Caption');
		if (caption && caption.text.length) { attrs.push(`Caption="${escapeAttr(caption.text)}"`); }
		const value = record.strings.get('Value');
		if (value && value.text.length) { attrs.push(`Value="${escapeAttr(value.text)}"`); }
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
		const pa = record.textProps?.values.get('ParagraphAlign');
		if (TEXT_ALIGN_KINDS.has(kind) && pa !== undefined && record.textProps
			&& recordHas(record.textProps, 'ParagraphAlign') && pa !== 1 && TEXT_ALIGN_WORDS[pa]) {
			attrs.push(`TextAlign="${TEXT_ALIGN_WORDS[pa]}"`);
		}
		if (kind === 'ComboBox' && recordHas(record, 'DisplayStyle')
			&& record.values.get('DisplayStyle') === 7) {
			// fmDisplayStyleDropList - VBA spells it Style 2 (fmStyleDropDownList).
			attrs.push('Style="2"');
		}
		const effectiveVpb = effectiveVariousPropertyBits(record, kind);
		if (effectiveVpb !== undefined && ALIGNMENT_KINDS.has(kind) && (effectiveVpb & ALIGNMENT_BIT) !== 0) {
			attrs.push('Alignment="0"');
		}
		if (effectiveVpb !== undefined) {
			const fallback = VPB_DEFAULTS[kind];
			for (const [attr, bit, kinds] of VPB_FLAGS) {
				if (!kinds.includes(kind)) { continue; }
				if (((effectiveVpb & bit) >>> 0) !== ((fallback & bit) >>> 0)) {
					attrs.push(`${attr}="${(effectiveVpb & bit) !== 0 ? 'True' : 'False'}"`);
				}
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
	pushSiteExtras(attrs, site, kind);
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
	pushSiteExtras(attrs, site, 'MultiPage');
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

function pushSiteExtras(attrs: string[], site: SiteModel | undefined, kind?: string): void {
	if (!site) { return; }
	const tab = site.values.get('TabIndex');
	if (tab !== undefined && tab >= 0) { attrs.push(`TabIndex="${tab}"`); }
	const tip = site.strings.get('ControlTipText');
	if (tip && tip.text.length) { attrs.push(`ControlTipText="${escapeAttr(tip.text)}"`); }
	const tag = site.strings.get('Tag');
	if (tag && tag.text.length) { attrs.push(`Tag="${escapeAttr(tag.text)}"`); }
	const source = site.strings.get('ControlSource');
	if (source && source.text.length) { attrs.push(`ControlSource="${escapeAttr(source.text)}"`); }
	const rowSource = site.strings.get('RowSource');
	if (rowSource && rowSource.text.length) { attrs.push(`RowSource="${escapeAttr(rowSource.text)}"`); }
	const help = site.values.get('HelpContextID');
	if (help !== undefined && help !== 0) { attrs.push(`HelpContextID="${help}"`); }
	const bits = (site.values.get('BitFlags') ?? SITE_BITFLAGS_DEFAULT) >>> 0;
	for (const [attr, bit] of SITE_FLAGS) {
		if (!siteFlagApplies(attr, kind)) { continue; }
		if (((bits & bit) >>> 0) !== ((SITE_BITFLAGS_DEFAULT & bit) >>> 0)) {
			attrs.push(`${attr}="${(bits & bit) !== 0 ? 'True' : 'False'}"`);
		}
	}
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
		if (effects & 0x4) { attrs.push('Font.Underline="True"'); }
		if (effects & 0x8) { attrs.push('Font.Strikethrough="True"'); }
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
	applyToPackage(pkg, root, outcome, true, pkg);
	return outcome;
}

function applyToPackage(
	pkg: FormPackage,
	element: MarkupElement,
	outcome: ApplyOutcome,
	isForm: boolean,
	root: FormPackage,
): void {
	// The container's own attributes. A Page's Caption is NOT its FormControl
	// caption - it lives in the parent MultiPage's tab items, where
	// applyToMultiPage routes it - and the FORM's lives in the VBFrame the
	// service owns; only a Frame's caption belongs to this record.
	const record = pkg.form.record;
	// A Frame's absent Caption means EMPTY - the printer stays quiet on an
	// empty legend, so a cleared one must clear here too.
	const caption = element.tag.toLowerCase() === 'page' || isForm
		? undefined
		: (element.attrs.get('Caption') ?? '');
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
	const owner = element.attrs.get('Name') ?? element.tag;
	const numericFields = isForm
		? ['SpecialEffect', ...FORM_EXTRA_FIELDS]
		: ['SpecialEffect'];
	for (const field of numericFields) {
		const text = element.attrs.get(field);
		if (text === undefined || !record.spec.data.some((f) => f.name === field)) { continue; }
		const value = Number(text);
		if (!Number.isFinite(value)) {
			throw new FormMarkupError(element.line, `${field}="${text}" is not a number`);
		}
		if (record.values.get(field) !== value || !recordHas(record, field)) {
			setRecordValue(record, field, value >>> 0);
			outcome.applied.push(`${field} of ${owner}`);
		}
	}
	if (isForm) { applyFormFontAttrs(pkg, element, outcome); }

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
			applyToExisting(pkg, existing, child, outcome, root);
			continue;
		}
		addControl(pkg, child, outcome, root);
	}

	// The SIBLING ORDER re-syncs to the document once everyone exists: a
	// reconcile-executed move lands before a diff-phase addition whatever
	// the document interleaved, and the document's order is the truth.
	const orderOf = new Map(childElements.map((c, i) => [c.attrs.get('Name')!.toLowerCase(), i]));
	const sorted = [...pkg.entries].sort((a, b) =>
		(orderOf.get(siteName(a.site).toLowerCase()) ?? 0) - (orderOf.get(siteName(b.site).toLowerCase()) ?? 0));
	if (sorted.some((e, i) => e !== pkg.entries[i])) {
		pkg.entries = sorted;
		pkg.form.sites = sorted.map((e) => e.site);
		pkg.form.sitesStructurallyChanged = true;
		outcome.applied.push(`sibling order of ${element.attrs.get('Name') ?? element.tag}`);
	}
}

function applyToExisting(
	pkg: FormPackage,
	entry: FormPackage['entries'][number],
	element: MarkupElement,
	outcome: ApplyOutcome,
	root: FormPackage,
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
	applySiteAttrs(site, element, outcome, true);
	if (entry.kind === 'record') {
		applyRecordAttrs(entry.record, element, actualKind, outcome, siteName(site), true);
		if (actualKind === 'TabStrip') {
			applyTabCaptions(entry.record, element, outcome, siteName(site));
		}
	}
	if (entry.kind === 'container') {
		const inner = pkg.containers.get(siteId(site));
		if (inner) {
			if (actualKind === 'Frame') {
				applyToPackage(inner, element, outcome, false, root);
				syncContainerSize(inner, element, site, outcome);
			} else if (actualKind === 'MultiPage') {
				applyToMultiPage(inner, element, outcome, siteName(site), root);
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
	root: FormPackage,
): void {
	const pages = element.children.filter((c) => c.tag.toLowerCase() === 'page');
	const tabStripEntry = mp.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
	const tabStrip = tabStripEntry && tabStripEntry.kind === 'record' ? tabStripEntry.record : undefined;

	// Pages HAVE names, so the diff is by name: only-in-document adds,
	// only-in-designer removes. The surviving pages must keep their relative
	// order - reordering moves the x bookkeeping, the tab arrays, and the
	// site list at once, and is refused until it is proven.
	const byName = new Map<string, MarkupElement>();
	for (const page of pages) {
		const pageName = page.attrs.get('Name');
		if (!pageName) {
			throw new FormMarkupError(page.line, '<Page> has no Name; the name is the key the diff matches by');
		}
		if (byName.has(pageName.toLowerCase())) {
			throw new FormMarkupError(page.line, `two pages are named ${pageName}`);
		}
		byName.set(pageName.toLowerCase(), page);
	}

	// A RENAMED page pairs before the diff can read it as remove-plus-add,
	// which rebuilt its contents from text alone and killed anything the
	// dialect cannot spell (hunt nine: a picture on a renamed page died on
	// save while the reprint stayed equal). The pairing is the unambiguous
	// single rename at the SAME position; anything else falls back honestly.
	const currentSites = (): SiteModel[] => mp.form.sites.filter((s) => siteCacheIndex(s) === 7);
	{
		const modelPages = currentSites();
		const removedPages = modelPages.filter((s) => !byName.has(siteName(s).toLowerCase()));
		const addedPages = pages.filter((p) =>
			!modelPages.some((s) => siteName(s).toLowerCase() === p.attrs.get('Name')!.toLowerCase()));
		if (removedPages.length === 1 && addedPages.length === 1) {
			const site = removedPages[0];
			const el = addedPages[0];
			const newName = el.attrs.get('Name')!;
			// Caption gate, as for controls: a rename keeps its caption (the
			// tab Items array is untouched); a fresh page brings its own.
			const captions = tabStrip ? decodeArrayStrings(tabStrip.arrays.get('Items') ?? Buffer.alloc(0)) : [];
			const modelIdx = modelPages.indexOf(site);
			const captionMatches = (el.attrs.get('Caption') ?? '') === (captions[modelIdx] ?? '');
			if (modelIdx === pages.indexOf(el) && captionMatches && LEGAL_CONTROL_NAME.test(newName)) {
				const oldName = siteName(site);
				renameSiteInPlace(site, newName);
				outcome.applied.push(`renamed page ${oldName} to ${newName} of ${mpName}`);
			}
		}
	}

	// Removals, from the end so indices stay stable.
	for (let index = currentSites().length - 1; index >= 0; index--) {
		const site = currentSites()[index];
		if (byName.has(siteName(site).toLowerCase())) { continue; }
		removePage(mp, site, index, tabStrip);
		outcome.applied.push(`removed page ${siteName(site)} of ${mpName}`);
	}

	// Survivors take the DOCUMENT's order. The reorder moves the page
	// sites, the tab arrays, the x bookkeeping, and the selected-page index
	// as one permutation, so the page a tab names, the storage it binds,
	// and the caption it wears travel together.
	const survivorOrder = currentSites().map((s) => siteName(s).toLowerCase());
	const documentOrder = pages
		.map((page) => page.attrs.get('Name')!.toLowerCase())
		.filter((n) => survivorOrder.includes(n));
	if (survivorOrder.join('|') !== documentOrder.join('|')) {
		reorderPages(mp, documentOrder, tabStrip);
		outcome.applied.push(`page order of ${mpName}`);
	}

	// Additions, at their document positions.
	pages.forEach((page, index) => {
		const pageName = page.attrs.get('Name')!;
		if (currentSites().some((s) => siteName(s).toLowerCase() === pageName.toLowerCase())) { return; }
		if (!LEGAL_CONTROL_NAME.test(pageName)) {
			throw new FormMarkupError(
				page.line,
				`${pageName}: not a legal page name - VBA wants a letter first, then letters, digits, or underscores`,
			);
		}
		addPage(mp, page, Math.min(index, currentSites().length), tabStrip, root);
		outcome.applied.push(`added page ${pageName} of ${mpName}`);
	});

	// Captions, positional against the final page order.
	if (tabStrip) {
		const items = tabStrip.arrays.get('Items');
		const captions = items ? decodeArrayStrings(items) : [];
		let changed = false;
		const finalSites = currentSites();
		pages.forEach((page) => {
			const caption = page.attrs.get('Caption');
			if (caption === undefined) { return; }
			const index = finalSites.findIndex(
				(s) => siteName(s).toLowerCase() === page.attrs.get('Name')!.toLowerCase(),
			);
			if (index >= 0 && captions[index] !== undefined && captions[index] !== caption) {
				captions[index] = caption;
				changed = true;
			}
		});
		if (changed) {
			const encoded = encodeArrayStrings(captions);
			tabStrip.arrays.set('Items', encoded);
			setRecordValue(tabStrip, 'ItemsSize', encoded.length);
			outcome.applied.push(`page captions of ${mpName}`);
		}
	}

	// The pages' own contents.
	for (const page of pages) {
		const site = currentSites().find(
			(s) => siteName(s).toLowerCase() === page.attrs.get('Name')!.toLowerCase(),
		);
		if (!site) { continue; }
		const inner = mp.containers.get(siteId(site));
		if (inner) { applyToPackage(inner, page, outcome, false, root); }
	}
}

/** Removes one page: its site, its storage package, its tab entry, its x row. */
function removePage(
	mp: FormPackage,
	site: SiteModel,
	index: number,
	tabStrip: ParsedRecord | undefined,
): void {
	mp.form.sites = mp.form.sites.filter((s) => s !== site);
	mp.entries = mp.entries.filter((e) => e.site !== site);
	mp.containers.delete(siteId(site));
	mp.form.sitesStructurallyChanged = true;
	if (tabStrip) { removeTabEntry(tabStrip, index); }
	if (mp.xRaw) {
		const book = parsePageBookkeeping(mp.xRaw);
		book.pageProps.splice(Math.min(index + 1, book.pageProps.length - 1), 1);
		book.pageIds.splice(index, 1);
		book.pageCount -= 1;
		mp.xRaw = serializePageBookkeeping(book);
	}
}

/**
 * Puts the surviving pages into `order` (lower-cased names). Everything a
 * page's position touches moves in one permutation: the page sites (in
 * place - the TabStrip site and any other neighbors stay put), each page's
 * TabIndex (position-tracking, the value set preserved), the tab strip's
 * per-tab arrays with their flags tail, the selected-tab ListIndex
 * (remapped so the same PAGE stays current, as the VBE keeps it), and the
 * x bookkeeping's PageProperties rows and PageIDs.
 */
function reorderPages(
	mp: FormPackage,
	order: readonly string[],
	tabStrip: ParsedRecord | undefined,
): void {
	const pageSites = mp.form.sites.filter((s) => siteCacheIndex(s) === 7);
	const perm = order.map((name) =>
		pageSites.findIndex((s) => siteName(s).toLowerCase() === name));
	const desired = perm.map((i) => pageSites[i]);

	if (desired.every((s) => s.values.get('TabIndex') !== undefined)) {
		const tabValues = desired
			.map((s) => s.values.get('TabIndex') as number)
			.sort((a, b) => a - b);
		desired.forEach((site, index) => { site.values.set('TabIndex', tabValues[index]); });
	}

	let slot = 0;
	mp.form.sites = mp.form.sites.map((s) => (siteCacheIndex(s) === 7 ? desired[slot++] : s));
	mp.form.sitesStructurallyChanged = true;

	if (tabStrip) {
		for (const [arrayName, sizeField] of TAB_PARALLEL_ARRAYS) {
			const raw = tabStrip.arrays.get(arrayName);
			if (raw === undefined) { continue; }
			const entries = decodeArrayStrings(raw);
			const encoded = encodeArrayStrings(perm.map((i) => entries[i] ?? ''));
			tabStrip.arrays.set(arrayName, encoded);
			setRecordValue(tabStrip, sizeField, encoded.length);
		}
		if (tabStrip.values.get('TabData') !== undefined && recordHas(tabStrip, 'TabData')) {
			const flags = tabStrip.tailRaw ?? Buffer.alloc(0);
			const chunks: Buffer[] = [];
			for (let i = 0; i + 4 <= flags.length; i += 4) { chunks.push(flags.subarray(i, i + 4)); }
			tabStrip.tailRaw = Buffer.concat(perm.map((i) => chunks[i] ?? Buffer.alloc(4)));
		}
		const selected = tabStrip.values.get('ListIndex');
		if (selected !== undefined && recordHas(tabStrip, 'ListIndex')) {
			const moved = perm.indexOf(selected);
			if (moved >= 0 && moved !== selected) {
				setRecordValue(tabStrip, 'ListIndex', moved);
			}
		}
	}

	if (mp.xRaw) {
		const book = parsePageBookkeeping(mp.xRaw);
		const props = book.pageProps;
		book.pageProps = [props[0], ...perm.map((i) => props[i + 1] ?? emptyPageProperties())];
		book.pageIds = perm.map((i) => book.pageIds[i]);
		mp.xRaw = serializePageBookkeeping(book);
	}
}

/** Adds one page: a fresh site, an empty storage package, arrays, x row. */
function addPage(
	mp: FormPackage,
	element: MarkupElement,
	index: number,
	tabStrip: ParsedRecord | undefined,
	root: FormPackage,
): void {
	const pageName = element.attrs.get('Name')!;
	// The page's site ID names its iNN storage, drawn from the same
	// tree-global pool as every control ID.
	const id = allocateControlId(root, mp);

	const reference = mp.form.sites.find((s) => siteCacheIndex(s) === 7);
	const site: SiteModel = { mask: 0, values: new Map(), strings: new Map(), pads: new Map() };
	site.mask = ((1 << 0) | (1 << 2) | (1 << 4) | (1 << 6) | (1 << 7) | (1 << 8)) >>> 0;
	site.strings.set('Name', {
		text: pageName,
		compressed: [...pageName].every((c) => c.charCodeAt(0) <= 0xff),
		raw: Buffer.alloc(0),
		edited: true,
	});
	site.values.set('NameData', 0);
	site.values.set('ID', id);
	// 0x00040021 is what Excel writes for a page that is NOT the current
	// one; the current page differs by one bit, and copying it from an
	// existing page would leave two pages claiming to be current.
	site.values.set('BitFlags', 0x00040021);
	site.values.set('TabIndex', mp.form.sites.filter((s) => siteCacheIndex(s) === 7).length + 1);
	site.values.set('ClsidCacheIndex', 7);
	site.position = reference?.position ? { ...reference.position } : { left: 0, top: 0 };

	// Page sites sit after the TabStrip site, in page order.
	const pageSites = mp.form.sites.filter((s) => siteCacheIndex(s) === 7);
	const anchor = index < pageSites.length
		? mp.form.sites.indexOf(pageSites[index])
		: mp.form.sites.length;
	mp.form.sites.splice(anchor, 0, site);
	mp.entries.push({ kind: 'container', site });
	mp.form.sitesStructurallyChanged = true;

	const inner = newEmptyContainerPackage(mp, element, 'Page');
	mp.containers.set(id, inner);

	if (tabStrip) {
		insertTabEntry(tabStrip, index, element.attrs.get('Caption') ?? pageName, pageName);
	}
	if (mp.xRaw) {
		const book = parsePageBookkeeping(mp.xRaw);
		book.pageProps.splice(index + 1, 0, emptyPageProperties());
		book.pageIds.splice(index, 0, id);
		book.pageCount += 1;
		mp.xRaw = serializePageBookkeeping(book);
	}
}

/**
 * Applies one element's site-level attrs. With `total`, the element speaks
 * for its WHOLE printed vocabulary: an absent value-gated attr (a flag, a
 * source, a tip, a help id) means its DEFAULT, so a document produced by the
 * printer round-trips edits that returned to a default. Without it (the
 * pane's single-property writes), an unspoken attr is never touched.
 */
export function applySiteAttrs(
	site: SiteModel,
	element: MarkupElement,
	outcome: ApplyOutcome,
	total = false,
): void {
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
		const text = element.attrs.get(attr) ?? (total ? '' : undefined);
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
	applyString('ControlSource', 'ControlSourceData', 'ControlSource', 13);
	applyString('RowSource', 'RowSourceData', 'RowSource', 14);
	const help = element.attrs.get('HelpContextID') ?? (total ? '0' : undefined);
	if (help !== undefined) {
		const v = Number(help);
		if (!Number.isFinite(v)) {
			throw new FormMarkupError(element.line, `HelpContextID="${help}" is not a number`);
		}
		const current = site.values.get('HelpContextID') ?? 0;
		if (v !== current) {
			site.values.set('HelpContextID', v);
			site.mask = (site.mask | (1 << 3)) >>> 0;
			outcome.applied.push(`HelpContextID of ${siteName(site)}`);
		}
	}
	for (const [attr, bit] of SITE_FLAGS) {
		const explicit = element.attrs.get(attr);
		const filled = explicit === undefined && total && siteFlagApplies(attr, element.tag);
		const text = explicit
			?? (filled ? (((SITE_BITFLAGS_DEFAULT & bit) >>> 0) !== 0 ? 'True' : 'False') : undefined);
		if (text === undefined) { continue; }
		if (!siteFlagApplies(attr, element.tag)) {
			throw new FormMarkupError(element.line, `a ${element.tag} has no ${attr}`);
		}
		if (!/^(true|false)$/i.test(text)) {
			throw new FormMarkupError(element.line, `${attr}="${text}" is not True or False`);
		}
		const stored = site.values.get('BitFlags');
		const base = (stored ?? SITE_BITFLAGS_DEFAULT) >>> 0;
		const next = (/^true$/i.test(text) ? (base | bit) : (base & ~bit)) >>> 0;
		// A default filled in for an unspoken attr must not materialize the
		// field; only an EXPLICIT spelling does that.
		if (next !== base || (stored === undefined && explicit !== undefined)) {
			site.values.set('BitFlags', next);
			site.mask = (site.mask | (1 << 4)) >>> 0;
			outcome.applied.push(`${attr} of ${siteName(site)}`);
		}
	}
}

export function applyRecordAttrs(
	record: ParsedRecord,
	element: MarkupElement,
	kind: string,
	outcome: ApplyOutcome,
	name: string,
	total = false,
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
		const explicit = element.attrs.get(attr);
		const canCarry = record.spec.extra.some((f) => f.name === field && f.kind === 'str');
		// TOTAL semantics: the printer stays quiet on an empty string, so an
		// absent attr in a whole-document apply means empty - a cleared
		// caption must not resurrect on save.
		const text = explicit ?? (total && canCarry ? '' : undefined);
		if (text === undefined) { continue; }
		if (!canCarry) {
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
	const canAlignText = TEXT_ALIGN_KINDS.has(kind) && record.textProps
		&& record.textProps.spec.data.some((f) => f.name === 'ParagraphAlign');
	const textAlign = element.attrs.get('TextAlign') ?? (total && canAlignText ? 'Left' : undefined);
	if (textAlign !== undefined) {
		const tp = record.textProps;
		if (!canAlignText || !tp) {
			throw new FormMarkupError(element.line, `a ${kind} has no TextAlign`);
		}
		const value = TEXT_ALIGN_VALUES[textAlign.toLowerCase()];
		if (value === undefined) {
			throw new FormMarkupError(element.line, `TextAlign="${textAlign}" is not Left, Center, or Right`);
		}
		const current = recordHas(tp, 'ParagraphAlign') ? tp.values.get('ParagraphAlign') : 1;
		if (current !== value || (!recordHas(tp, 'ParagraphAlign') && element.attrs.has('TextAlign'))) {
			setRecordValue(tp, 'ParagraphAlign', value);
			outcome.applied.push(`TextAlign of ${name}`);
		}
	}
	const style = element.attrs.get('Style') ?? (total && kind === 'ComboBox' ? '0' : undefined);
	if (style !== undefined) {
		if (kind !== 'ComboBox') {
			throw new FormMarkupError(element.line, `a ${kind} has no Style`);
		}
		// VBA Style 0 (fmStyleDropDownCombo) is DisplayStyle 3; Style 2
		// (fmStyleDropDownList) is DisplayStyle 7. There is no Style 1.
		const display = style === '0' ? 3 : style === '2' ? 7 : undefined;
		if (display === undefined) {
			throw new FormMarkupError(element.line, `Style="${style}" is not 0 or 2`);
		}
		if (record.values.get('DisplayStyle') !== display) {
			setRecordValue(record, 'DisplayStyle', display);
			outcome.applied.push(`Style of ${name}`);
		}
	}
	const alignment = element.attrs.get('Alignment') ?? (total && ALIGNMENT_KINDS.has(kind) ? '1' : undefined);
	if (alignment !== undefined) {
		if (!ALIGNMENT_KINDS.has(kind)) {
			throw new FormMarkupError(element.line, `a ${kind} has no Alignment`);
		}
		if (!/^[01]$/.test(alignment)) {
			throw new FormMarkupError(element.line, `Alignment="${alignment}" is not 0 or 1`);
		}
		const base = effectiveVariousPropertyBits(record, kind) ?? 0x1B;
		const next = (alignment === '0' ? (base | ALIGNMENT_BIT) : (base & ~ALIGNMENT_BIT)) >>> 0;
		if (next !== base || (!recordHas(record, 'VariousPropertyBits') && element.attrs.has('Alignment'))) {
			setRecordValue(record, 'VariousPropertyBits', next);
			outcome.applied.push(`Alignment of ${name}`);
		}
	}
	for (const [attr, bit, kinds] of VPB_FLAGS) {
		const explicit = element.attrs.get(attr);
		const canFlag = kinds.includes(kind) && record.spec.data.some((f) => f.name === 'VariousPropertyBits');
		const fallback = VPB_DEFAULTS[kind];
		const text = explicit
			?? (total && canFlag && fallback !== undefined
				? (((fallback & bit) >>> 0) !== 0 ? 'True' : 'False')
				: undefined);
		if (text === undefined) { continue; }
		if (!canFlag) {
			throw new FormMarkupError(element.line, `a ${kind} has no ${attr}`);
		}
		if (!/^(true|false)$/i.test(text)) {
			throw new FormMarkupError(element.line, `${attr}="${text}" is not True or False`);
		}
		const base = effectiveVariousPropertyBits(record, kind) ?? 0x1B;
		const next = (/^true$/i.test(text) ? (base | bit) : (base & ~bit)) >>> 0;
		if (next !== base || (!recordHas(record, 'VariousPropertyBits') && explicit !== undefined)) {
			setRecordValue(record, 'VariousPropertyBits', next);
			outcome.applied.push(`${attr} of ${name}`);
		}
	}
	applyFontAttrs(record, element, outcome, name, total);
	for (const [attr, field] of [['PasswordChar', 'PasswordChar'], ['Accelerator', 'Accelerator']] as const) {
		const explicit = element.attrs.get(attr);
		const canCarry = record.spec.data.some((f) => f.name === field);
		const text = explicit ?? (total && canCarry ? '' : undefined);
		if (text === undefined) { continue; }
		if (!canCarry) {
			throw new FormMarkupError(element.line, `a ${kind} has no ${attr}`);
		}
		const value = text.length ? text.charCodeAt(0) : 0;
		const current = recordHas(record, field) ? record.values.get(field) : 0;
		if (current !== value || (!recordHas(record, field) && explicit !== undefined && text.length > 0)) {
			setRecordValue(record, field, value);
			outcome.applied.push(`${attr} of ${name}`);
		}
	}
}

/**
 * Font.* on the FORM lands on its StdFont blob: the fBold flag stays zero
 * per the spec (weight 700 carries bold), and a fresh enablement seeds the
 * 0xFFFF Font marker.
 */
function applyFormFontAttrs(
	pkg: FormPackage,
	element: MarkupElement,
	outcome: ApplyOutcome,
): void {
	const wantsFont = [...element.attrs.keys()].some((k) => k.startsWith('Font.'));
	if (!wantsFont) { return; }
	const current = pkg.form.fontRaw ? parseStdFont(pkg.form.fontRaw) : undefined;
	const face = element.attrs.get('Font.Name') ?? current?.face ?? 'Tahoma';
	let heightTT = current?.heightTenThousandthsPt ?? 82500;
	const sizeAttr = element.attrs.get('Font.Size');
	if (sizeAttr !== undefined) {
		const pt = Number(sizeAttr);
		if (!Number.isFinite(pt) || pt <= 0) {
			throw new FormMarkupError(element.line, `Font.Size="${sizeAttr}" is not a size`);
		}
		heightTT = Math.round(pt * 10000);
	}
	// The document is TOTAL for the form's font: the printer speaks every
	// set style, so an absent style attr means unset.
	const decide = (attr: string | undefined): boolean =>
		attr !== undefined && /^true$/i.test(attr);
	const next = composeStdFont(face, heightTT, {
		bold: decide(element.attrs.get('Font.Bold')),
		italic: decide(element.attrs.get('Font.Italic')),
		underline: decide(element.attrs.get('Font.Underline')),
		strikeout: decide(element.attrs.get('Font.Strikethrough')),
		charset: current?.charset ?? 0,
	});
	if (!pkg.form.fontRaw || !next.equals(pkg.form.fontRaw)) {
		pkg.form.fontRaw = next;
		pkg.form.record.maskLo = (pkg.form.record.maskLo | (1 << 20)) >>> 0;
		pkg.form.record.values.set('Font', 0xffff);
		outcome.applied.push('Font of the form');
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
	total = false,
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
	const underlineAttr = element.attrs.get('Font.Underline');
	const strikeAttr = element.attrs.get('Font.Strikethrough');
	if (total || boldAttr !== undefined || italicAttr !== undefined
		|| underlineAttr !== undefined || strikeAttr !== undefined) {
		const current = recordHas(tp, 'FontEffects') ? (tp.values.get('FontEffects') ?? 0) : 0;
		// The printer speaks a set style and stays quiet on an unset one, so
		// a TOTAL apply reads absence as unset; the pane's single write keeps
		// the styles it does not mention.
		const decide = (attr: string | undefined, bit: number): boolean =>
			attr !== undefined ? /^true$/i.test(attr) : (total ? false : (current & bit) !== 0);
		const bold = decide(boldAttr, 0x1);
		const italic = decide(italicAttr, 0x2);
		const underline = decide(underlineAttr, 0x4);
		const strike = decide(strikeAttr, 0x8);
		const next = (current & ~0xf) | (bold ? 0x1 : 0) | (italic ? 0x2 : 0)
			| (underline ? 0x4 : 0) | (strike ? 0x8 : 0);
		const anyExplicit = boldAttr !== undefined || italicAttr !== undefined
			|| underlineAttr !== undefined || strikeAttr !== undefined;
		if (next !== current || (!recordHas(tp, 'FontEffects') && anyExplicit)) {
			setRecordValue(tp, 'FontEffects', next >>> 0);
			setRecordValue(tp, 'FontWeight', bold ? 700 : 400);
			outcome.applied.push(`Font style of ${name}`);
		}
	}
}

/** The TabStrip's per-tab arrays and their DataBlock size fields, by name. */
const TAB_PARALLEL_ARRAYS: ReadonlyArray<readonly [string, string]> = [
	['Items', 'ItemsSize'],
	['TipStrings', 'TipStringsSize'],
	['TabNames', 'NamesSize'],
	['Tags', 'TagsSize'],
	['Accelerators', 'AcceleratorsSize'],
];

/**
 * Inserts one tab's entries at `index` across every stored per-tab array of a
 * TabStrip record, plus the flags tail and the counters - the spec requires
 * every stored array to carry an entry for every tab, so they move together.
 */
function insertTabEntry(record: ParsedRecord, index: number, caption: string, tabName: string): void {
	for (const [arrayName, sizeField] of TAB_PARALLEL_ARRAYS) {
		const raw = record.arrays.get(arrayName);
		if (raw === undefined) { continue; }
		const entries = decodeArrayStrings(raw);
		const value = arrayName === 'Items' ? caption : arrayName === 'TabNames' ? tabName : '';
		entries.splice(index, 0, value);
		const encoded = encodeArrayStrings(entries);
		record.arrays.set(arrayName, encoded);
		setRecordValue(record, sizeField, encoded.length);
	}
	adjustTabFlags(record, index, 'insert');
}

function removeTabEntry(record: ParsedRecord, index: number): void {
	for (const [arrayName, sizeField] of TAB_PARALLEL_ARRAYS) {
		const raw = record.arrays.get(arrayName);
		if (raw === undefined) { continue; }
		const entries = decodeArrayStrings(raw);
		entries.splice(index, 1);
		const encoded = encodeArrayStrings(entries);
		record.arrays.set(arrayName, encoded);
		setRecordValue(record, sizeField, encoded.length);
	}
	adjustTabFlags(record, index, 'remove');
}

/** Keeps TabData, the TabStripTabFlags tail, and TabsAllocated in step. */
function adjustTabFlags(record: ParsedRecord, index: number, op: 'insert' | 'remove'): void {
	const count = record.values.get('TabData');
	if (count !== undefined && recordHas(record, 'TabData')) {
		const flags = record.tailRaw ?? Buffer.alloc(0);
		const entries: Buffer[] = [];
		for (let i = 0; i + 4 <= flags.length; i += 4) { entries.push(flags.subarray(i, i + 4)); }
		if (op === 'insert') {
			// Visible and enabled, the flags every fixture tab carries.
			const fresh = Buffer.alloc(4);
			fresh.writeUInt32LE(0x00000003);
			entries.splice(index, 0, fresh);
		} else {
			entries.splice(index, 1);
		}
		record.tailRaw = Buffer.concat(entries);
		setRecordValue(record, 'TabData', count + (op === 'insert' ? 1 : -1));
	}
	const itemsRaw = record.arrays.get('Items');
	const newCount = itemsRaw ? decodeArrayStrings(itemsRaw).length : 0;
	const allocated = record.values.get('TabsAllocated');
	if (allocated !== undefined && recordHas(record, 'TabsAllocated') && newCount > allocated) {
		setRecordValue(record, 'TabsAllocated', newCount);
	}
}

/**
 * <Tab> edits on a standalone TabStrip. Tabs carry no name, so the diff is
 * POSITIONAL: same count recaptions in place, a longer document appends new
 * tabs at the end, a shorter one truncates from the end.
 */
function applyTabCaptions(
	record: ParsedRecord,
	element: MarkupElement,
	outcome: ApplyOutcome,
	name: string,
): void {
	const tabs = element.children.filter((c) => c.tag.toLowerCase() === 'tab');
	if (tabs.length === 0 && element.children.length === 0) { return; }
	const items = record.arrays.get('Items');
	const captions = items ? decodeArrayStrings(items) : [];

	while (captions.length > tabs.length) {
		removeTabEntry(record, captions.length - 1);
		captions.pop();
		outcome.applied.push(`removed a tab of ${name}`);
	}
	for (let i = captions.length; i < tabs.length; i++) {
		const caption = tabs[i].attrs.get('Caption') ?? `Tab${i + 1}`;
		insertTabEntry(record, i, caption, `Tab${i + 1}`);
		captions.push(caption);
		outcome.applied.push(`added a tab of ${name}`);
	}

	let recaptioned = false;
	tabs.forEach((tab, index) => {
		const caption = tab.attrs.get('Caption');
		if (caption !== undefined && caption !== captions[index]) {
			captions[index] = caption;
			recaptioned = true;
		}
	});
	if (recaptioned) {
		const encoded = encodeArrayStrings(captions);
		record.arrays.set('Items', encoded);
		setRecordValue(record, 'ItemsSize', encoded.length);
		outcome.applied.push(`tab captions of ${name}`);
	}
}

/**
 * Allocates a control ID unique across the WHOLE form tree. The fixture's
 * IDs prove the scope: the root form runs 1..20 with gaps at exactly the IDs
 * its nested containers' controls hold (7,8 in the Frame, 11,12,13 in the
 * MultiPage, 14 on a Page) - one counter serves everything, and a page
 * control re-using a root-level ID broke the page's binding in Excel.
 * Every NextAvailableID on the path records the assignment, the way the VBE
 * leaves each container's counter at the last ID it handed out.
 */
function allocateControlId(root: FormPackage, local: FormPackage): number {
	let id = 1;
	const consider = (candidate: number | undefined): void => {
		if (candidate !== undefined && candidate >= id) { id = candidate + 1; }
	};
	const walk = (pkg: FormPackage): void => {
		consider(pkg.form.record.values.get('NextAvailableID'));
		for (const site of pkg.form.sites) { consider(site.values.get('ID')); }
		for (const child of pkg.containers.values()) { walk(child); }
	};
	walk(root);
	setRecordValue(root.form.record, 'NextAvailableID', id);
	if (local !== root) {
		setRecordValue(local.form.record, 'NextAvailableID', id);
	}
	return id;
}

// ------------------------------------------------------------- additions

/** The canvas's entry into the same authoring path the markup diff uses. */
export function addControlForDesigner(
	pkg: FormPackage,
	element: MarkupElement,
	root: FormPackage,
): void {
	addControl(pkg, element, { applied: [] }, root);
}

function addControl(
	pkg: FormPackage,
	element: MarkupElement,
	outcome: ApplyOutcome,
	root: FormPackage,
): void {
	const kind = element.tag;
	const name = element.attrs.get('Name')!;
	if (!LEGAL_CONTROL_NAME.test(name)) {
		throw new FormMarkupError(
			element.line,
			`${name}: not a legal control name - VBA wants a letter first, then letters, digits, or underscores`,
		);
	}
	if (kind === 'ActiveX') {
		throw new FormMarkupError(
			element.line,
			`${name}: creating an ActiveX control needs a class table entry this engine does not author yet`,
		);
	}
	if (kind === 'MultiPage') {
		throw new FormMarkupError(
			element.line,
			`${name}: adding a whole MultiPage is not supported yet; pages of an existing one are`,
		);
	}
	if (kind === 'Page' || kind === 'Tab') {
		throw new FormMarkupError(
			element.line,
			`${name}: a ${kind} lives inside a ${kind === 'Page' ? 'MultiPage' : 'TabStrip'}, not on the form`,
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
	// ShapeCookie: an Excel-authored container that holds controls carries
	// one, and a page without one bound its content only once the rest of the
	// authorship matched Excel - but BUMPING an existing cookie broke a form
	// that loaded fine before the bump (measured: a top-level add with 17->18
	// failed, the same add leaving 17 alone loaded). The cookie is the design
	// surface's own bookkeeping: set one only where none exists, never touch
	// one that does.
	if (formRecord.values.get('ShapeCookie') === undefined) {
		setRecordValue(formRecord, 'ShapeCookie', 1);
	}
	const id = allocateControlId(root, pkg);

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
		const inner = newEmptyContainerPackage(pkg, element, 'Frame');
		pkg.containers.set(id, inner);
		pkg.entries.push({ kind: 'container', site });
		// The document's site-level say - its own TabIndex, a Tag, the site
		// flags - lands on the fresh site too; skipping this dropped them on
		// every creation (the chain fuzz's find).
		applySiteAttrs(site, element, outcome);
		applyToPackage(inner, element, outcome, false, root);
		outcome.applied.push(`added Frame ${name}`);
		return;
	}

	const record = newRecordForKind(kind);
	record.sizes.set('Size', {
		width: pointsToHimetric(Number(element.attrs.get('Width') ?? '72')),
		height: pointsToHimetric(Number(element.attrs.get('Height') ?? '18')),
	});
	pkg.entries.push({ kind: 'record', site, record });
	applySiteAttrs(site, element, outcome);
	applyRecordAttrs(record, element, kind, outcome, name);
	outcome.applied.push(`added ${kind} ${name}`);
}

/** Renames a site where it stands: the string, its length field, the mask. */
export function renameSiteInPlace(site: SiteModel, newName: string): void {
	site.strings.set('Name', {
		text: newName,
		compressed: [...newName].every((c) => c.charCodeAt(0) <= 0xff),
		raw: Buffer.alloc(0),
		edited: true,
	});
	site.values.set('NameData', 0);
	site.mask = (site.mask | (1 << 0)) >>> 0;
}

export function nextTabIndex(pkg: FormPackage): number {
	let max = -1;
	for (const site of pkg.form.sites) {
		const t = site.values.get('TabIndex');
		if (t !== undefined && t > max) { max = t; }
	}
	return max + 1;
}

/**
 * A fresh control record, authored field-for-field as live Excel authors one
 * of the same kind - measured from the fixture form the VBE built. The
 * differences are not cosmetic: a TextBox without Excel's
 * VariousPropertyBits, or a MorphData without its reserved mask bit, loads
 * at top level but silently breaks the binding of a MultiPage page that
 * carries it.
 */
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
	if (spec.mask64) {
		// MorphData's mask bit 31 is reserved-MUST-be-1 ([MS-OFORMS] 2.2.5.2).
		record.maskLo = (record.maskLo | (1 << 31)) >>> 0;
		const style = ({ TextBox: 1, ListBox: 2, ComboBox: 3, CheckBox: 4, OptionButton: 5, ToggleButton: 6 } as Record<string, number>)[kind];
		if (style !== undefined && style !== 1) {
			setRecordValue(record, 'DisplayStyle', style);
		}
		switch (kind) {
			case 'TextBox':
				setRecordValue(record, 'VariousPropertyBits', 0x2c80481b);
				break;
			case 'ComboBox':
				setRecordValue(record, 'VariousPropertyBits', 0x2c80481b);
				setRecordValue(record, 'MatchEntry', 1);
				setRecordValue(record, 'ShowDropButtonWhen', 2);
				break;
			case 'ListBox':
				setRecordValue(record, 'ScrollBars', 3);
				setRecordValue(record, 'MatchEntry', 0);
				break;
			case 'CheckBox':
			case 'OptionButton':
			case 'ToggleButton':
				setRecordValue(record, 'BackColor', 0x8000000f);
				setRecordValue(record, 'ForeColor', 0x80000012);
				setRecordString(record, 'Value', '0');
				break;
		}
	} else if (kind === 'SpinButton' || kind === 'ScrollBar') {
		setRecordValue(record, 'Orientation', -1);
	}
	if (spec.textProps) {
		// Every record Excel writes carries Tahoma at 8.25pt (165 twips) with
		// charset 0 and pitch-and-family 2; button-like kinds centre their
		// text with ParagraphAlign 3.
		const textProps: ParsedRecord = {
			spec: TEXT_PROPS_REF,
			maskLo: 0, maskHi: 0,
			values: new Map(), strings: new Map(), sizes: new Map(), arrays: new Map(),
			pads: new Map(), streamData: new Map(),
		};
		setRecordString(textProps, 'FontName', 'Tahoma');
		setRecordValue(textProps, 'FontHeight', 165);
		setRecordValue(textProps, 'FontCharSet', 0);
		setRecordValue(textProps, 'FontPitchAndFamily', 2);
		if (kind === 'CommandButton' || kind === 'ToggleButton') {
			setRecordValue(textProps, 'ParagraphAlign', 3);
		}
		record.textProps = textProps;
	}
	return record;
}

function requireSpec(cacheIndex: number) {
	const spec = RECORD_SPECS_BY_CACHE_INDEX.get(cacheIndex);
	if (!spec) { throw new RangeError(`no record spec for cache index ${cacheIndex}`); }
	return spec;
}

/**
 * A fresh, empty container package: the minimal FormControl an Excel-authored
 * Frame or Page carries. BooleanProperties 0x8004 (enabled, class table not
 * saved), DrawBuffer 32000, LogicalSize zero - measured from the fixture's
 * own containers, where a page authored without them was silently not bound.
 */
function newEmptyContainerPackage(
	parent: FormPackage,
	element: MarkupElement,
	kind: 'Frame' | 'Page',
): FormPackage {
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
	setBit(6); record.values.set('BooleanProperties', 0x00008004);
	setBit(10); record.sizes.set('DisplayedSize', { width, height });
	setBit(11); record.sizes.set('LogicalSize', { width: 0, height: 0 });
	setBit(27); record.values.set('DrawBuffer', 32000);
	if (kind === 'Frame') {
		const caption = element.attrs.get('Caption');
		if (caption !== undefined) {
			setBit(19);
			record.strings.set('Caption', {
				text: caption,
				compressed: [...caption].every((c) => c.charCodeAt(0) <= 0xff),
				raw: Buffer.alloc(0),
				edited: true,
			});
			record.values.set('Caption', 0);
		}
	}
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
		compObjRaw: kind === 'Page' ? PAGE_COMPOBJ : FRAME_COMPOBJ,
	};
}
