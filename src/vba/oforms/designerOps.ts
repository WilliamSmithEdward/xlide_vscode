// Canvas gestures, applied to the designer model: move, resize, add, remove.
//
// These are the same mutations the markup diff performs, addressed by control
// name instead of by document - VBA requires control names to be unique
// form-wide, so a name finds its control wherever it nests. Additions run
// through the same authoring path the markup uses, which is the path live
// Excel verified: per-kind recipes, tree-global IDs, the cookie rules.

import { pointsToHimetric, formatPointsShortest as pts } from './bytes';
import { composeStdFont, parseStdFont, siteName, siteId, siteIsContainer, siteCacheIndex, type SiteModel } from './formStream';
import { controlKindOfSite, walkPackages, type FormPackage } from './formPackage';
import {
	addControlForDesigner,
	applyRecordAttrs,
	applySiteAttrs,
	decodeArrayStrings,
	effectiveVariousPropertyBits,
	encodeArrayStrings,
	FORM_EXTRA_FIELDS,
	formatOleColor,
	FormMarkupError,
	LEGAL_CONTROL_NAME,
	nextTabIndex,
	parseOleColor,
	PRINTED_FIELDS,
	ALIGNMENT_KINDS,
	SITE_BITFLAGS_DEFAULT,
	SITE_FLAGS,
	TEXT_ALIGN_KINDS,
	VPB_FLAGS,
	type ApplyOutcome,
	type MarkupElement,
} from './markup';
import { recordHas, setRecordString, setRecordValue, type ParsedRecord } from './records';

export interface ControlLocation {
	pkg: FormPackage;
	site: SiteModel;
	entry: FormPackage['entries'][number];
}

/** Finds a control by name anywhere in the form tree. */
export function findControl(root: FormPackage, name: string): ControlLocation | undefined {
	let found: ControlLocation | undefined;
	walkPackages(root, (pkg) => {
		if (found) { return; }
		for (const entry of pkg.entries) {
			if (siteName(entry.site).toLowerCase() === name.toLowerCase()) {
				found = { pkg, site: entry.site, entry };
				return;
			}
		}
	});
	return found;
}

/** Finds a container surface (a Frame or Page package) by its name; '' is the root. */
export function findSurface(root: FormPackage, containerName: string): FormPackage | undefined {
	if (!containerName) { return root; }
	let found: FormPackage | undefined;
	walkPackages(root, (pkg) => {
		if (found) { return; }
		for (const site of pkg.form.sites) {
			if (!siteIsContainer(site)) { continue; }
			if (siteName(site).toLowerCase() === containerName.toLowerCase()) {
				found = pkg.containers.get(siteId(site));
				return;
			}
		}
	});
	return found;
}

export interface GeometryPt {
	left?: number;
	top?: number;
	width?: number;
	height?: number;
}

/**
 * Moves and/or resizes one control, in points. Position lands on the site;
 * size lands on the record, or on the container's own form for a Frame,
 * MultiPage, or Page.
 */
export function setControlGeometry(root: FormPackage, name: string, geometry: GeometryPt): string[] {
	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	const applied: string[] = [];
	const { site, entry, pkg } = location;
	if (geometry.left !== undefined || geometry.top !== undefined) {
		const current = site.position ?? { left: 0, top: 0 };
		const next = {
			left: geometry.left !== undefined ? pointsToHimetric(geometry.left) : current.left,
			top: geometry.top !== undefined ? pointsToHimetric(geometry.top) : current.top,
		};
		if (next.left !== current.left || next.top !== current.top) {
			site.position = next;
			site.mask = (site.mask | (1 << 8)) >>> 0;
			applied.push(`position of ${siteName(site)}`);
		}
	}
	if (geometry.width !== undefined || geometry.height !== undefined) {
		const target = entry.kind === 'record'
			? { record: entry.record, key: 'Size' as const }
			: undefined;
		if (target) {
			const size = target.record.sizes.get(target.key) ?? { width: 0, height: 0 };
			const next = {
				width: geometry.width !== undefined ? pointsToHimetric(geometry.width) : size.width,
				height: geometry.height !== undefined ? pointsToHimetric(geometry.height) : size.height,
			};
			if (next.width !== size.width || next.height !== size.height) {
				target.record.sizes.set(target.key, next);
				applied.push(`size of ${siteName(site)}`);
			}
		} else {
			const inner = pkg.containers.get(siteId(site));
			const record = inner?.form.record;
			const size = record?.sizes.get('DisplayedSize');
			if (record && size) {
				const next = {
					width: geometry.width !== undefined ? pointsToHimetric(geometry.width) : size.width,
					height: geometry.height !== undefined ? pointsToHimetric(geometry.height) : size.height,
				};
				if (next.width !== size.width || next.height !== size.height) {
					record.sizes.set('DisplayedSize', next);
					const logical = record.sizes.get('LogicalSize');
					if (logical && (logical.width !== 0 || logical.height !== 0)) {
						record.sizes.set('LogicalSize', next);
					}
					applied.push(`size of ${siteName(site)}`);
				}
			}
		}
	}
	return applied;
}

/** A fresh name for a kind, unique across the whole tree: Label1, Label2... */
export function nextControlName(root: FormPackage, kind: string): string {
	const taken = new Set<string>();
	walkPackages(root, (pkg) => {
		for (const site of pkg.form.sites) { taken.add(siteName(site).toLowerCase()); }
	});
	for (let i = 1; ; i++) {
		const candidate = `${kind}${i}`;
		if (!taken.has(candidate.toLowerCase())) { return candidate; }
	}
}

/** The default size a new control of `kind` takes on the canvas, in points. */
export function defaultSizePt(kind: string): { width: number; height: number } {
	switch (kind) {
		case 'Label': return { width: 72, height: 12 };
		case 'CommandButton': case 'ToggleButton': return { width: 72, height: 24 };
		case 'CheckBox': case 'OptionButton': return { width: 90, height: 15 };
		case 'ComboBox': case 'TextBox': return { width: 96, height: 18 };
		case 'ListBox': return { width: 96, height: 60 };
		case 'Frame': return { width: 120, height: 90 };
		case 'Image': return { width: 72, height: 54 };
		case 'SpinButton': return { width: 13, height: 36 };
		case 'ScrollBar': return { width: 13, height: 90 };
		case 'TabStrip': return { width: 150, height: 90 };
		default: return { width: 72, height: 24 };
	}
}

/**
 * Adds a control of `kind` at (left, top) points on the named surface,
 * through the markup's own authoring path. Returns the new control's name.
 */
export function addControlAt(
	root: FormPackage,
	containerName: string,
	kind: string,
	leftPt: number,
	topPt: number,
): string {
	const surface = findSurface(root, containerName);
	if (!surface) {
		throw new FormMarkupError(0, `no container named ${containerName}`);
	}
	const name = nextControlName(root, kind);
	const size = defaultSizePt(kind);
	const attrs = new Map<string, string>([
		['Name', name],
		['Left', String(leftPt)],
		['Top', String(topPt)],
		['Width', String(size.width)],
		['Height', String(size.height)],
	]);
	if (['Label', 'CommandButton', 'ToggleButton', 'CheckBox', 'OptionButton', 'Frame'].includes(kind)) {
		attrs.set('Caption', name);
	}
	const element: MarkupElement = { tag: kind, attrs, children: [], line: 0 };
	addControlForDesigner(surface, element, root);
	return name;
}

/**
 * Removes one control by name. Pages are refused here - removing a page
 * moves four structures at once and belongs to the markup diff, where the
 * whole document states the intent.
 */
export function removeControl(root: FormPackage, name: string): void {
	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	if (siteCacheIndex(location.site) === 7) {
		throw new FormMarkupError(0, `${name} is a Page; remove pages through the form markup`);
	}
	const { pkg, site } = location;
	pkg.form.sites = pkg.form.sites.filter((s) => s !== site);
	pkg.entries = pkg.entries.filter((e) => e.site !== site);
	pkg.containers.delete(siteId(site));
	pkg.form.sitesStructurallyChanged = true;
}

/**
 * Moves one control onto a different surface - the VBE's drag between
 * containers. The site, its record bytes, and (for a container) its whole
 * nested storage move intact, so every property survives; only position,
 * TabIndex, and group membership answer to the new home. Pages stay with
 * their MultiPage, and a container never lands inside itself.
 */
export function reparentControl(
	root: FormPackage,
	name: string,
	containerName: string,
	leftPt: number,
	topPt: number,
): void {
	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	const { pkg: source, site, entry } = location;
	if (siteCacheIndex(site) === 7) {
		throw new FormMarkupError(0, `${name} is a Page; pages stay inside their MultiPage`);
	}
	if ((siteCacheIndex(site) & 0x8000) !== 0) {
		// A class-table index points into ITS container's SiteClassInfo list;
		// the entry does not exist in the target, so the site cannot move.
		throw new FormMarkupError(0, `${name}: an ActiveX control's class entry lives in its container and cannot move`);
	}
	const target = findSurface(root, containerName);
	if (!target) {
		throw new FormMarkupError(0, `no container named ${containerName}`);
	}
	if (target === source) {
		// The same surface: an ordinary move.
		setControlGeometry(root, name, { left: leftPt, top: topPt });
		return;
	}
	const subtree = siteIsContainer(site) ? source.containers.get(siteId(site)) : undefined;
	if (subtree) {
		let cyclic = false;
		walkPackages(subtree, (p) => { if (p === target) { cyclic = true; } });
		if (cyclic) {
			throw new FormMarkupError(0, `${name} cannot move into itself`);
		}
	}

	source.form.sites = source.form.sites.filter((s) => s !== site);
	source.entries = source.entries.filter((e) => e.site !== site);
	source.containers.delete(siteId(site));
	source.form.sitesStructurallyChanged = true;

	// Tab order joins the end of the target's, as the VBE assigns on a drop,
	// and a group stays behind with its container. The ShapeCookie rule from
	// the page-binding work applies to any surface gaining its first control.
	if (site.values.get('TabIndex') !== undefined) {
		site.values.set('TabIndex', nextTabIndex(target));
	}
	if ((site.values.get('GroupID') ?? 0) !== 0) {
		site.values.set('GroupID', 0);
	}
	if (target.form.record.values.get('ShapeCookie') === undefined) {
		setRecordValue(target.form.record, 'ShapeCookie', 1);
	}
	target.form.sites.push(site);
	target.entries.push(entry);
	if (subtree) { target.containers.set(siteId(site), subtree); }
	target.form.sitesStructurallyChanged = true;

	// Position is container-relative; the drop names the point in the new one.
	site.position = { left: pointsToHimetric(leftPt), top: pointsToHimetric(topPt) };
	site.mask = (site.mask | (1 << 8)) >>> 0;
}

// ------------------------------------------------------------ property pane
//
// The pane's rows are the markup dialect's own vocabulary: everything a row
// shows is exactly what the document could spell, drawn from the same record
// and site fields the printer reads - one answer behind both faces. A blank
// value is a property at its default, spelled nowhere.

export interface PropertyRow {
	prop: string;
	value: string;
}

const COLOR_PROPS = ['BackColor', 'ForeColor', 'BorderColor'] as const;
const GEOMETRY_PROPS = ['Left', 'Top', 'Width', 'Height'] as const;

function fontRows(record: ParsedRecord): PropertyRow[] {
	const tp = record.textProps;
	if (!tp) { return []; }
	const face = tp.strings.get('FontName');
	const height = tp.values.get('FontHeight');
	const effects = recordHas(tp, 'FontEffects') ? (tp.values.get('FontEffects') ?? 0) : 0;
	const weight = recordHas(tp, 'FontWeight') ? (tp.values.get('FontWeight') ?? 0) : 0;
	const bold = (effects & 0x1) !== 0 || weight >= 600;
	return [
		{ prop: 'Font.Name', value: face && recordHas(tp, 'FontName') ? face.text : '' },
		{ prop: 'Font.Size', value: height !== undefined && recordHas(tp, 'FontHeight') ? String(Math.round((height / 20) * 100) / 100) : '' },
		{ prop: 'Font.Bold', value: bold ? 'True' : '' },
		{ prop: 'Font.Italic', value: (effects & 0x2) !== 0 ? 'True' : '' },
		{ prop: 'Font.Underline', value: (effects & 0x4) !== 0 ? 'True' : '' },
		{ prop: 'Font.Strikethrough', value: (effects & 0x8) !== 0 ? 'True' : '' },
	];
}

function siteRows(site: SiteModel, kind: string): PropertyRow[] {
	const tab = site.values.get('TabIndex');
	const rows: PropertyRow[] = [
		{ prop: 'TabIndex', value: tab !== undefined && tab >= 0 ? String(tab) : '' },
		{ prop: 'ControlTipText', value: site.strings.get('ControlTipText')?.text ?? '' },
		{ prop: 'Tag', value: site.strings.get('Tag')?.text ?? '' },
	];
	if (CONTROL_SOURCE_KINDS.has(kind)) {
		rows.push({ prop: 'ControlSource', value: site.strings.get('ControlSource')?.text ?? '' });
	}
	if (ROW_SOURCE_KINDS.has(kind)) {
		rows.push({ prop: 'RowSource', value: site.strings.get('RowSource')?.text ?? '' });
	}
	if (kind !== 'Page') {
		const help = site.values.get('HelpContextID');
		rows.push({ prop: 'HelpContextID', value: help !== undefined && help !== 0 ? String(help) : '' });
	}
	const bits = (site.values.get('BitFlags') ?? SITE_BITFLAGS_DEFAULT) >>> 0;
	for (const [prop, bit] of SITE_FLAGS) {
		if (kind === 'Page' || (prop === 'TabStop' && (kind === 'Label' || kind === 'Image'))) { continue; }
		if ((prop === 'Default' || prop === 'Cancel') && kind !== 'CommandButton') { continue; }
		rows.push({ prop, value: (bits & bit) !== 0 ? 'True' : 'False' });
	}
	return rows;
}

function recordControlRows(kind: string, site: SiteModel, record: ParsedRecord): PropertyRow[] {
	const rows: PropertyRow[] = [{ prop: 'Name', value: siteName(site) }];
	const pos = site.position;
	rows.push({ prop: 'Left', value: pos ? pts(pos.left) : '' }, { prop: 'Top', value: pos ? pts(pos.top) : '' });
	const size = record.sizes.get('Size');
	rows.push({ prop: 'Width', value: size ? pts(size.width) : '' }, { prop: 'Height', value: size ? pts(size.height) : '' });
	for (const field of ['Caption', 'Value', 'GroupName']) {
		if (!record.spec.extra.some((f) => f.name === field && f.kind === 'str')) { continue; }
		rows.push({ prop: field, value: record.strings.get(field)?.text ?? '' });
	}
	for (const field of COLOR_PROPS) {
		if (!record.spec.data.some((f) => f.name === field)) { continue; }
		const v = record.values.get(field);
		rows.push({ prop: field, value: v !== undefined && recordHas(record, field) ? formatOleColor(v) : '' });
	}
	for (const field of PRINTED_FIELDS[kind] ?? []) {
		if ((COLOR_PROPS as readonly string[]).includes(field)) { continue; }
		if (!record.spec.data.some((f) => f.name === field)) { continue; }
		const v = record.values.get(field);
		rows.push({ prop: field, value: v !== undefined && recordHas(record, field) ? String(v) : '' });
	}
	for (const field of ['PasswordChar', 'Accelerator']) {
		if (!record.spec.data.some((f) => f.name === field)) { continue; }
		const v = record.values.get(field);
		rows.push({ prop: field, value: v !== undefined && recordHas(record, field) && v !== 0 ? String.fromCharCode(v) : '' });
	}
	if (kind === 'ComboBox') {
		rows.push({ prop: 'Style', value: record.values.get('DisplayStyle') === 7 ? '2' : '0' });
	}
	if (ALIGNMENT_KINDS.has(kind)) {
		const vpbNow = effectiveVariousPropertyBits(record, kind) ?? 0;
		rows.push({ prop: 'Alignment', value: (vpbNow & 0x2000) !== 0 ? '0' : '1' });
	}
	if (TEXT_ALIGN_KINDS.has(kind) && record.textProps) {
		const pa = recordHas(record.textProps, 'ParagraphAlign')
			? (record.textProps.values.get('ParagraphAlign') ?? 1) : 1;
		rows.push({ prop: 'TextAlign', value: pa === 2 ? 'Right' : pa === 3 ? 'Center' : 'Left' });
	}
	const effectiveVpb = effectiveVariousPropertyBits(record, kind);
	if (effectiveVpb !== undefined) {
		for (const [prop, bit, kinds] of VPB_FLAGS) {
			if (!kinds.includes(kind)) { continue; }
			rows.push({ prop, value: (effectiveVpb & bit) !== 0 ? 'True' : 'False' });
		}
	}
	rows.push(...fontRows(record));
	rows.push(...siteRows(site, kind));
	return rows;
}

function innerRecordRows(record: ParsedRecord): PropertyRow[] {
	const rows: PropertyRow[] = [];
	for (const field of [...COLOR_PROPS, 'SpecialEffect']) {
		if (!record.spec.data.some((f) => f.name === field)) { continue; }
		const v = record.values.get(field);
		const shown = v !== undefined && recordHas(record, field)
			? (field === 'SpecialEffect' ? String(v) : formatOleColor(v))
			: '';
		rows.push({ prop: field, value: shown });
	}
	return rows;
}

function containerRows(
	kind: string,
	site: SiteModel,
	inner: FormPackage | undefined,
	pageCaption: string | undefined,
): PropertyRow[] {
	const rows: PropertyRow[] = [{ prop: 'Name', value: siteName(site) }];
	const pos = site.position;
	if (pos) { rows.push({ prop: 'Left', value: pts(pos.left) }, { prop: 'Top', value: pts(pos.top) }); }
	const size = inner?.form.record.sizes.get('DisplayedSize');
	rows.push({ prop: 'Width', value: size ? pts(size.width) : '' }, { prop: 'Height', value: size ? pts(size.height) : '' });
	if (kind === 'Frame') {
		rows.push({ prop: 'Caption', value: inner?.form.record.strings.get('Caption')?.text ?? '' });
	}
	if (kind === 'Page') {
		rows.push({ prop: 'Caption', value: pageCaption ?? '' });
	}
	if (inner) { rows.push(...innerRecordRows(inner.form.record)); }
	rows.push(...siteRows(site, kind));
	return rows;
}

/** VBFrame-held form properties the service reads out of the frame text. */
export interface VbFrameProps {
	showModal?: string;
	startUpPosition?: string;
	whatsThisButton?: string;
}

// One list, one concept: the dialect's own form extras (markup.ts).
const FORM_NUMERIC_PROPS = FORM_EXTRA_FIELDS;

const CONTROL_SOURCE_KINDS = new Set(['TextBox', 'ComboBox', 'ListBox', 'CheckBox', 'OptionButton', 'ToggleButton', 'ScrollBar', 'SpinButton']);
const ROW_SOURCE_KINDS = new Set(['ComboBox', 'ListBox']);

function formRows(
	root: FormPackage,
	formName: string,
	captionFallback?: string,
	vbFrame?: VbFrameProps,
): PropertyRow[] {
	const record = root.form.record;
	const rows: PropertyRow[] = [{ prop: 'Name', value: formName }];
	rows.push({ prop: 'Caption', value: record.strings.get('Caption')?.text ?? captionFallback ?? '' });
	const size = record.sizes.get('DisplayedSize');
	rows.push({ prop: 'Width', value: size ? pts(size.width) : '' }, { prop: 'Height', value: size ? pts(size.height) : '' });
	rows.push(...innerRecordRows(record));
	for (const field of FORM_NUMERIC_PROPS) {
		if (!record.spec.data.some((f) => f.name === field)) { continue; }
		const v = record.values.get(field);
		rows.push({ prop: field, value: v !== undefined && recordHas(record, field) ? String(v) : '' });
	}
	// The form's font is a StdFont blob, not TextProps. Its fBold flag MUST
	// stay zero per the spec, so bold reads from the weight.
	const font = root.form.fontRaw ? parseStdFont(root.form.fontRaw) : undefined;
	rows.push({ prop: 'Font.Name', value: font?.face ?? '' });
	rows.push({ prop: 'Font.Size', value: font ? String(Math.round(font.heightTenThousandthsPt / 100) / 100) : '' });
	rows.push({ prop: 'Font.Bold', value: font && ((font.flags & 0x1) !== 0 || font.weight >= 600) ? 'True' : 'False' });
	rows.push({ prop: 'Font.Italic', value: font && (font.flags & 0x2) !== 0 ? 'True' : 'False' });
	rows.push({ prop: 'Font.Underline', value: font && (font.flags & 0x4) !== 0 ? 'True' : 'False' });
	rows.push({ prop: 'Font.Strikethrough', value: font && (font.flags & 0x8) !== 0 ? 'True' : 'False' });
	if (vbFrame) {
		rows.push({ prop: 'StartUpPosition', value: vbFrame.startUpPosition ?? '1' });
		rows.push({ prop: 'ShowModal', value: vbFrame.showModal ?? 'True' });
		rows.push({ prop: 'WhatsThisButton', value: vbFrame.whatsThisButton ?? 'False' });
	}
	return rows;
}

/** The tab captions of a package that owns pages (a MultiPage's inside). */
function tabCaptionsOf(pkg: FormPackage): string[] {
	const entry = pkg.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
	return entry && entry.kind === 'record'
		? decodeArrayStrings(entry.record.arrays.get('Items') ?? Buffer.alloc(0))
		: [];
}

/**
 * Every target's property rows, keyed by control name - '' is the form. The
 * designer's Properties pane renders straight from this.
 */
export function listFormProperties(
	root: FormPackage,
	formName: string,
	captionFallback?: string,
	vbFrame?: VbFrameProps,
): Record<string, { kind: string; rows: PropertyRow[] }> {
	const out: Record<string, { kind: string; rows: PropertyRow[] }> = {
		'': { kind: 'Form', rows: formRows(root, formName, captionFallback, vbFrame) },
	};
	walkPackages(root, (pkg) => {
		const captions = tabCaptionsOf(pkg);
		let pageIndex = 0;
		for (const entry of pkg.entries) {
			const site = entry.site;
			const name = siteName(site);
			if (!name) { continue; }
			const kind = controlKindOfSite(site, entry.kind === 'record' ? entry.record : undefined);
			if (siteIsContainer(site)) {
				const pageCaption = siteCacheIndex(site) === 7 ? captions[pageIndex++] : undefined;
				out[name] = { kind, rows: containerRows(kind, site, pkg.containers.get(siteId(site)), pageCaption) };
			} else if (entry.kind === 'record') {
				out[name] = { kind, rows: recordControlRows(kind, site, entry.record) };
			} else {
				out[name] = { kind: 'ActiveX', rows: [{ prop: 'Name', value: name }, ...siteRows(site, 'ActiveX')] };
			}
		}
	});
	return out;
}

function setColorValue(record: ParsedRecord, field: string, text: string, applied: string[], owner: string): void {
	const value = parseOleColor(text);
	if (value === undefined) {
		throw new FormMarkupError(0, `${field}="${text}" is not a color this dialect knows`);
	}
	if (record.values.get(field) !== value || !recordHas(record, field)) {
		setRecordValue(record, field, value);
		applied.push(`${field} of ${owner}`);
	}
}

function setNumericValue(record: ParsedRecord, field: string, text: string, applied: string[], owner: string): void {
	const value = Number(text);
	if (!Number.isFinite(value)) {
		throw new FormMarkupError(0, `${field}="${text}" is not a number`);
	}
	if (record.values.get(field) !== value || !recordHas(record, field)) {
		setRecordValue(record, field, value);
		applied.push(`${field} of ${owner}`);
	}
}

/**
 * Writes ONE property of one target - the Properties pane's gesture, the
 * vbide's setProperty. '' targets the form (its Caption and size belong to
 * the service, which owns the VBFrame). Renaming returns the new name so the
 * pane can follow its control.
 */
export function setControlProperty(
	root: FormPackage,
	name: string,
	prop: string,
	value: string,
): { applied: string[]; renamed?: string } {
	if (name === '') {
		const record = root.form.record;
		const applied: string[] = [];
		if ((COLOR_PROPS as readonly string[]).includes(prop) && record.spec.data.some((f) => f.name === prop)) {
			setColorValue(record, prop, value, applied, 'the form');
			return { applied };
		}
		if ((prop === 'SpecialEffect' || (FORM_NUMERIC_PROPS as readonly string[]).includes(prop))
			&& record.spec.data.some((f) => f.name === prop)) {
			setNumericValue(record, prop, value, applied, 'the form');
			return { applied };
		}
		if (prop.startsWith('Font.')) {
			const current = root.form.fontRaw ? parseStdFont(root.form.fontRaw) : undefined;
			const face = prop === 'Font.Name' ? value : current?.face ?? 'Tahoma';
			let heightTT = current?.heightTenThousandthsPt ?? 82500;
			if (prop === 'Font.Size') {
				const pt = Number(value);
				if (!Number.isFinite(pt) || pt <= 0) {
					throw new FormMarkupError(0, `Font.Size="${value}" is not a size`);
				}
				heightTT = Math.round(pt * 10000);
			}
			const wasBold = !!current && ((current.flags & 0x1) !== 0 || current.weight >= 600);
			const wasItalic = !!current && (current.flags & 0x2) !== 0;
			const wasUnderline = !!current && (current.flags & 0x4) !== 0;
			const wasStrikeout = !!current && (current.flags & 0x8) !== 0;
			const bold = prop === 'Font.Bold' ? /^true$/i.test(value) : wasBold;
			const italic = prop === 'Font.Italic' ? /^true$/i.test(value) : wasItalic;
			const underline = prop === 'Font.Underline' ? /^true$/i.test(value) : wasUnderline;
			const strikeout = prop === 'Font.Strikethrough' ? /^true$/i.test(value) : wasStrikeout;
			if (prop === 'Font.Bold' || prop === 'Font.Italic'
				|| prop === 'Font.Underline' || prop === 'Font.Strikethrough') {
				if (!/^(true|false)$/i.test(value)) {
					throw new FormMarkupError(0, `${prop}="${value}" is not True or False`);
				}
			}
			root.form.fontRaw = composeStdFont(face, heightTT, {
				bold, italic, underline, strikeout, charset: current?.charset ?? 0,
			});
			root.form.record.maskLo = (root.form.record.maskLo | (1 << 20)) >>> 0;
			// The DataBlock's Font marker MUST be 0xFFFF ([MS-OFORMS]); a
			// fresh enablement has no captured value to replay.
			root.form.record.values.set('Font', 0xffff);
			return { applied: [`${prop} of the form`] };
		}
		throw new FormMarkupError(0, `the form has no ${prop} this pane can write`);
	}

	const location = findControl(root, name);
	if (!location) {
		throw new FormMarkupError(0, `no control named ${name}`);
	}
	const { pkg, site, entry } = location;
	const kind = controlKindOfSite(site, entry.kind === 'record' ? entry.record : undefined);

	if (prop === 'Name') {
		if (!LEGAL_CONTROL_NAME.test(value)) {
			throw new FormMarkupError(0, `${value} is not a legal control name`);
		}
		if (siteName(site) === value) { return { applied: [] }; }
		if (value.toLowerCase() !== name.toLowerCase()) {
			let taken = false;
			walkPackages(root, (p) => {
				for (const s of p.form.sites) {
					if (siteName(s).toLowerCase() === value.toLowerCase()) { taken = true; }
				}
			});
			if (taken) {
				throw new FormMarkupError(0, `a control named ${value} already exists`);
			}
		}
		site.strings.set('Name', {
			text: value,
			compressed: [...value].every((c) => c.charCodeAt(0) <= 0xff),
			raw: Buffer.alloc(0),
			edited: true,
		});
		site.values.set('NameData', 0);
		site.mask = (site.mask | (1 << 0)) >>> 0;
		return { applied: [`Name of ${name}`], renamed: value };
	}

	// The pane's vocabulary is the contract: a prop that has no row is not a
	// property of this kind.
	const rows = siteIsContainer(site)
		? containerRows(kind, site, pkg.containers.get(siteId(site)), '')
		: entry.kind === 'record'
			? recordControlRows(kind, site, entry.record)
			: [{ prop: 'Name', value: name }, ...siteRows(site, 'ActiveX')];
	if (!rows.some((r) => r.prop === prop)) {
		throw new FormMarkupError(0, `a ${kind} has no ${prop}`);
	}

	if ((GEOMETRY_PROPS as readonly string[]).includes(prop)) {
		const n = Number(value);
		if (!Number.isFinite(n)) {
			throw new FormMarkupError(0, `${prop}="${value}" is not a number`);
		}
		const applied = setControlGeometry(root, name, { [prop.toLowerCase()]: n });
		return { applied };
	}

	if (siteCacheIndex(site) === 7 && prop === 'Caption') {
		// A Page's caption lives in its MultiPage's tab items, positionally.
		const tabEntry = pkg.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
		if (!tabEntry || tabEntry.kind !== 'record') {
			throw new FormMarkupError(0, `${name}: its MultiPage carries no tab record`);
		}
		const captions = decodeArrayStrings(tabEntry.record.arrays.get('Items') ?? Buffer.alloc(0));
		const index = pkg.form.sites.filter((s) => siteCacheIndex(s) === 7).indexOf(site);
		if (index < 0 || captions[index] === undefined) {
			throw new FormMarkupError(0, `${name} has no tab caption to edit`);
		}
		if (captions[index] === value) { return { applied: [] }; }
		captions[index] = value;
		const encoded = encodeArrayStrings(captions);
		tabEntry.record.arrays.set('Items', encoded);
		setRecordValue(tabEntry.record, 'ItemsSize', encoded.length);
		return { applied: [`Caption of ${name}`] };
	}

	if (siteIsContainer(site)) {
		const inner = pkg.containers.get(siteId(site));
		const applied: string[] = [];
		if (prop === 'Caption' && kind === 'Frame' && inner) {
			if ((inner.form.record.strings.get('Caption')?.text ?? '') !== value) {
				setRecordString(inner.form.record, 'Caption', value);
				applied.push(`Caption of ${name}`);
			}
			return { applied };
		}
		if ((COLOR_PROPS as readonly string[]).includes(prop) && inner) {
			setColorValue(inner.form.record, prop, value, applied, name);
			return { applied };
		}
		if (prop === 'SpecialEffect' && inner) {
			setNumericValue(inner.form.record, prop, value, applied, name);
			return { applied };
		}
		const outcome: ApplyOutcome = { applied };
		const el: MarkupElement = { tag: kind, attrs: new Map([[prop, value]]), children: [], line: 0 };
		applySiteAttrs(site, el, outcome);
		return { applied: outcome.applied };
	}

	const outcome: ApplyOutcome = { applied: [] };
	const el: MarkupElement = { tag: kind, attrs: new Map([[prop, value]]), children: [], line: 0 };
	applySiteAttrs(site, el, outcome);
	if (entry.kind === 'record') {
		applyRecordAttrs(entry.record, el, kind, outcome, name);
	}
	return { applied: outcome.applied };
}

// ------------------------------------------------- identity reconciliation
//
// The markup diff is keyed by NAME, so a renamed or reparented control used
// to read as remove-plus-add - and the rebuilt site kept only what the
// dialect can spell. Its picture, its mouse icon, an ActiveX payload: gone
// from the saved workbook while the canvas (whose scratch renamed in place)
// still showed them. This pre-pass pairs the document's identities with the
// model's and executes the pairs through the IN-PLACE primitives, so the
// diff that follows sees matches. Every pairing demands an unambiguous
// match; anything uncertain falls back to the old remove-plus-add.

interface ReconcileDocEntry {
	owner: string;
	el: MarkupElement;
	kind: string;
	nameLower: string;
}

interface ReconcileModelEntry {
	owner: string;
	name: string;
	kind: string;
	/** Geometry AS THE PRINTER SPELLS IT - stored himetrics do not round-trip
	 *  to whole points (258pt persists as 9102hm = 258.0094pt), so the only
	 *  fingerprint that matches the document is the printed string itself. */
	geo: { l: string; t: string; w: string; h: string };
	/** The stored caption ('' when the kind has none), for the global key. */
	caption: string;
	nameLower: string;
}

function collectDocEntries(el: MarkupElement, owner: string, out: ReconcileDocEntry[]): void {
	for (const child of el.children) {
		const tag = child.tag.toLowerCase();
		if (tag === 'tab') { continue; }
		const name = child.attrs.get('Name') ?? '';
		if (tag === 'page') {
			collectDocEntries(child, name, out);
			continue;
		}
		out.push({ owner, el: child, kind: child.tag, nameLower: name.toLowerCase() });
		if (tag === 'frame') { collectDocEntries(child, name, out); }
		if (tag === 'multipage') { collectDocEntries(child, name, out); }
	}
}

function collectModelEntries(pkg: FormPackage, owner: string, out: ReconcileModelEntry[]): void {
	for (const entry of pkg.entries) {
		const site = entry.site;
		if (siteCacheIndex(site) === 7) { continue; } // pages walk below
		const record = entry.kind === 'record' ? entry.record : undefined;
		const inner = entry.kind === 'container' ? pkg.containers.get(siteId(site)) : undefined;
		const size = record?.sizes.get('Size') ?? inner?.form.record.sizes.get('DisplayedSize');
		out.push({
			owner,
			name: siteName(site),
			nameLower: siteName(site).toLowerCase(),
			kind: controlKindOfSite(site, record),
			geo: {
				l: pts(site.position?.left ?? 0),
				t: pts(site.position?.top ?? 0),
				w: pts(size?.width ?? 0),
				h: pts(size?.height ?? 0),
			},
			caption: record?.strings.get('Caption')?.text
				?? inner?.form.record.strings.get('Caption')?.text
				?? '',
		});
		if (inner) {
			if (controlKindOfSite(site, undefined) === 'MultiPage') {
				for (const pageSite of inner.form.sites) {
					if (siteCacheIndex(pageSite) !== 7) { continue; }
					const pagePkg = inner.containers.get(siteId(pageSite));
					if (pagePkg) { collectModelEntries(pagePkg, siteName(pageSite), out); }
				}
			} else {
				collectModelEntries(inner, siteName(site), out);
			}
		}
	}
}

const docFpOf = (d: ReconcileDocEntry): string | undefined => {
	const l = d.el.attrs.get('Left'), t = d.el.attrs.get('Top');
	const w = d.el.attrs.get('Width'), h = d.el.attrs.get('Height');
	if (l === undefined || t === undefined || w === undefined || h === undefined) { return undefined; }
	return `${d.kind.toLowerCase()}|${l}|${t}|${w}|${h}`;
};


/**
 * Pairs renamed and reparented controls between the document and the model,
 * executing them in place BEFORE the name-keyed diff runs. One pass in
 * DOCUMENT order - two controls moved into the same container must land in
 * the order the document lists them, and a renamed container is visited
 * before its children, so their moves resolve against its new name. Returns
 * what it did, in the apply outcome's own voice.
 */
export function reconcileMarkupIdentities(root: FormPackage, doc: MarkupElement): string[] {
	const applied: string[] = [];
	const docs: ReconcileDocEntry[] = [];
	collectDocEntries(doc, '', docs);
	const docNames = new Set(docs.map((d) => d.nameLower));
	const models = (): ReconcileModelEntry[] => {
		const out: ReconcileModelEntry[] = [];
		collectModelEntries(root, '', out);
		return out;
	};
	const posFpModel = (m: ReconcileModelEntry): string =>
		`${m.kind.toLowerCase()}|${m.geo.l}|${m.geo.t}|${m.geo.w}|${m.geo.h}`;
	const sizeFpDoc = (d: ReconcileDocEntry): string | undefined => {
		const w = d.el.attrs.get('Width');
		const h = d.el.attrs.get('Height');
		if (w === undefined || h === undefined) { return undefined; }
		return `${d.kind.toLowerCase()}|${w}|${h}|${d.el.attrs.get('Caption') ?? ''}`;
	};
	const sizeFpModel = (m: ReconcileModelEntry): string =>
		`${m.kind.toLowerCase()}|${m.geo.w}|${m.geo.h}|${m.caption}`;

	// ONE collection, maintained incrementally: the first version rebuilt the
	// whole model index per document entry, and the parser fuzz measured the
	// quadratic (a 4000-control paste spent seconds in here).
	const current = models();
	const byName = new Map(current.map((entry) => [entry.nameLower, entry]));
	const removed = current.filter((entry) => !docNames.has(entry.nameLower));
	const ownerPosDoc = (a: ReconcileDocEntry): string | undefined => {
		const fp = docFpOf(a);
		return fp === undefined ? undefined : `${a.owner.toLowerCase()}::${fp}`;
	};
	const ownerPosModel = (entry: ReconcileModelEntry): string =>
		`${entry.owner.toLowerCase()}::${posFpModel(entry)}`;
	const bump = (map: Map<string, number>, key: string | undefined, by: number): void => {
		if (key !== undefined) { map.set(key, (map.get(key) ?? 0) + by); }
	};
	const docPos = new Map<string, number>();
	const docSize = new Map<string, number>();
	for (const a of docs) {
		if (a.nameLower && !byName.has(a.nameLower)) {
			bump(docPos, ownerPosDoc(a), 1);
			bump(docSize, sizeFpDoc(a), 1);
		}
	}
	const modelPos = new Map<string, number>();
	const modelSize = new Map<string, number>();
	for (const entry of removed) {
		bump(modelPos, ownerPosModel(entry), 1);
		bump(modelSize, sizeFpModel(entry), 1);
	}

	for (const d of docs) {
		if (!d.nameLower) { continue; }
		const m = byName.get(d.nameLower);

		if (m) {
			// The same name under a different owner is a REPARENT, at the
			// document's own coordinates.
			if (m.owner.toLowerCase() === d.owner.toLowerCase()) { continue; }
			if (d.kind.toLowerCase() !== m.kind.toLowerCase() && d.kind !== 'ActiveX') { continue; }
			const left = Number(d.el.attrs.get('Left'));
			const top = Number(d.el.attrs.get('Top'));
			if (!Number.isFinite(left) || !Number.isFinite(top)) { continue; }
			try {
				reparentControl(root, m.name, d.owner, left, top);
				m.owner = d.owner;
				applied.push(`moved ${m.name} into ${d.owner || 'the form'}`);
			} catch { /* an unmovable site falls back to remove-plus-add */ }
			continue;
		}
		if (removed.length === 0) { continue; } // plain additions pair with nothing

		// A document-only name: a RENAME when exactly one model-only control
		// matches - same owner by kind + full printed geometry, or anywhere
		// by kind + size + caption (position changes with a move; size and
		// caption travel, and the caption tells a moved "Start" button from
		// a FRESH default-sized add). Anything ambiguous falls back to the
		// old remove-plus-add.
		const newName = d.el.attrs.get('Name') ?? '';
		if (!LEGAL_CONTROL_NAME.test(newName)) { continue; }
		const posKey = ownerPosDoc(d);
		let pairTo: ReconcileModelEntry | undefined;
		if (posKey !== undefined && docPos.get(posKey) === 1 && modelPos.get(posKey) === 1) {
			pairTo = removed.find((entry) => ownerPosModel(entry) === posKey);
		}
		const sizeKey = sizeFpDoc(d);
		if (!pairTo && sizeKey !== undefined && docSize.get(sizeKey) === 1 && modelSize.get(sizeKey) === 1) {
			pairTo = removed.find((entry) => sizeFpModel(entry) === sizeKey);
		}
		if (!pairTo) { continue; }
		const left = Number(d.el.attrs.get('Left'));
		const top = Number(d.el.attrs.get('Top'));
		const moves = pairTo.owner.toLowerCase() !== d.owner.toLowerCase();
		const oldName = pairTo.name;
		const oldNameLower = pairTo.nameLower;
		const oldPosKey = ownerPosModel(pairTo);
		const oldSizeKey = sizeFpModel(pairTo);
		try {
			if (moves) {
				if (!Number.isFinite(left) || !Number.isFinite(top)) { continue; }
				reparentControl(root, oldName, d.owner, left, top);
			}
			setControlProperty(root, oldName, 'Name', newName);
			applied.push(`renamed ${oldName} to ${newName}${moves ? ` in ${d.owner || 'the form'}` : ''}`);
			// The pair is CONSUMED: fix every index the loop still reads.
			removed.splice(removed.indexOf(pairTo), 1);
			bump(modelPos, oldPosKey, -1);
			bump(modelSize, oldSizeKey, -1);
			bump(docPos, posKey, -1);
			bump(docSize, sizeKey, -1);
			byName.delete(oldNameLower);
			pairTo.name = newName;
			pairTo.nameLower = newName.toLowerCase();
			if (moves) { pairTo.owner = d.owner; }
			byName.set(pairTo.nameLower, pairTo);
			// A renamed CONTAINER re-homes its children's owner strings.
			for (const entry of current) {
				if (entry.owner.toLowerCase() === oldNameLower) { entry.owner = newName; }
			}
		} catch { /* an uncertain pairing falls back to remove-plus-add */ }
	}

	return applied;
}

/** Resizes the form's own client area, in points. */
export function setFormSize(root: FormPackage, widthPt: number, heightPt: number): void {
	const record = root.form.record;
	const size = record.sizes.get('DisplayedSize');
	if (!size) { return; }
	const next = { width: pointsToHimetric(widthPt), height: pointsToHimetric(heightPt) };
	record.sizes.set('DisplayedSize', next);
	const logical = record.sizes.get('LogicalSize');
	if (logical && (logical.width !== 0 || logical.height !== 0)) {
		record.sizes.set('LogicalSize', next);
	}
}
