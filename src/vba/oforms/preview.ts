// The form preview: the parsed designer model rendered to self-contained
// HTML for a webview canvas.
//
// This is a RENDERER of the same projection the markup document edits, not a
// rival editing model - the sequencing the vbide design chose. It draws what
// it honestly knows: real bounds, captions, colors, and fonts from the
// binary. Where fidelity runs out - a picture, a third-party control - it
// shows the control's true bounds and identity rather than an approximation
// that lies.

import { himetricToPoints, formatPointsShortest as pts } from './bytes';
import { recordHas, type ParsedRecord } from './records';
import { siteName, siteId, siteIsContainer, siteCacheIndex, type SiteModel } from './formStream';
import { controlKindOfSite, type FormPackage } from './formPackage';
import { decodeArrayStrings, effectiveVariousPropertyBits, formatOleColor } from './markup';
import { parseStdFont } from './formStream';

// The Windows default palette, by GetSysColor value - what MSForms actually
// paints with. CSS system-color KEYWORDS are useless here: modern Chromium
// computes the deprecated ButtonShadow as rgb(240,240,240), identical to
// ButtonFace, so a keyword-colored border on a ButtonFace surface is
// invisible by definition (measured; the page and frame borders vanished).
const WINDOWS_PALETTE: Readonly<Record<string, string>> = {
	ScrollBars: '#c8c8c8', Desktop: '#000000', ActiveTitleBar: '#99b4d1',
	InactiveTitleBar: '#bfcddb', MenuBar: '#f0f0f0', WindowBackground: '#ffffff',
	WindowFrame: '#646464', MenuText: '#000000', WindowText: '#000000',
	TitleBarText: '#000000', ActiveBorder: '#b4b4b4', InactiveBorder: '#f4f7fc',
	ApplicationWorkspace: '#ababab', Highlight: '#0078d7', HighlightText: '#ffffff',
	ButtonFace: '#f0f0f0', ButtonShadow: '#a0a0a0', GrayText: '#6d6d6d',
	ButtonText: '#000000', InactiveTitleBarText: '#000000', ButtonHighlight: '#ffffff',
	ButtonDarkShadow: '#696969', ButtonLight: '#e3e3e3', InfoText: '#000000',
	InfoBackground: '#ffffe1', HotTracking: '#0066cc', GradientActiveTitleBar: '#b9d1ea',
	GradientInactiveTitleBar: '#d7e4f2', MenuHighlight: '#3399ff', MenuBackground: '#f0f0f0',
};

/** OLE_COLOR -> CSS: literals verbatim, system indices via the palette. */
function cssColor(value: number): string {
	const spelled = formatOleColor(value);
	if (spelled.startsWith('#')) { return spelled; }
	return WINDOWS_PALETTE[spelled] ?? WINDOWS_PALETTE.ButtonFace;
}

function esc(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

/** HIMETRIC box, formatted to points at print time. */
interface Box { left: number; top: number; width: number; height: number; }

function siteBox(site: SiteModel, record?: ParsedRecord, containerPkg?: FormPackage): Box {
	const pos = site.position ?? { left: 0, top: 0 };
	const size = record?.sizes.get('Size')
		?? containerPkg?.form.record.sizes.get('DisplayedSize')
		?? { width: 0, height: 0 };
	return {
		left: pos.left,
		top: pos.top,
		width: size.width,
		height: size.height,
	};
}

function fontCss(record: ParsedRecord | undefined): string {
	const tp = record?.textProps;
	if (!tp) { return ''; }
	const parts: string[] = [];
	const name = tp.strings.get('FontName');
	if (name && recordHas(tp, 'FontName')) { parts.push(`font-family:'${esc(name.text)}',Tahoma,sans-serif;`); }
	const height = tp.values.get('FontHeight');
	if (height !== undefined && recordHas(tp, 'FontHeight')) { parts.push(`font-size:${height / 20}pt;`); }
	const effects = tp.values.get('FontEffects') ?? 0;
	const weight = tp.values.get('FontWeight') ?? 0;
	if ((recordHas(tp, 'FontEffects') && (effects & 0x1)) || (recordHas(tp, 'FontWeight') && weight >= 600)) {
		parts.push('font-weight:bold;');
	}
	if (recordHas(tp, 'FontEffects') && (effects & 0x2)) { parts.push('font-style:italic;'); }
	const deco = [
		recordHas(tp, 'FontEffects') && (effects & 0x4) ? 'underline' : '',
		recordHas(tp, 'FontEffects') && (effects & 0x8) ? 'line-through' : '',
	].filter(Boolean).join(' ');
	if (deco) { parts.push(`text-decoration:${deco};`); }
	const pa = tp.values.get('ParagraphAlign');
	if (pa !== undefined && recordHas(tp, 'ParagraphAlign')) {
		if (pa === 2) { parts.push('text-align:right;justify-content:flex-end;'); }
		if (pa === 3) { parts.push('text-align:center;justify-content:center;'); }
	}
	return parts.join('');
}

/**
 * A container's own font, from its inner form's StdFont - falling back to
 * the CONTROL default (Tahoma 8.25), never to the form's font: MSForms
 * containers do not inherit the form's typeface, so a large form font must
 * not swell frame legends and page tabs (measured against the native
 * render).
 */
function containerFontCss(inner: FormPackage | undefined): string {
	const font = inner?.form.fontRaw ? parseStdFont(inner.form.fontRaw) : undefined;
	if (!font) { return "font-family:Tahoma,sans-serif;font-size:8.25pt;"; }
	return `font-family:'${esc(font.face)}',Tahoma,sans-serif;`
		+ `font-size:${font.heightTenThousandthsPt / 10000}pt;`
		+ (((font.flags & 0x1) !== 0 || font.weight >= 600) ? 'font-weight:bold;' : '')
		+ ((font.flags & 0x2) !== 0 ? 'font-style:italic;' : '');
}

/**
 * The border a control's stored SpecialEffect / BorderStyle / BorderColor
 * ask for, overriding the class defaults inline. fmSpecialEffect: 0 flat,
 * 1 raised, 2 sunken, 3 etched, 6 bump - CSS's outset/inset/groove/ridge
 * are their direct ancestors.
 */
function borderCss(record: ParsedRecord | undefined): string {
	if (!record) { return ''; }
	const hasStyle = record.spec.data.some((f) => f.name === 'BorderStyle') && recordHas(record, 'BorderStyle');
	const hasEffect = record.spec.data.some((f) => f.name === 'SpecialEffect') && recordHas(record, 'SpecialEffect');
	if (!hasStyle && !hasEffect) { return ''; }
	const borderColor = record.spec.data.some((f) => f.name === 'BorderColor') && recordHas(record, 'BorderColor')
		? cssColor(record.values.get('BorderColor') ?? 0)
		: '#7a7a7a';
	if (hasStyle && (record.values.get('BorderStyle') ?? 0) === 1) {
		return `border:1px solid ${borderColor};box-shadow:none;`;
	}
	const effect = hasEffect ? (record.values.get('SpecialEffect') ?? 0) : 0;
	const styleOf: Record<number, string> = { 0: 'none', 1: 'outset', 2: 'inset', 3: 'groove', 6: 'ridge' };
	const line = styleOf[effect];
	if (line === undefined) { return ''; }
	return line === 'none'
		? 'border:none;box-shadow:none;'
		: `border:2px ${line} #d9d9d9;box-shadow:none;`;
}

/** The GrayText the face wears when a control is disabled. */
function stateCss(record: ParsedRecord | undefined, kind: string): string {
	if (!record) { return ''; }
	const vpb = effectiveVariousPropertyBits(record, kind);
	if (vpb !== undefined && (vpb & 0x2) === 0) { return `color:${WINDOWS_PALETTE.GrayText};`; }
	return '';
}

// A GuidAndPicture is a 16-byte CLSID, a 4-byte preamble, a 4-byte size,
// then the picture bytes in their ORIGINAL format. Browsers paint BMP, PNG,
// JPEG, GIF, and ICO from data URIs; WMF and EMF have no browser decoder,
// so those keep the honest hatched placeholder.
function pictureDataUri(guidAndPicture: Buffer | undefined): string | undefined {
	if (!guidAndPicture || guidAndPicture.length <= 24) { return undefined; }
	const bytes = guidAndPicture.subarray(24);
	const mime = bytes[0] === 0x42 && bytes[1] === 0x4d ? 'image/bmp'
		: bytes[0] === 0x89 && bytes[1] === 0x50 ? 'image/png'
			: bytes[0] === 0xff && bytes[1] === 0xd8 ? 'image/jpeg'
				: bytes.subarray(0, 4).toString('latin1') === 'GIF8' ? 'image/gif'
					: bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 1 && bytes[3] === 0 ? 'image/x-icon'
						: undefined;
	return mime ? `data:${mime};base64,${Buffer.from(bytes).toString('base64')}` : undefined;
}

const PICTURE_ALIGNMENTS: Readonly<Record<number, string>> = {
	0: 'left top', 1: 'right top', 2: 'center center', 3: 'left bottom', 4: 'right bottom',
};

/** Background CSS for a control's stored picture, honoring size mode and alignment. */
function pictureCss(record: ParsedRecord | undefined, kind: string): string {
	const uri = pictureDataUri(record?.streamData.get('Picture'));
	if (!uri || !record) { return ''; }
	if (kind !== 'Image') {
		// Only fmPicturePosition 12 - the picture BEHIND the caption - is a
		// background; the other eleven positions render as a real <img> the
		// runtime script dresses (see captionPictureHtml).
		const position = recordHas(record, 'PicturePosition') ? (record.values.get('PicturePosition') ?? 7) : 7;
		if (position !== 12) { return ''; }
		return `background-image:url('${uri}');background-repeat:no-repeat;background-position:center center;`;
	}
	const mode = recordHas(record, 'PictureSizeMode') ? (record.values.get('PictureSizeMode') ?? 0) : 0;
	const align = recordHas(record, 'PictureAlignment') ? (record.values.get('PictureAlignment') ?? 2) : 2;
	const size = mode === 1 ? '100% 100%' : mode === 3 ? 'contain' : 'auto';
	const position = PICTURE_ALIGNMENTS[align] ?? 'center center';
	return `background-image:url('${uri}');background-repeat:no-repeat;`
		+ `background-size:${size};background-position:${position};`;
}

/**
 * A picture that sits WITH a caption, as an <img> the runtime script dresses.
 * MSForms draws these through a color key and stretches an oversized one over
 * the whole face (measured in xlide vbide off the running form) - decisions
 * that need the picture's natural size, which only a loaded <img> knows.
 */
function captionPictureHtml(
	record: ParsedRecord | undefined,
	kind: string,
	pictures: Map<string, string>,
): string {
	if (!record || kind === 'Image') { return ''; }
	const uri = pictureDataUri(record.streamData.get('Picture'));
	if (!uri) { return ''; }
	const position = recordHas(record, 'PicturePosition') ? (record.values.get('PicturePosition') ?? 7) : 7;
	if (position === 12) { return ''; }
	const key = `p${pictures.size}`;
	pictures.set(key, uri);
	return `<img class="cpic" data-pic="${key}" data-pos="${position}" draggable="false">`;
}

function colorCss(record: ParsedRecord | undefined, kind?: string): string {
	if (!record) { return ''; }
	const parts: string[] = [];
	const back = record.values.get('BackColor');
	if (back !== undefined && recordHas(record, 'BackColor')) {
		parts.push(`background:${cssColor(back)};`);
	} else if (kind) {
		const vpb = effectiveVariousPropertyBits(record, kind);
		if (vpb !== undefined && (vpb & 0x8) !== 0) {
			parts.push(`background:${WINDOWS_PALETTE.ButtonFace};`);
		}
	}
	const fore = record.values.get('ForeColor');
	if (fore !== undefined && recordHas(record, 'ForeColor')) { parts.push(`color:${cssColor(fore)};`); }
	return parts.join('');
}

/** Renders one form package to the inner HTML of its surface. */
function renderSurface(
	pkg: FormPackage,
	idPrefix: string,
	selected: string | undefined,
	pictures: Map<string, string>,
): string {
	const parts: string[] = [];
	pkg.entries.forEach((entry, index) => {
		const site = entry.site;
		const record = entry.kind === 'record' ? entry.record : undefined;
		const kind = controlKindOfSite(site, record);
		const name = siteName(site);
		const inner = siteIsContainer(site) ? pkg.containers.get(siteId(site)) : undefined;
		const box = siteBox(site, record, inner);
		const style = `left:${pts(box.left)}pt;top:${pts(box.top)}pt;width:${pts(box.width)}pt;height:${pts(box.height)}pt;`
			+ fontCss(record) + colorCss(record, kind) + stateCss(record, kind) + pictureCss(record, kind)
			+ borderCss(record)
			+ (kind === 'TextBox' && record && ((effectiveVariousPropertyBits(record, kind) ?? 0) & 0x80000000) !== 0
				? 'white-space:pre-wrap;'
				: '');
		const sel = selected !== undefined && selected.toLowerCase() === name.toLowerCase() ? ' selected' : '';
		const dn = `data-name="${esc(name)}"`;
		const caption = record?.strings.get('Caption')?.text
			?? inner?.form.record.strings.get('Caption')?.text
			?? '';
		const value = record?.strings.get('Value')?.text ?? '';
		const childId = `${idPrefix}-${index}`;

		switch (kind) {
			case 'Label':
				parts.push(`<div class="ctl label${sel}" ${dn} style="${style}" title="${esc(name)}">${captionPictureHtml(record, kind, pictures)}${esc(caption)}</div>`);
				break;
			case 'TextBox':
				parts.push(`<div class="ctl edit${sel}" ${dn} style="${style}" title="${esc(name)}">${esc(value)}</div>`);
				break;
			case 'ComboBox':
				parts.push(`<div class="ctl edit combo${sel}" ${dn} style="${style}" title="${esc(name)}"><span>${esc(value)}</span><span class="drop">&#9662;</span></div>`);
				break;
			case 'ListBox':
				parts.push(`<div class="ctl edit list${sel}" ${dn} style="${style}" title="${esc(name)}"></div>`);
				break;
			case 'CheckBox':
				parts.push(`<div class="ctl opt${sel}" ${dn} style="${style}" title="${esc(name)}"><span class="box${value === '1' ? ' on' : ''}"></span>${esc(caption)}</div>`);
				break;
			case 'OptionButton':
				parts.push(`<div class="ctl opt${sel}" ${dn} style="${style}" title="${esc(name)}"><span class="radio${value === '1' ? ' on' : ''}"></span>${esc(caption)}</div>`);
				break;
			case 'CommandButton':
			case 'ToggleButton':
				parts.push(`<div class="ctl button${kind === 'ToggleButton' && value === '1' ? ' pressed' : ''}${sel}" ${dn} style="${style}" title="${esc(name)}">${captionPictureHtml(record, kind, pictures)}${esc(caption)}</div>`);
				break;
			case 'Image':
				parts.push(pictureDataUri(record?.streamData.get('Picture'))
					? `<div class="ctl image pictured${sel}" ${dn} style="${style}" title="${esc(name)}"></div>`
					: `<div class="ctl image${sel}" ${dn} style="${style}" title="${esc(name)}"><span>${esc(name)}</span></div>`);
				break;
			case 'SpinButton':
				parts.push(`<div class="ctl spin${sel}" ${dn} style="${style}" title="${esc(name)}"><span>&#9652;</span><span>&#9662;</span></div>`);
				break;
			case 'ScrollBar':
				parts.push(`<div class="ctl scroll${sel}" ${dn} style="${style}" title="${esc(name)}"></div>`);
				break;
			case 'TabStrip': {
				const items = record?.arrays.get('Items');
				const tabs = items ? decodeArrayStrings(items) : [];
				parts.push(`<div class="ctl tabstrip${sel}" ${dn} style="${style}" title="${esc(name)}"><div class="tabs">${tabs.map((t, i) => `<span class="tab${i === 0 ? ' active' : ''}">${esc(t)}</span>`).join('')}</div><div class="tabbody"></div></div>`);
				break;
			}
			case 'Frame':
				parts.push(`<div class="ctl frame${sel}" ${dn} style="${style}${containerFontCss(inner)}" title="${esc(name)}"><span class="legend">${esc(caption)}</span><div class="surface" data-surface="${esc(name)}" style="${inner ? colorCss(inner.form.record) : ''}">${inner ? renderSurface(inner, childId, selected, pictures) : ''}</div></div>`);
				break;
			case 'MultiPage': {
				if (!inner) { break; }
				const tabStripEntry = inner.entries.find((e) => e.kind === 'record' && siteCacheIndex(e.site) === 18);
				const captions = tabStripEntry && tabStripEntry.kind === 'record'
					? decodeArrayStrings(tabStripEntry.record.arrays.get('Items') ?? Buffer.alloc(0))
					: [];
				const pageSites = inner.form.sites.filter((s) => siteCacheIndex(s) === 7);
				const headers = pageSites.map((pageSite, i) =>
					`<span class="tab${i === 0 ? ' active' : ''}" data-page="${childId}-p${i}">${esc(captions[i] ?? siteName(pageSite))}</span>`).join('');
				const pages = pageSites.map((pageSite, i) => {
					const pagePkg = inner.containers.get(siteId(pageSite));
					return `<div class="page" id="${childId}-p${i}" data-surface="${esc(siteName(pageSite))}"${i === 0 ? '' : ' hidden'}>${pagePkg ? renderSurface(pagePkg, `${childId}-p${i}`, selected, pictures) : ''}</div>`;
				}).join('');
				parts.push(`<div class="ctl multipage${sel}" ${dn} style="${style}${containerFontCss(inner)}" title="${esc(name)}"><div class="tabs">${headers}</div><div class="pagearea">${pages}</div></div>`);
				break;
			}
			default:
				parts.push(`<div class="ctl foreign${sel}" ${dn} style="${style}" title="${esc(name)}"><span>${esc(name)}</span></div>`);
				break;
		}
	});
	return parts.join('\n');
}

export interface FormPreviewOptions {
	formName: string;
	caption?: string;
	/** Control name to render selected, restoring selection across re-renders. */
	selected?: string;
	/** False renders a static picture; true adds the designer interactions. */
	interactive?: boolean;
	/** Property rows per target ('' is the form), rendered in the pane. */
	properties?: Record<string, { kind: string; rows: Array<{ prop: string; value: string }> }>;
	/** Workbook path and module name, stamped into webview state so VS Code
	 *  can restore the panel after a window reload. */
	identity?: { workbook: string; module: string };
}

/**
 * The whole preview document: the form's dialog chrome, its control tree,
 * and the few lines of script that switch MultiPage pages. Self-contained -
 * no external resources - so it drops straight into a webview.
 */
export function renderFormPreviewHtml(pkg: FormPackage, options: FormPreviewOptions): string {
	const record = pkg.form.record;
	const size = record.sizes.get('DisplayedSize') ?? { width: 0, height: 0 };
	const width = pts(size.width);
	const height = pts(size.height);
	const caption = options.caption ?? record.strings.get('Caption')?.text ?? options.formName;
	const back = record.values.get('BackColor');
	const formBack = back !== undefined && recordHas(record, 'BackColor') ? cssColor(back) : WINDOWS_PALETTE.ButtonFace;
	const formPictureUri = pictureDataUri(pkg.form.pictureRaw);
	const formPicture = formPictureUri
		? `background-image:url('${formPictureUri}');background-repeat:no-repeat;background-position:center center;`
		: '';
	const fore = record.values.get('ForeColor');
	const formFore = fore !== undefined && recordHas(record, 'ForeColor') ? cssColor(fore) : '#000';
	const formBorder = (() => {
		const color = recordHas(record, 'BorderColor') ? cssColor(record.values.get('BorderColor') ?? 0) : '#7a7a7a';
		if (recordHas(record, 'BorderStyle') && (record.values.get('BorderStyle') ?? 0) === 1) {
			return `border:1px solid ${color};`;
		}
		const styleOf: Record<number, string> = { 1: 'outset', 2: 'inset', 3: 'groove', 6: 'ridge' };
		const effect = recordHas(record, 'SpecialEffect') ? (record.values.get('SpecialEffect') ?? 0) : 0;
		return styleOf[effect] ? `border:2px ${styleOf[effect]} #d9d9d9;` : '';
	})();
	const scrollBars = recordHas(record, 'ScrollBars') ? (record.values.get('ScrollBars') ?? 0) : 0;
	const scrollRails = `${(scrollBars & 2) !== 0 ? '<div class="rail v"></div>' : ''}${(scrollBars & 1) !== 0 ? '<div class="rail h"></div>' : ''}`;
	const formZoom = recordHas(record, 'Zoom') ? Math.max(10, Math.min(400, record.values.get('Zoom') ?? 100)) : 100;
	const stdFont = pkg.form.fontRaw ? parseStdFont(pkg.form.fontRaw) : undefined;
	const formFont = stdFont
		? `font-family:'${esc(stdFont.face)}',Tahoma,sans-serif;font-size:${stdFont.heightTenThousandthsPt / 10000}pt;`
		: "font-family:Tahoma,sans-serif;font-size:8.25pt;";

	const toolbox = ['Label', 'TextBox', 'ComboBox', 'ListBox', 'CheckBox', 'OptionButton',
		'ToggleButton', 'Frame', 'CommandButton', 'TabStrip', 'ScrollBar', 'SpinButton', 'Image'];
	const interactive = options.interactive !== false;
	const propsJson = JSON.stringify(options.properties ?? {}).replace(/</g, '\\u003c');
	const pictures = new Map<string, string>();
	const surfaceHtml = renderSurface(pkg, 'c', options.selected, pictures);
	const picturesJson = JSON.stringify(Object.fromEntries(pictures)).replace(/</g, '\\u003c');
	return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
	:root { color-scheme: light; }
	body { margin: 0; padding: 0; background: #808080; user-select: none; }
	.toolbar { position: sticky; top: 0; z-index: 10; display: flex; flex-wrap: wrap; gap: 4px;
		align-items: center; background: #2d2d2d; color: #ccc; padding: 6px 10px;
		font: 12px sans-serif; }
	.toolbar .tool { background: #3c3c3c; color: #ddd; border: 1px solid #555; padding: 2px 8px;
		border-radius: 3px; cursor: pointer; }
	/* The blue is ROLLOVER feedback, not a latch: an armed tool shows no
	   fill (the canvas's crosshair announces placing mode instead). */
	.toolbar .tool:hover { background: #0e639c; color: #fff; border-color: #1177bb; }
	.toolbar label { display: flex; gap: 4px; align-items: center; margin-left: 8px; }
	.toolbar select { background: #3c3c3c; color: #ddd; border: 1px solid #555;
		border-radius: 3px; font: inherit; }
	.main { display: flex; align-items: flex-start; }
	.stage { padding: 24px; flex: 1; min-width: 0; }
	/* The pane floats: the aside's gutter wears the designer's own gray, and
	   the content is a dark card with its own edge and shadow. */
	.props { width: 240px; flex: none; position: sticky; top: 40px; box-sizing: border-box;
		max-height: calc(100vh - 48px); overflow-y: auto; background: #808080; color: #ccc;
		font: 12px sans-serif; padding: 10px; }
	#propsBody { background: #252526; border: 1px solid #3c3c3c; border-radius: 4px;
		box-shadow: 0 2px 8px rgba(0, 0, 0, 0.35); overflow: hidden; padding-bottom: 6px; }
	.props .props-head { padding: 6px 10px; font-weight: bold; background: #2d2d2d;
		position: sticky; top: 0; display: flex; justify-content: space-between; align-items: center; }
	.props .props-sash { position: absolute; left: 0; top: 0; bottom: 0; width: 5px;
		cursor: col-resize; z-index: 2; }
	.props .props-sash:hover { background: rgba(14, 99, 156, 0.4); }
	.props-collapse { background: none; border: none; color: #ccc; cursor: pointer;
		font: inherit; padding: 0 2px; }
	.props-collapse:hover { color: #fff; }
	.props.collapsed { width: 24px !important; overflow: hidden; }
	.props.collapsed .row, .props.collapsed .props-head span, .props.collapsed .props-sash { display: none; }
	.props .row { display: grid; grid-template-columns: 45% 55%; align-items: center;
		padding: 1px 8px 1px 10px; column-gap: 4px; }
	.props .row:hover { background: #2a2d2e; }
	.props label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
		cursor: default; }
	.props input { width: 100%; box-sizing: border-box; background: #3c3c3c; color: #ddd;
		border: 1px solid transparent; padding: 2px 4px; font: inherit; }
	.props input:focus { border-color: #0e639c; outline: none; }
	.props input:disabled { color: #888; }
	.props select { width: 100%; box-sizing: border-box; background: #3c3c3c; color: #ddd;
		border: 1px solid transparent; padding: 2px 2px; font: inherit; }
	.props select:focus { border-color: #0e639c; outline: none; }
	.props .colorcell { display: flex; gap: 4px; align-items: center; }
	.props .swatch { width: 16px; height: 16px; flex: none; border: 1px solid #555;
		cursor: pointer; background: transparent; padding: 0; }
	.colorpop { position: fixed; z-index: 30; background: #2d2d2d; border: 1px solid #555;
		border-radius: 4px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.5); padding: 8px; width: 208px; }
	.colorpop .grid { display: grid; grid-template-columns: repeat(8, 1fr); gap: 3px; }
	.colorpop .grid button { width: 20px; height: 16px; border: 1px solid #555; cursor: pointer;
		padding: 0; }
	.colorpop .grid button:hover { outline: 2px solid #0e639c; }
	.colorpop .sys { margin-top: 6px; max-height: 120px; overflow-y: auto; }
	.colorpop .sys button { display: flex; gap: 6px; align-items: center; width: 100%;
		background: none; border: none; color: #ccc; cursor: pointer; font: inherit;
		padding: 1px 2px; text-align: left; }
	.colorpop .sys button:hover { background: #3c3c3c; }
	.colorpop .sys i { width: 12px; height: 12px; flex: none; border: 1px solid #555;
		display: inline-block; }
	.dialog { width: ${width}pt; box-shadow: 2px 2px 8px rgba(0,0,0,.5); position: relative; }
	.dialog.form-selected { outline: 1px dashed #0e639c; outline-offset: 2px; }
	.form-handle { position: absolute; width: 7px; height: 7px; background: #fff;
		border: 1px solid #0e639c; z-index: 7; }
	.titlebar { background: #fff; color: #000; padding: 4px 8px; border-bottom: 1px solid #e5e5e5;
		font: 9pt 'Segoe UI', Tahoma, sans-serif; display: flex; justify-content: space-between; }
	.client { position: relative; width: ${width}pt; height: ${height}pt;
		background: ${formBack}; overflow: hidden; color: ${formFore}; ${formBorder}${formPicture}${formFont} }
	.rail { position: absolute; background: #d4d0c8; border: 1px solid #a0a0a0; z-index: 3;
		pointer-events: none; }
	.rail.v { right: 0; top: 0; bottom: 0; width: 11pt; }
	.rail.h { left: 0; right: 0; bottom: 0; height: 11pt; }
	img.cpic { max-width: none; max-height: none; }
	/* While grid snap is on, every design surface shows the 6pt lattice the
	   snapping answers to, as the VBE's dotted face does. The half-cell
	   offset centers a dot on each grid point, so dots mark exactly where a
	   snapped edge lands. */
	body.grid-on [data-surface]::before { content: ''; position: absolute; inset: 0;
		background-image: radial-gradient(circle, #666 1px, transparent 1px);
		background-size: 6pt 6pt; background-position: -3pt -3pt; pointer-events: none; }
	.ctl { position: absolute; box-sizing: border-box; overflow: hidden; white-space: nowrap; }
	.ctl.selected { outline: 1px dashed #0e639c; outline-offset: 1px; }
	/* Hover ergonomics, adopted from the vbide designer (settled 2026-08-15):
	   the control a click would select lights up, and only the DEEPEST hovered
	   one does - without :has, a Frame outlines under its own children. The
	   HAND across the whole face because every inch responds to a press; the
	   four-way MOVE only on the control a press would actually pick up and
	   carry; a handle keeps its own resize cursor; a drag in flight paints
	   the whole canvas with its gesture's cursor. */
	.ctl:hover:not(:has(.ctl:hover)) { outline: 1px solid rgba(74, 106, 157, 0.8); outline-offset: 1px; }
	.client, .ctl { cursor: pointer; }
	.ctl.selected { cursor: move; }
	.ctl.dragging { z-index: 5; opacity: 0.85; pointer-events: none; }
	.client.gesture-move, .client.gesture-move * { cursor: move !important; }
	.client.gesture-resize, .client.gesture-resize * { cursor: inherit; }
	.label { display: flex; align-items: flex-start; }
	.edit { background: #fff; color: #000; border: 1px solid #7a7a7a; padding: 1px 2px;
		box-shadow: inset 1px 1px 2px rgba(0, 0, 0, 0.08); }
	.combo { display: flex; justify-content: space-between; align-items: center; }
	.combo .drop { border-left: 1px solid #a0a0a0; background: #f0f0f0; align-self: stretch;
		display: flex; align-items: center; padding: 0 2px; }
	.opt { display: flex; align-items: center; gap: 4px; }
	.opt .box { width: 9pt; height: 9pt; background: #fff; border: 1px solid #7a7a7a; flex: none; }
	.opt .box.on::after { content: '\\2713'; display: block; text-align: center; line-height: 9pt; }
	.opt .radio { width: 9pt; height: 9pt; background: #fff; border: 1px solid #7a7a7a;
		border-radius: 50%; flex: none; }
	.opt .radio.on { background: radial-gradient(circle at center, #000 35%, #fff 40%); }
	.button { background: linear-gradient(#f6f6f6, #e8e8e8); border: 1px solid #8b8b8b;
		border-radius: 2px; box-shadow: 0 1px 1px rgba(0, 0, 0, 0.10);
		display: flex; align-items: center; justify-content: center; }
	.button.pressed { border-style: inset; }
	.image, .foreign { border: 1px solid #a0a0a0;
		background: repeating-linear-gradient(45deg, #ddd 0 6px, #eee 6px 12px);
		display: flex; align-items: center; justify-content: center; color: #555; font-size: 7pt; }
	.image.pictured { background: #fff; }
	.spin { display: flex; flex-direction: column; }
	.spin span { flex: 1; background: #f0f0f0; border: 1px solid #a0a0a0;
		display: flex; align-items: center; justify-content: center; font-size: 6pt; }
	.scroll { background: #d4d0c8; border: 1px solid #a0a0a0; position: relative; }
	.scroll::after { content: ''; position: absolute; left: 1px; right: 1px; top: 15%; height: 30%;
		background: #f0f0f0; border: 1px solid #808080; }
	.frame { border: 1px solid #bdbdbd; box-shadow: inset 0 0 0 1px #fbfbfb;
		overflow: visible; }
	.frame .legend { position: absolute; top: 0; left: 6pt; transform: translateY(-55%);
		background: ${formBack}; padding: 0 3px; line-height: 1.1; z-index: 1; }
	.frame .surface { position: absolute; inset: 0; overflow: hidden; background: #f0f0f0; }
	.tabs { display: flex; gap: 1px; padding: 0 2px; height: 14pt; align-items: flex-end;
		position: relative; z-index: 1; }
	.tab { background: #f0f0f0; border: 1px solid #a0a0a0; border-bottom: none;
		padding: 0 8px 1px; border-radius: 3px 3px 0 0; cursor: pointer; }
	.tab.active { background: #fff; position: relative; top: 1px; }
	.tabstrip, .multipage { background: #f0f0f0; overflow: visible; }
	.tabbody, .pagearea { position: absolute; inset: 14pt 0 0; border: 1px solid #bdbdbd;
		box-shadow: inset 0 0 0 1px #fbfbfb; background: #f0f0f0; overflow: hidden; }
	.page { position: absolute; inset: 0; background: #f0f0f0; }
	.handle { position: absolute; width: 6px; height: 6px; background: #fff;
		border: 1px solid #0e639c; z-index: 5; }
	.guide { position: absolute; background: #e51400; z-index: 4; pointer-events: none; }
	.guide.v { width: 1px; top: 0; bottom: 0; }
	.guide.h { height: 1px; left: 0; right: 0; }
	.cpic { flex: none; pointer-events: none; }
	.diag { color: #ff6b6b; margin-left: 12px; font-weight: bold; }
	.placing, .placing .ctl { cursor: crosshair !important; }
	.ghost { position: absolute; border: 1px dashed #0e639c; background: rgba(14, 99, 156, 0.12);
		z-index: 6; pointer-events: none; }
	.drop-target { outline: 2px solid #0e639c; outline-offset: -2px; }
	/* The vbide splitter chrome, docked at the bottom edge: an arrow either
	   side of the dots collapses one half or the other, and the dots DRAG
	   the split itself. */
	.split-grip { position: fixed; left: 0; right: 0; bottom: 0; height: 14px; z-index: 11;
		background: #2d2d2d; border-top: 1px solid #3c3c3c; display: flex;
		align-items: stretch; justify-content: center; }
	.grip-dots { display: flex; align-items: center; justify-content: center; gap: 3px;
		padding: 0 18px; cursor: ns-resize; }
	.grip-dots span { width: 3px; height: 3px; border-radius: 50%; background: #777; }
	.grip-dots:hover span { background: #ccc; }
	.grip-btn { background: none; border: none; color: #777; cursor: pointer;
		font-size: 8px; padding: 0 8px; }
	.grip-btn:hover { color: #fff; }
</style>
</head>
<body>
${interactive ? `	<div class="toolbar" id="toolbar">
		<span>Add:</span>
		${toolbox.map((k) => `<button class="tool" data-kind="${k}">${k}</button>`).join('')}
		<label><input type="checkbox" id="snapGrid" checked>Grid 6pt</label>
		<label><input type="checkbox" id="snapNeighbors">Snap to neighbors</label>
		<label>Zoom <select id="zoomPick">
			<option value="0.5">50%</option><option value="0.75">75%</option>
			<option value="1" selected>100%</option><option value="1.25">125%</option>
			<option value="1.5">150%</option><option value="2">200%</option>
		</select> <span id="zoomShow"></span></label>
	</div>` : ''}
	<div class="main">
	<div class="stage">
	<div class="dialog${options.selected === '' ? ' form-selected' : ''}">
		<div class="titlebar"><span>${esc(caption)}</span><span>&#10005;</span></div>
		<div class="client" data-surface="">${scrollRails}
${surfaceHtml}
		</div>
	</div>
	</div>
${interactive ? `	<aside class="props" id="props"><div class="props-sash" id="propsSash"></div><div id="propsBody"></div></aside>` : ''}
	</div>
${interactive ? `	<div class="split-grip" id="splitGrip"><button class="grip-btn" id="gripUp" title="Push the split up: the markup below takes the designer's space">&#9650;</button><div class="grip-dots" id="gripDots" title="Drag to resize the split between the designer and the markup below"><span></span><span></span><span></span><span></span><span></span></div><button class="grip-btn" id="gripDown" title="Push the split down: the designer takes the markup's space">&#9660;</button></div>` : ''}
	<script type="application/json" id="cpicData">${picturesJson}</script>
	<script>
		// Caption pictures, the way MSForms actually draws them (measured in
		// xlide vbide off the running form, 2026-08-17): the top-left pixel
		// is a COLOR KEY - exactly-matching pixels go transparent, and the
		// anti-aliased halo that misses the key stays, because it does on the
		// real form too. A picture too big for its control is STRETCHED over
		// the whole face with the caption underneath; one that fits keeps its
		// natural size beside the caption where fmPicturePosition says.
		// Surface pictures (Image, the form) draw the same pixels SOLID -
		// MSForms treats the two surfaces differently, so this canvas does.
		const cpicData = (() => {
			try { return JSON.parse(document.getElementById('cpicData')?.textContent ?? '{}'); }
			catch { return {}; }
		})();
		document.querySelectorAll('img.cpic').forEach((image) => {
			const box = image.parentElement;
			const position = Number(image.dataset.pos) || 7;
			// A picture that cannot load must not take its control with it:
			// the img leaves, the chrome and caption stay, and the console
			// says why.
			image.addEventListener('error', () => {
				console.error('xlide designer: caption picture failed to load for', box?.getAttribute('data-name') ?? '(unknown)');
				image.remove();
			});
			image.src = cpicData[image.dataset.pic] ?? '';
			box.style.flexDirection = position <= 5 ? 'row' : 'column';
			box.style.alignItems = ['flex-start', 'center', 'flex-end'][position % 3] || 'center';
			if (position <= 2 || (position >= 6 && position <= 8)) { box.prepend(image); }
			const keyOut = () => {
				if (image.naturalWidth === 0 || image.dataset.keyed === 'yes') { return; }
				const sheet = document.createElement('canvas');
				sheet.width = image.naturalWidth;
				sheet.height = image.naturalHeight;
				const ink = sheet.getContext('2d');
				if (!ink) { return; }
				ink.drawImage(image, 0, 0);
				let field;
				try { field = ink.getImageData(0, 0, sheet.width, sheet.height); } catch { return; }
				const px = field.data;
				const r = px[0], g = px[1], b = px[2];
				for (let at = 0; at < px.length; at += 4) {
					if (px[at] === r && px[at + 1] === g && px[at + 2] === b) { px[at + 3] = 0; }
				}
				ink.putImageData(field, 0, 0);
				image.dataset.keyed = 'yes';
				image.src = sheet.toDataURL('image/png');
			};
			const stretchIfOversized = () => {
				const room = box.getBoundingClientRect();
				if (room.width <= 0 || room.height <= 0 || image.naturalWidth === 0) { return; }
				if (image.naturalWidth > room.width || image.naturalHeight > room.height) {
					if (getComputedStyle(box).position === 'static') { box.style.position = 'relative'; }
					image.style.position = 'absolute';
					image.style.inset = '0';
					image.style.width = '100%';
					image.style.height = '100%';
					image.style.objectFit = 'fill';
				}
			};
			const dress = () => { keyOut(); stretchIfOversized(); };
			image.addEventListener('load', dress);
			if (image.complete) { dress(); }
		});

		document.querySelectorAll('.multipage .tabs .tab').forEach((tab) => {
			tab.addEventListener('click', (e) => {
				if (document.body.dataset.dragging) { return; }
				const area = tab.closest('.multipage');
				area.querySelectorAll(':scope > .tabs > .tab').forEach((t) => t.classList.remove('active'));
				tab.classList.add('active');
				area.querySelectorAll(':scope > .pagearea > .page').forEach((p) => { p.hidden = true; });
				const page = document.getElementById(tab.dataset.page);
				if (page) { page.hidden = false; }
			});
		});
	</script>
${interactive ? `	<script>
	(() => {
		const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : undefined;
		const post = (msg) => { if (vscode) { vscode.postMessage(msg); } };

		// The panel's identity rides in webview state, so a window reload can
		// deserialize the designer back onto its form.
		if (vscode?.setState) {
			vscode.setState({ ...(vscode.getState?.() ?? {}), ...${JSON.stringify({
				wb: options.identity?.workbook ?? '',
				mod: options.identity?.module ?? '',
			})} });
		}

		// The Properties pane, following the selection the way the VBE's
		// Properties window does - '' is the form itself. Rows arrive baked
		// from the engine (the same fields the markup prints); committing an
		// edit posts ONE property write, and the canvas re-renders from the
		// bytes. Escape in a row reverts it without touching the selection.
		const PROPS = ${propsJson};
		const FORM_NAME = ${JSON.stringify(options.formName)};
		// The arrows are TOGGLES, as in xlide vbide: a collapse flips the
		// glyph, and the flipped arrow restores the split it replaced. The
		// panel owns the saved sizes and answers with the current state,
		// which also survives re-renders (the canvas asks on every load).
		const gripUp = document.getElementById('gripUp');
		const gripDown = document.getElementById('gripDown');
		const paintSplitState = (collapsed) => {
			if (gripUp) {
				gripUp.innerHTML = collapsed === 'self' ? '&#9660;' : '&#9650;';
				gripUp.title = collapsed === 'self'
					? 'Restore the split'
					: 'Push the split up: the markup below takes the space';
			}
			if (gripDown) {
				gripDown.innerHTML = collapsed === 'below' ? '&#9650;' : '&#9660;';
				gripDown.title = collapsed === 'below'
					? 'Restore the split'
					: 'Push the split down: the designer takes the space';
			}
		};
		window.addEventListener('message', (e) => {
			const m = e.data;
			if (m && m.type === 'splitState') { paintSplitState(m.collapsed ?? null); }
		});
		gripUp?.addEventListener('click', () => post({ type: 'splitCollapse', which: 'self' }));
		gripDown?.addEventListener('click', () => post({ type: 'splitCollapse', which: 'below' }));
		post({ type: 'splitStateQuery' });
		// The dots drag with POINTER CAPTURE (a downward drag leaves the
		// iframe within millimeters, and uncaptured events stop at the edge),
		// and every way a pointer can abandon a drag ends it - pointercancel,
		// capture loss, window blur. The panel does the actual sizing with
		// layout-tree pixel math, so the motion is CONTINUOUS and the
		// direction is the mouse's by construction; deltas stream a few
		// pixels apart.
		const gripDots = document.getElementById('gripDots');
		let gripDrag = null;
		const endGripDrag = () => {
			if (!gripDrag) { return; }
			gripDrag = null;
			post({ type: 'splitDrag', phase: 'end' });
		};
		gripDots?.addEventListener('pointerdown', (e) => {
			gripDrag = { y: e.clientY, lastSent: 0, delta: 0, raf: 0 };
			try { gripDots.setPointerCapture(e.pointerId); } catch { /* capture is a nicety */ }
			post({ type: 'splitDrag', phase: 'start' });
			e.preventDefault();
		});
		gripDots?.addEventListener('pointermove', (e) => {
			if (!gripDrag) { return; }
			// One send per FRAME, latest position wins - the panel coalesces
			// on its side too, so the border tracks without flooding.
			gripDrag.delta = e.clientY - gripDrag.y;
			if (gripDrag.raf) { return; }
			gripDrag.raf = requestAnimationFrame(() => {
				if (!gripDrag) { return; }
				gripDrag.raf = 0;
				if (gripDrag.delta !== gripDrag.lastSent) {
					gripDrag.lastSent = gripDrag.delta;
					post({ type: 'splitDrag', phase: 'move', delta: gripDrag.delta });
				}
			});
		});
		gripDots?.addEventListener('pointerup', endGripDrag);
		gripDots?.addEventListener('pointercancel', endGripDrag);
		gripDots?.addEventListener('lostpointercapture', endGripDrag);
		window.addEventListener('blur', endGripDrag);
		const propsPane = document.getElementById('props');
		const propsBody = document.getElementById('propsBody');
		const mergeState = (patch) => {
			if (vscode?.setState) {
				vscode.setState({ ...(vscode.getState?.() ?? {}), ...patch });
			}
		};
		const paneState = vscode?.getState?.() ?? {};
		let propsCollapsed = paneState.propsCollapsed === true;
		let propsWidth = typeof paneState.propsWidth === 'number' ? paneState.propsWidth : 240;
		const applyPropsChrome = () => {
			if (!propsPane) { return; }
			propsPane.classList.toggle('collapsed', propsCollapsed);
			propsPane.style.width = propsCollapsed ? '' : propsWidth + 'px';
		};
		applyPropsChrome();
		const BOOL_PROPS = new Set(['Enabled', 'Locked', 'MultiLine', 'WordWrap', 'AutoSize',
			'Visible', 'TabStop', 'Default', 'Cancel', 'Font.Bold', 'Font.Italic',
			'Font.Underline', 'Font.Strikethrough', 'ShowModal', 'WhatsThisButton']);
		const ENUM_OPTIONS = {
			SpecialEffect: [['0', 'Flat'], ['1', 'Raised'], ['2', 'Sunken'], ['3', 'Etched'], ['6', 'Bump']],
			BorderStyle: [['0', 'None'], ['1', 'Single']],
			MultiSelect: [['0', 'Single'], ['1', 'Multi'], ['2', 'Extended']],
			ListStyle: [['0', 'Plain'], ['1', 'Option']],
			Style: [['0', 'DropDownCombo'], ['2', 'DropDownList']],
			PictureSizeMode: [['0', 'Clip'], ['1', 'Stretch'], ['3', 'Zoom']],
			PictureAlignment: [['0', 'TopLeft'], ['1', 'TopRight'], ['2', 'Center'], ['3', 'BottomLeft'], ['4', 'BottomRight']],
			Orientation: [['-1', 'Auto'], ['0', 'Vertical'], ['1', 'Horizontal']],
			ScrollBars: [['0', 'None'], ['1', 'Horizontal'], ['2', 'Vertical'], ['3', 'Both']],
			TextAlign: [['Left', 'Left'], ['Center', 'Center'], ['Right', 'Right']],
			Cycle: [['0', 'AllForms'], ['2', 'CurrentForm']],
			Alignment: [['0', 'Left'], ['1', 'Right']],
			MousePointer: [['0', 'Default'], ['1', 'Arrow'], ['2', 'Cross'], ['3', 'IBeam'],
				['6', 'SizeNESW'], ['7', 'SizeNS'], ['8', 'SizeNWSE'], ['9', 'SizeWE'],
				['10', 'UpArrow'], ['11', 'HourGlass'], ['12', 'NoDrop'], ['13', 'AppStarting'],
				['14', 'Help'], ['15', 'SizeAll'], ['99', 'Custom']],
			StartUpPosition: [['0', 'Manual'], ['1', 'CenterOwner'], ['2', 'CenterScreen'], ['3', 'WindowsDefault']],
		};
		const COLOR_PROPS = new Set(['BackColor', 'ForeColor', 'BorderColor']);
		const SYSTEM_COLORS = ${JSON.stringify(WINDOWS_PALETTE)};
		const FONT_FACES = ['Tahoma', 'Segoe UI', 'Arial', 'Calibri', 'Verdana', 'Georgia',
			'Times New Roman', 'Courier New', 'Consolas', 'MS Sans Serif'];
		// vbide's ramp: a greys row, then eight hues by seven lightness steps.
		const paletteSwatches = (() => {
			const rows = [['#000000', '#404040', '#808080', '#a6a6a6', '#c0c0c0', '#d9d9d9', '#f0f0f0', '#ffffff']];
			const hsl = (h, sPct, l) => {
				const a = (sPct * Math.min(l, 100 - l)) / 100;
				const ch = (n) => {
					const k = (n + h / 30) % 12;
					const v = l - a * Math.max(-1, Math.min(k - 3, Math.min(9 - k, 1)));
					return Math.round((v * 255) / 100).toString(16).padStart(2, '0');
				};
				return '#' + ch(0) + ch(8) + ch(4);
			};
			for (const l of [22, 32, 42, 52, 62, 74, 86]) {
				rows.push([0, 30, 60, 120, 180, 210, 270, 330].map((h) => hsl(h, 68, l)));
			}
			return rows.flat();
		})();
		const cssOfColorValue = (value) => value.startsWith('#') ? value : (SYSTEM_COLORS[value] ?? 'transparent');
		let colorPop = null;
		const closeColorPop = () => { colorPop?.remove(); colorPop = null; };
		document.addEventListener('pointerdown', (e) => {
			if (colorPop && !e.target.closest('.colorpop') && !e.target.closest('.swatch')) { closeColorPop(); }
		}, true);
		const openColorPop = (anchor, apply) => {
			closeColorPop();
			colorPop = document.createElement('div');
			colorPop.className = 'colorpop';
			const grid = document.createElement('div');
			grid.className = 'grid';
			for (const hex of paletteSwatches) {
				const b = document.createElement('button');
				b.style.background = hex;
				b.title = hex;
				b.addEventListener('click', () => { apply(hex); closeColorPop(); });
				grid.appendChild(b);
			}
			const sys = document.createElement('div');
			sys.className = 'sys';
			for (const name of Object.keys(SYSTEM_COLORS)) {
				const b = document.createElement('button');
				const chip = document.createElement('i');
				chip.style.background = SYSTEM_COLORS[name];
				b.append(chip, document.createTextNode(name));
				b.addEventListener('click', () => { apply(name); closeColorPop(); });
				sys.appendChild(b);
			}
			colorPop.append(grid, sys);
			document.body.appendChild(colorPop);
			const at = anchor.getBoundingClientRect();
			const width = 208;
			colorPop.style.left = Math.max(4, Math.min(window.innerWidth - width - 4, at.left - width + at.width)) + 'px';
			colorPop.style.top = Math.min(window.innerHeight - colorPop.offsetHeight - 4, at.bottom + 2) + 'px';
		};

		let lastPropsTarget = '';
		const renderProps = (target) => {
			if (!propsPane || !propsBody) { return; }
			lastPropsTarget = target;
			const info = PROPS[target];
			propsBody.textContent = '';
			if (!info) { return; }
			const head = document.createElement('div');
			head.className = 'props-head';
			const title = document.createElement('span');
			title.textContent = (target === '' ? FORM_NAME : target) + ' (' + info.kind + ')';
			const collapse = document.createElement('button');
			collapse.className = 'props-collapse';
			collapse.textContent = propsCollapsed ? '\u00ab' : '\u00bb';
			collapse.title = propsCollapsed ? 'Expand the properties pane' : 'Collapse the properties pane';
			collapse.addEventListener('click', () => {
				propsCollapsed = !propsCollapsed;
				applyPropsChrome();
				mergeState({ propsCollapsed });
				renderProps(lastPropsTarget);
			});
			head.append(title, collapse);
			propsBody.appendChild(head);
			for (const row of info.rows) {
				const div = document.createElement('div');
				div.className = 'row';
				const label = document.createElement('label');
				label.textContent = row.prop;
				label.title = row.prop;
				div.appendChild(label);
				const commit = (value) => {
					if (value === row.value) { return; }
					post({ type: 'setProp', name: target, prop: row.prop, value });
				};
				if (BOOL_PROPS.has(row.prop)) {
					const pick = document.createElement('select');
					for (const v of ['True', 'False']) {
						const o = document.createElement('option');
						o.value = v; o.textContent = v;
						pick.appendChild(o);
					}
					pick.value = row.value === 'True' ? 'True' : 'False';
					pick.addEventListener('change', () => commit(pick.value));
					div.appendChild(pick);
				} else if (ENUM_OPTIONS[row.prop]) {
					const pick = document.createElement('select');
					if (!ENUM_OPTIONS[row.prop].some(([v]) => v === row.value)) {
						const blank = document.createElement('option');
						blank.value = row.value; blank.textContent = row.value || '(default)';
						pick.appendChild(blank);
					}
					for (const [v, word] of ENUM_OPTIONS[row.prop]) {
						const o = document.createElement('option');
						o.value = v;
						o.textContent = /^[A-Za-z]/.test(v) ? word : v + ' - ' + word;
						pick.appendChild(o);
					}
					pick.value = row.value;
					pick.addEventListener('change', () => commit(pick.value));
					div.appendChild(pick);
				} else if (COLOR_PROPS.has(row.prop)) {
					const cell = document.createElement('div');
					cell.className = 'colorcell';
					const swatch = document.createElement('button');
					swatch.className = 'swatch';
					swatch.title = 'Pick a color';
					swatch.style.background = row.value ? cssOfColorValue(row.value) : 'transparent';
					const input = document.createElement('input');
					input.value = row.value;
					input.addEventListener('keydown', (e) => {
						if (e.key === 'Enter') { input.blur(); }
						if (e.key === 'Escape') { input.value = row.value; input.blur(); }
					});
					input.addEventListener('change', () => commit(input.value));
					swatch.addEventListener('click', () => openColorPop(swatch, (value) => {
						input.value = value;
						swatch.style.background = cssOfColorValue(value);
						commit(value);
					}));
					cell.append(swatch, input);
					div.appendChild(cell);
				} else {
					const input = document.createElement('input');
					input.value = row.value;
					input.disabled = target === '' && row.prop === 'Name';
					if (row.prop === 'Font.Name') { input.setAttribute('list', 'fontFaces'); }
					if (row.prop === 'Font.Size') { input.type = 'number'; input.step = '0.25'; input.min = '1'; }
					input.addEventListener('keydown', (e) => {
						if (e.key === 'Enter') { input.blur(); }
						if (e.key === 'Escape') { input.value = row.value; input.blur(); }
					});
					input.addEventListener('change', () => commit(input.value));
					div.appendChild(input);
				}
				propsBody.appendChild(div);
			}
			if (!document.getElementById('fontFaces')) {
				const list = document.createElement('datalist');
				list.id = 'fontFaces';
				for (const face of FONT_FACES) {
					const o = document.createElement('option');
					o.value = face;
					list.appendChild(o);
				}
				document.body.appendChild(list);
			}
		};

		// The sash on the pane's left edge resizes it; the width persists.
		const propsSash = document.getElementById('propsSash');
		let sashDrag = false;
		propsSash?.addEventListener('pointerdown', (e) => { sashDrag = true; e.preventDefault(); });
		document.addEventListener('pointermove', (e) => {
			if (!sashDrag || !propsPane) { return; }
			const edge = propsPane.getBoundingClientRect().right;
			propsWidth = Math.min(480, Math.max(140, edge - e.clientX));
			propsPane.style.width = propsWidth + 'px';
		});
		document.addEventListener('pointerup', () => {
			if (sashDrag) { sashDrag = false; mergeState({ propsWidth }); }
		});

		// The VBE's double-click: jump to the control's DEFAULT event
		// handler in the code face - Click for buttons and labels, Change
		// for inputs - creating the stub when none exists. The empty face is
		// the form itself: UserForm_Click, whatever the form is named. A
		// dblclick inside a container resolves to the DEEPEST control under
		// the pointer, so a page's empty area belongs to its MultiPage.
		const DEFAULT_EVENTS = {
			Form: 'Click', CommandButton: 'Click', Label: 'Click', TextBox: 'Change',
			ComboBox: 'Change', ListBox: 'Click', CheckBox: 'Click', OptionButton: 'Click',
			ToggleButton: 'Click', Frame: 'Click', MultiPage: 'Change', TabStrip: 'Change',
			ScrollBar: 'Change', SpinButton: 'Change', Image: 'Click',
		};
		document.addEventListener('dblclick', (e) => {
			if (e.target.closest('.toolbar') || e.target.closest('.props')) { return; }
			if (!e.target.closest('.dialog')) { return; }
			const ctl = e.target.closest('.ctl');
			const name = ctl ? ctl.dataset.name : '';
			const kind = PROPS[name] ? PROPS[name].kind : '';
			post({ type: 'openHandler', name, event: DEFAULT_EVENTS[kind] || 'Click' });
		});

		// Every gesture re-renders the whole page, which reset the snap
		// toggles to their defaults - a toggle the user turned OFF came back
		// ON after the next drag. The webview state store survives reloads,
		// so the toggles read from it and every change writes back.
		//
		// The two snaps are rivals - grid pulls to fixed lines, neighbors to
		// living ones, and both at once fight over the same drag - so checking
		// either one clears the other. Both may be off. While grid snap is on,
		// the surfaces paint the lattice it answers to.
		const savedState = vscode?.getState?.() ?? {};
		const gridBox = document.getElementById('snapGrid');
		const neighborsBox = document.getElementById('snapNeighbors');
		if (savedState.grid !== undefined) { gridBox.checked = savedState.grid; }
		if (savedState.neighbors !== undefined) { neighborsBox.checked = savedState.neighbors; }
		if (gridBox.checked && neighborsBox.checked) { neighborsBox.checked = false; }
		const syncGridDots = () => { document.body.classList.toggle('grid-on', gridBox.checked); };
		const saveToggles = () => {
			if (vscode?.setState) {
				vscode.setState({
					...(vscode.getState?.() ?? {}),
					grid: gridBox.checked,
					neighbors: neighborsBox.checked,
					zoom: ZOOM,
				});
			}
		};
		gridBox.addEventListener('change', () => {
			if (gridBox.checked) { neighborsBox.checked = false; }
			syncGridDots();
			saveToggles();
		});
		neighborsBox.addEventListener('change', () => {
			if (neighborsBox.checked) { gridBox.checked = false; }
			syncGridDots();
			saveToggles();
		});
		syncGridDots();
		// Self-check: every control the engine LISTED must be in the DOM
		// with real bounds (hidden pages excepted). One control kept
		// vanishing on one machine while every harness showed it; this names
		// the victim on the face of the canvas instead of leaving the gap
		// silent.
		setTimeout(() => {
			const missing = [];
			for (const name of Object.keys(PROPS)) {
				if (!name) { continue; }
				const el = document.querySelector('[data-name="' + name.replace(/"/g, '\\"') + '"], [data-surface="' + name.replace(/"/g, '\\"') + '"]');
				if (!el) { missing.push(name + ' (absent)'); continue; }
				if (el.closest('[hidden]')) { continue; }
				const r = el.getBoundingClientRect();
				if (r.width <= 0 || r.height <= 0) { missing.push(name + ' (zero-sized)'); }
			}
			for (const name of Object.keys(PROPS)) {
				if (!name) { continue; }
				const el = document.querySelector('[data-name="' + name.replace(/"/g, '\\"') + '"], [data-surface="' + name.replace(/"/g, '\\"') + '"]');
				if (!el || el.closest('[hidden]')) { continue; }
				// A Page matches by its surface and stores an inner offset
				// that says nothing about where it draws; placement talk is
				// for controls only.
				if (!el.hasAttribute('data-name')) { continue; }
				const rows = PROPS[name].rows;
				const expectedLeft = Number(rows.find((r) => r.prop === 'Left')?.value);
				const expectedTop = Number(rows.find((r) => r.prop === 'Top')?.value);
				if (!Number.isFinite(expectedLeft) || !Number.isFinite(expectedTop)) { continue; }
				const surface = el.parentElement.closest('[data-surface]') || el.parentElement;
				const rect = el.getBoundingClientRect();
				const srect = surface.getBoundingClientRect();
				const actualLeft = px2pt(rect.left - srect.left);
				const actualTop = px2pt(rect.top - srect.top);
				if (Math.abs(actualLeft - expectedLeft) > 2 || Math.abs(actualTop - expectedTop) > 2) {
					missing.push(name + ' misplaced (' + expectedLeft + ',' + expectedTop + ')->('
						+ Math.round(actualLeft) + ',' + Math.round(actualTop) + ') pos=' + getComputedStyle(el).position);
					console.error('xlide designer: misplaced control', name, {
						expected: [expectedLeft, expectedTop],
						actual: [actualLeft, actualTop],
						inlineStyle: el.getAttribute('style'),
						computedPosition: getComputedStyle(el).position,
						parent: surface.getAttribute('data-surface'),
					});
				}
			}
			if (missing.length) {
				console.error('xlide designer: controls not rendered:', missing);
				console.error('xlide designer: data-names present:',
					[...document.querySelectorAll('[data-name]')].map((c) => c.getAttribute('data-name')));
				const diag = document.createElement('span');
				diag.className = 'diag';
				diag.textContent = '\u26a0 not rendered: ' + missing.join(', ');
				document.getElementById('toolbar')?.appendChild(diag);
			}
		}, 300);

		const PX_PER_PT = 96 / 72;
		const GRID = 6;
		const SNAP_TOL = 3;
		const FORM_ZOOM = ${formZoom} / 100;
		let ZOOM = FORM_ZOOM;
		const px2pt = (px) => px / (PX_PER_PT * ZOOM);
		const gridOn = () => document.getElementById('snapGrid').checked;
		const neighborsOn = () => document.getElementById('snapNeighbors').checked;

		let selected = document.querySelector('.ctl.selected') || null;
		let armedKind = null;
		const handles = [];

		const geometryOf = (el) => ({
			left: parseFloat(el.style.left), top: parseFloat(el.style.top),
			width: parseFloat(el.style.width), height: parseFloat(el.style.height),
		});
		const setGeometry = (el, g) => {
			el.style.left = g.left + 'pt'; el.style.top = g.top + 'pt';
			if (g.width !== undefined) { el.style.width = g.width + 'pt'; }
			if (g.height !== undefined) { el.style.height = g.height + 'pt'; }
		};

		const clearHandles = () => { handles.splice(0).forEach((h) => h.remove()); };
		const layHandles = () => {
			clearHandles();
			if (!selected) { return; }
			const g = geometryOf(selected);
			const spots = [
				['nw', 0, 0], ['n', 0.5, 0], ['ne', 1, 0], ['e', 1, 0.5],
				['se', 1, 1], ['s', 0.5, 1], ['sw', 0, 1], ['w', 0, 0.5],
			];
			for (const [dir, fx, fy] of spots) {
				const h = document.createElement('div');
				h.className = 'handle';
				h.dataset.dir = dir;
				h.style.left = 'calc(' + (g.left + g.width * fx) + 'pt - 4px)';
				h.style.top = 'calc(' + (g.top + g.height * fy) + 'pt - 4px)';
				h.style.cursor = dir + '-resize';
				selected.parentElement.appendChild(h);
				handles.push(h);
			}
		};
		const select = (el) => {
			document.querySelectorAll('.ctl.selected').forEach((c) => c.classList.remove('selected'));
			selected = el;
			if (el) { el.classList.add('selected'); }
			layHandles();
			renderProps(el ? el.dataset.name : '');
		};
		if (selected) { layHandles(); }
		renderProps(selected ? selected.dataset.name : '');

		// Neighbor snapping: edges and centers of siblings on the same surface.
		const guides = [];
		const clearGuides = () => { guides.splice(0).forEach((g) => g.remove()); };
		const snapLines = (surface, except) => {
			const xs = []; const ys = [];
			surface.querySelectorAll(':scope > .ctl').forEach((c) => {
				if (c === except) { return; }
				const g = geometryOf(c);
				xs.push(g.left, g.left + g.width, g.left + g.width / 2);
				ys.push(g.top, g.top + g.height, g.top + g.height / 2);
			});
			return { xs, ys };
		};
		const snapValue = (value, lines) => {
			for (const line of lines) {
				if (Math.abs(value - line) <= SNAP_TOL) { return line; }
			}
			return null;
		};
		const drawGuide = (surface, axis, pt) => {
			const g = document.createElement('div');
			g.className = 'guide ' + axis;
			if (axis === 'v') { g.style.left = pt + 'pt'; } else { g.style.top = pt + 'pt'; }
			surface.appendChild(g);
			guides.push(g);
		};

		// The FORM itself is selectable: a click on the empty face activates
		// it (the VBE's own gesture), outlining the dialog and laying resize
		// handles on the edges a form can grow by - east, south, southeast.
		// Position is not a form property, so there is no move gesture.
		const dialog = document.querySelector('.dialog');
		const client = document.querySelector('.client');
		const formHandles = [];
		const clearFormHandles = () => { formHandles.splice(0).forEach((h) => h.remove()); };
		const layFormHandles = () => {
			clearFormHandles();
			if (!dialog.classList.contains('form-selected')) { return; }
			for (const dir of ['e', 's', 'se']) {
				const h = document.createElement('div');
				h.className = 'form-handle';
				h.dataset.dir = dir;
				h.style.cursor = dir + '-resize';
				if (dir === 'e') { h.style.right = '-4px'; h.style.top = 'calc(50% - 4px)'; }
				if (dir === 's') { h.style.bottom = '-4px'; h.style.left = 'calc(50% - 4px)'; }
				if (dir === 'se') { h.style.right = '-4px'; h.style.bottom = '-4px'; }
				dialog.appendChild(h);
				formHandles.push(h);
			}
		};
		const selectForm = () => {
			select(null);
			dialog.classList.add('form-selected');
			layFormHandles();
			renderProps('');
		};
		const deselectForm = () => {
			dialog.classList.remove('form-selected');
			clearFormHandles();
		};
		if (dialog.classList.contains('form-selected')) { layFormHandles(); }

		// Zoom scales the DIALOG - handles, guides, and grid dots ride along -
		// and every pointer-to-points conversion divides the factor back out,
		// so the model always sees true points. The select picks presets;
		// Ctrl+wheel glides on an exponential curve between them; the label
		// always tells the truth. The choice survives re-renders.
		const zoomPick = document.getElementById('zoomPick');
		const zoomShow = document.getElementById('zoomShow');
		const paintZoom = () => {
			const ui = parseFloat(zoomPick.value) || 1;
			ZOOM = ui * FORM_ZOOM;
			dialog.style.transform = ZOOM === 1 ? '' : 'scale(' + ZOOM + ')';
			dialog.style.transformOrigin = 'top left';
			if (zoomShow) { zoomShow.textContent = Math.round(ui * 100) + '%'; }
		};
		if (typeof savedState.zoom === 'number' && [...zoomPick.options].some((o) => Number(o.value) === savedState.zoom)) {
			zoomPick.value = String(savedState.zoom);
		}
		paintZoom();
		zoomPick.addEventListener('change', () => { paintZoom(); saveToggles(); });
		document.addEventListener('wheel', (e) => {
			if (!e.ctrlKey) { return; }
			e.preventDefault();
			// The wheel glides the TOTAL scale; paintZoom would snap it back
			// to the select's preset, so the transform applies directly and
			// the label and saved state speak in user-zoom units.
			ZOOM = Math.min(4, Math.max(0.25, ZOOM * Math.exp(-e.deltaY * 0.0015)));
			dialog.style.transform = ZOOM === 1 ? '' : 'scale(' + ZOOM + ')';
			dialog.style.transformOrigin = 'top left';
			if (zoomShow) { zoomShow.textContent = Math.round((ZOOM / FORM_ZOOM) * 100) + '%'; }
			mergeState({ zoom: ZOOM / FORM_ZOOM });
		}, { passive: false });

		let formDrag = null;
		let drag = null;
		document.addEventListener('pointerdown', (e) => {
			if (e.target.closest('.props') || e.target.closest('.split-grip')) { return; }
			const formHandle = e.target.closest('.form-handle');
			if (formHandle) {
				// The client's size lives in the STYLESHEET until the first
				// live resize writes an inline value, so the seed measures
				// the box - a NaN seed made every frame a silent no-op.
				const rect = client.getBoundingClientRect();
				formDrag = {
					dir: formHandle.dataset.dir,
					x: e.clientX, y: e.clientY,
					width: px2pt(rect.width),
					height: px2pt(rect.height),
				};
				client.classList.add('gesture-resize');
				client.style.cursor = formHandle.dataset.dir + '-resize';
				e.preventDefault();
				return;
			}
			const handle = e.target.closest('.handle');
			if (handle && selected) {
				drag = { kind: 'resize', dir: handle.dataset.dir, el: selected,
					start: geometryOf(selected), x: e.clientX, y: e.clientY };
				document.body.dataset.dragging = '1';
				e.preventDefault();
				return;
			}
			if (armedKind) {
				const surface = e.target.closest('[data-surface]');
				if (surface) {
					const rect = surface.getBoundingClientRect();
					let left = px2pt(e.clientX - rect.left);
					let top = px2pt(e.clientY - rect.top);
					if (gridOn()) { left = Math.round(left / GRID) * GRID; top = Math.round(top / GRID) * GRID; }
					post({ type: 'add', container: surface.dataset.surface, controlKind: armedKind,
						left: Math.max(0, left), top: Math.max(0, top) });
					disarm();
				}
				return;
			}
			const ctl = e.target.closest('.ctl[data-name]');
			if (ctl) {
				// The innermost control under the pointer wins; a click on a
				// frame's child selects the child, not the frame.
				select(ctl);
				drag = { kind: 'move', el: ctl, start: geometryOf(ctl), x: e.clientX, y: e.clientY, moved: false };
				document.body.dataset.dragging = '';
				e.preventDefault();
			} else if (e.target.closest('.client')) {
				// The empty face: activate the form itself.
				selectForm();
			} else if (!e.target.closest('.toolbar')) {
				select(null);
				deselectForm();
			}
			if (ctl) { deselectForm(); }
			if (drag) {
				const client = drag.el.closest('.client');
				if (client) {
					client.classList.add(drag.kind === 'move' ? 'gesture-move' : 'gesture-resize');
					if (drag.kind === 'resize') { client.style.cursor = drag.dir + '-resize'; }
				}
			}
		});

		document.addEventListener('pointermove', (e) => {
			if (formDrag) {
				let width = formDrag.width;
				let height = formDrag.height;
				if (formDrag.dir.includes('e')) { width = Math.max(60, formDrag.width + px2pt(e.clientX - formDrag.x)); }
				if (formDrag.dir.includes('s')) { height = Math.max(40, formDrag.height + px2pt(e.clientY - formDrag.y)); }
				if (gridOn()) {
					width = Math.round(width / GRID) * GRID;
					height = Math.round(height / GRID) * GRID;
				}
				formDrag.liveWidth = width;
				formDrag.liveHeight = height;
				client.style.width = width + 'pt';
				client.style.height = height + 'pt';
				dialog.style.width = width + 'pt';
				return;
			}
			if (!drag) { return; }
			const dx = px2pt(e.clientX - drag.x);
			const dy = px2pt(e.clientY - drag.y);
			if (Math.abs(dx) + Math.abs(dy) > 0.5) {
				drag.moved = true;
				document.body.dataset.dragging = '1';
				// A control being carried lifts above its siblings and goes
				// transparent to the pointer, so it is never hidden by what it
				// passes over and never answers its own hit test.
				if (drag.kind === 'move') { drag.el.classList.add('dragging'); }
			}
			clearGuides();
			const surface = drag.el.parentElement.closest('[data-surface]') || drag.el.parentElement;
			const lines = neighborsOn() ? snapLines(drag.el.parentElement, drag.el) : { xs: [], ys: [] };
			if (drag.kind === 'move') {
				let left = drag.start.left + dx;
				let top = drag.start.top + dy;
				if (gridOn()) { left = Math.round(left / GRID) * GRID; top = Math.round(top / GRID) * GRID; }
				const g = drag.start;
				for (const [edge, base] of [[left, 0], [left + g.width, 0], [left + g.width / 2, 0]]) {
					const hit = snapValue(edge, lines.xs);
					if (hit !== null) { left += hit - edge; drawGuide(drag.el.parentElement, 'v', hit); break; }
				}
				for (const [edge] of [[top], [top + g.height], [top + g.height / 2]]) {
					const hit = snapValue(edge, lines.ys);
					if (hit !== null) { top += hit - edge; drawGuide(drag.el.parentElement, 'h', hit); break; }
				}
				setGeometry(drag.el, { left: Math.max(0, left), top: Math.max(0, top) });
				// Carrying over a DIFFERENT surface offers reparenting: the
				// prospective home lights up, and the drop lands there. The
				// carried control is pointer-transparent, so the hit test sees
				// through it and a Frame can never offer itself.
				document.querySelectorAll('.drop-target').forEach((s) => s.classList.remove('drop-target'));
				if (drag.moved) {
					const under = document.elementFromPoint(e.clientX, e.clientY);
					const over = under && under.closest('[data-surface]');
					if (over && over !== drag.el.parentElement) { over.classList.add('drop-target'); }
				}
			} else {
				const g = { ...drag.start };
				const dir = drag.dir;
				if (dir.includes('e')) { g.width = Math.max(6, drag.start.width + dx); }
				if (dir.includes('s')) { g.height = Math.max(6, drag.start.height + dy); }
				if (dir.includes('w')) { g.left = drag.start.left + dx; g.width = Math.max(6, drag.start.width - dx); }
				if (dir.includes('n')) { g.top = drag.start.top + dy; g.height = Math.max(6, drag.start.height - dy); }
				if (gridOn()) {
					if (dir.includes('e')) { g.width = Math.round((g.left + g.width) / GRID) * GRID - g.left; }
					if (dir.includes('s')) { g.height = Math.round((g.top + g.height) / GRID) * GRID - g.top; }
					if (dir.includes('w')) { const r = g.left + g.width; g.left = Math.round(g.left / GRID) * GRID; g.width = r - g.left; }
					if (dir.includes('n')) { const b = g.top + g.height; g.top = Math.round(g.top / GRID) * GRID; g.height = b - g.top; }
				}
				const right = snapValue(g.left + g.width, lines.xs);
				if (right !== null && dir.includes('e')) { g.width = right - g.left; drawGuide(drag.el.parentElement, 'v', right); }
				const bottom = snapValue(g.top + g.height, lines.ys);
				if (bottom !== null && dir.includes('s')) { g.height = bottom - g.top; drawGuide(drag.el.parentElement, 'h', bottom); }
				setGeometry(drag.el, g);
			}
			layHandles();
		});

		document.addEventListener('pointerup', (e) => {
			if (formDrag) {
				// Commit the TRACKED size: the inline style is untouched when
				// the pointer never moved, and a no-move click posts nothing.
				const { width, height, liveWidth, liveHeight } = formDrag;
				formDrag = null;
				client.classList.remove('gesture-resize');
				client.style.cursor = '';
				if (liveWidth !== undefined && (liveWidth !== width || liveHeight !== height)) {
					post({ type: 'formResize', width: liveWidth, height: liveHeight });
				}
				return;
			}
			if (!drag) { return; }
			clearGuides();
			drag.el.classList.remove('dragging');
			const client = drag.el.closest('.client');
			if (client) {
				client.classList.remove('gesture-move', 'gesture-resize');
				client.style.cursor = '';
			}
			document.querySelectorAll('.drop-target').forEach((s) => s.classList.remove('drop-target'));
			if (drag.kind === 'move' && drag.moved) {
				const under = document.elementFromPoint(e.clientX, e.clientY);
				const surf = under && under.closest('[data-surface]');
				if (surf && surf !== drag.el.parentElement) {
					// The drop crossed containers: map the carried position into
					// the new surface and let the engine move the site there.
					const rect = drag.el.getBoundingClientRect();
					const srect = surf.getBoundingClientRect();
					let left = px2pt(rect.left - srect.left);
					let top = px2pt(rect.top - srect.top);
					if (gridOn()) { left = Math.round(left / GRID) * GRID; top = Math.round(top / GRID) * GRID; }
					post({ type: 'reparent', name: drag.el.dataset.name, container: surf.dataset.surface,
						left: Math.max(0, left), top: Math.max(0, top) });
					drag = null;
					delete document.body.dataset.dragging;
					return;
				}
			}
			const g = geometryOf(drag.el);
			const name = drag.el.dataset.name;
			const changed = drag.moved || drag.kind === 'resize';
			const different = g.left !== drag.start.left || g.top !== drag.start.top
				|| g.width !== drag.start.width || g.height !== drag.start.height;
			if (changed && different) {
				post({ type: 'geometry', name, left: g.left, top: g.top,
					...(drag.kind === 'resize' ? { width: g.width, height: g.height } : {}) });
			}
			drag = null;
			delete document.body.dataset.dragging;
		});

		document.addEventListener('keydown', (e) => {
			if (e.target.closest('.props')) { return; }
			if (!selected) { return; }
			const name = selected.dataset.name;
			const step = e.shiftKey ? GRID : 1;
			const g = geometryOf(selected);
			const move = (dl, dt) => {
				setGeometry(selected, { left: Math.max(0, g.left + dl), top: Math.max(0, g.top + dt) });
				layHandles();
				const now = geometryOf(selected);
				post({ type: 'geometry', name, left: now.left, top: now.top });
			};
			if (e.key === 'ArrowLeft') { move(-step, 0); e.preventDefault(); }
			else if (e.key === 'ArrowRight') { move(step, 0); e.preventDefault(); }
			else if (e.key === 'ArrowUp') { move(0, -step); e.preventDefault(); }
			else if (e.key === 'ArrowDown') { move(0, step); e.preventDefault(); }
			else if (e.key === 'Delete') { post({ type: 'remove', name }); e.preventDefault(); }
		});

		const disarm = () => {
			armedKind = null;
			document.body.classList.remove('placing');
			document.querySelectorAll('.tool.armed').forEach((t) => t.classList.remove('armed'));
		};

		// The toolbox works both ways: press-and-DRAG carries a ghost of the
		// control onto a surface and drops it there; a plain CLICK arms the
		// kind for click-to-place. The default size the ghost shows is the
		// size the drop creates.
		const DEFAULT_SIZES = {
			Label: [72, 12], CommandButton: [72, 24], ToggleButton: [72, 24],
			CheckBox: [90, 15], OptionButton: [90, 15], ComboBox: [96, 18],
			TextBox: [96, 18], ListBox: [96, 60], Frame: [120, 90],
			Image: [72, 54], SpinButton: [13, 36], ScrollBar: [13, 90], TabStrip: [150, 90],
		};
		let toolDrag = null;
		const dropPoint = (surface, clientX, clientY) => {
			const rect = surface.getBoundingClientRect();
			let left = px2pt(clientX - rect.left);
			let top = px2pt(clientY - rect.top);
			if (gridOn()) { left = Math.round(left / GRID) * GRID; top = Math.round(top / GRID) * GRID; }
			return { left: Math.max(0, left), top: Math.max(0, top) };
		};
		document.querySelectorAll('.tool').forEach((tool) => {
			tool.addEventListener('pointerdown', (e) => {
				toolDrag = { kind: tool.dataset.kind, x: e.clientX, y: e.clientY, ghost: null, tool };
				e.preventDefault();
			});
		});
		document.addEventListener('pointermove', (e) => {
			if (!toolDrag) { return; }
			if (!toolDrag.ghost) {
				if (Math.abs(e.clientX - toolDrag.x) + Math.abs(e.clientY - toolDrag.y) < 4) { return; }
				toolDrag.ghost = document.createElement('div');
				toolDrag.ghost.className = 'ghost';
				const [w, h] = DEFAULT_SIZES[toolDrag.kind] ?? [72, 24];
				toolDrag.ghost.style.width = w + 'pt';
				toolDrag.ghost.style.height = h + 'pt';
				document.body.classList.add('placing');
			}
			// The ghost lives on the surface under the pointer, at the snapped
			// point the drop would use - so what you see is what you get.
			toolDrag.ghost.remove();
			const under = document.elementFromPoint(e.clientX, e.clientY);
			const surface = under && under.closest('[data-surface]');
			if (!surface) { return; }
			const p = dropPoint(surface, e.clientX, e.clientY);
			toolDrag.ghost.style.left = p.left + 'pt';
			toolDrag.ghost.style.top = p.top + 'pt';
			surface.appendChild(toolDrag.ghost);
		});
		document.addEventListener('pointerup', (e) => {
			if (!toolDrag) { return; }
			const wasDragging = !!toolDrag.ghost;
			const kind = toolDrag.kind;
			const tool = toolDrag.tool;
			if (toolDrag.ghost) { toolDrag.ghost.remove(); }
			toolDrag = null;
			document.body.classList.remove('placing');
			if (wasDragging) {
				const under = document.elementFromPoint(e.clientX, e.clientY);
				const surface = under && under.closest('[data-surface]');
				if (surface) {
					const p = dropPoint(surface, e.clientX, e.clientY);
					post({ type: 'add', container: surface.dataset.surface, controlKind: kind,
						left: p.left, top: p.top });
				}
				return;
			}
			// No movement: a plain click, which toggles click-to-place.
			if (armedKind === kind) { disarm(); return; }
			disarm();
			armedKind = kind;
			tool.classList.add('armed');
			document.body.classList.add('placing');
		});
		document.addEventListener('keyup', (e) => {
			if (e.target.closest('.props')) { return; }
			if (e.key === 'Escape') { disarm(); select(null); deselectForm(); }
		});
		// F5 posted from INSIDE the canvas: key forwarding from webviews to
		// workbench keybindings has not proven reliable for it.
		document.addEventListener('keydown', (e) => {
			if (e.key === 'F5') { e.preventDefault(); post({ type: 'launchHost' }); }
		});
	})();
	</script>` : ''}
</body>
</html>`;
}
