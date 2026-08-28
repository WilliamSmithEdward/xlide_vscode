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
import { decodeArrayStrings, formatOleColor } from './markup';
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
	return parts.join('');
}

function colorCss(record: ParsedRecord | undefined): string {
	if (!record) { return ''; }
	const parts: string[] = [];
	const back = record.values.get('BackColor');
	if (back !== undefined && recordHas(record, 'BackColor')) { parts.push(`background:${cssColor(back)};`); }
	const fore = record.values.get('ForeColor');
	if (fore !== undefined && recordHas(record, 'ForeColor')) { parts.push(`color:${cssColor(fore)};`); }
	return parts.join('');
}

/** Renders one form package to the inner HTML of its surface. */
function renderSurface(pkg: FormPackage, idPrefix: string, selected?: string): string {
	const parts: string[] = [];
	pkg.entries.forEach((entry, index) => {
		const site = entry.site;
		const record = entry.kind === 'record' ? entry.record : undefined;
		const kind = controlKindOfSite(site, record);
		const name = siteName(site);
		const inner = siteIsContainer(site) ? pkg.containers.get(siteId(site)) : undefined;
		const box = siteBox(site, record, inner);
		const style = `left:${pts(box.left)}pt;top:${pts(box.top)}pt;width:${pts(box.width)}pt;height:${pts(box.height)}pt;`
			+ fontCss(record) + colorCss(record);
		const sel = selected !== undefined && selected.toLowerCase() === name.toLowerCase() ? ' selected' : '';
		const dn = `data-name="${esc(name)}"`;
		const caption = record?.strings.get('Caption')?.text
			?? inner?.form.record.strings.get('Caption')?.text
			?? '';
		const value = record?.strings.get('Value')?.text ?? '';
		const childId = `${idPrefix}-${index}`;

		switch (kind) {
			case 'Label':
				parts.push(`<div class="ctl label${sel}" ${dn} style="${style}" title="${esc(name)}">${esc(caption)}</div>`);
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
				parts.push(`<div class="ctl button${kind === 'ToggleButton' && value === '1' ? ' pressed' : ''}${sel}" ${dn} style="${style}" title="${esc(name)}">${esc(caption)}</div>`);
				break;
			case 'Image':
				parts.push(`<div class="ctl image${sel}" ${dn} style="${style}" title="${esc(name)}"><span>${esc(name)}</span></div>`);
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
				parts.push(`<div class="ctl frame${sel}" ${dn} style="${style}" title="${esc(name)}"><span class="legend">${esc(caption)}</span><div class="surface" data-surface="${esc(name)}">${inner ? renderSurface(inner, childId, selected) : ''}</div></div>`);
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
					return `<div class="page" id="${childId}-p${i}" data-surface="${esc(siteName(pageSite))}"${i === 0 ? '' : ' hidden'}>${pagePkg ? renderSurface(pagePkg, `${childId}-p${i}`, selected) : ''}</div>`;
				}).join('');
				parts.push(`<div class="ctl multipage${sel}" ${dn} style="${style}" title="${esc(name)}"><div class="tabs">${headers}</div><div class="pagearea">${pages}</div></div>`);
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
	const stdFont = pkg.form.fontRaw ? parseStdFont(pkg.form.fontRaw) : undefined;
	const formFont = stdFont
		? `font-family:'${esc(stdFont.face)}',Tahoma,sans-serif;font-size:${stdFont.heightTenThousandthsPt / 10000}pt;`
		: "font-family:Tahoma,sans-serif;font-size:8.25pt;";

	const toolbox = ['Label', 'TextBox', 'ComboBox', 'ListBox', 'CheckBox', 'OptionButton',
		'ToggleButton', 'Frame', 'CommandButton', 'TabStrip', 'ScrollBar', 'SpinButton', 'Image'];
	const interactive = options.interactive !== false;
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
	.toolbar .tool.armed { background: #0e639c; color: #fff; border-color: #1177bb; }
	.toolbar label { display: flex; gap: 4px; align-items: center; margin-left: 8px; }
	.stage { padding: 24px; }
	.dialog { width: ${width}pt; box-shadow: 2px 2px 8px rgba(0,0,0,.5); }
	.titlebar { background: #99b4d1; color: #000; padding: 2px 6px;
		font: bold 9pt Tahoma, sans-serif; display: flex; justify-content: space-between; }
	.client { position: relative; width: ${width}pt; height: ${height}pt;
		background: ${formBack}; overflow: hidden; ${formFont} }
	.ctl { position: absolute; box-sizing: border-box; overflow: hidden; white-space: nowrap; }
	.ctl.selected { outline: 1px dashed #0e639c; outline-offset: 1px; }
	.label { display: flex; align-items: flex-start; }
	.edit { background: #fff; color: #000; border: 1px solid #7a7a7a; padding: 1px 2px; }
	.combo { display: flex; justify-content: space-between; align-items: center; }
	.combo .drop { border-left: 1px solid #a0a0a0; background: #f0f0f0; align-self: stretch;
		display: flex; align-items: center; padding: 0 2px; }
	.opt { display: flex; align-items: center; gap: 4px; }
	.opt .box { width: 9pt; height: 9pt; background: #fff; border: 1px solid #7a7a7a; flex: none; }
	.opt .box.on::after { content: '\\2713'; display: block; text-align: center; line-height: 9pt; }
	.opt .radio { width: 9pt; height: 9pt; background: #fff; border: 1px solid #7a7a7a;
		border-radius: 50%; flex: none; }
	.opt .radio.on { background: radial-gradient(circle at center, #000 35%, #fff 40%); }
	.button { background: #f0f0f0; border: 1px solid #707070;
		box-shadow: inset 1px 1px 0 #fff, inset -1px -1px 0 #a0a0a0;
		display: flex; align-items: center; justify-content: center; }
	.button.pressed { border-style: inset; }
	.image, .foreign { border: 1px solid #a0a0a0;
		background: repeating-linear-gradient(45deg, #ddd 0 6px, #eee 6px 12px);
		display: flex; align-items: center; justify-content: center; color: #555; font-size: 7pt; }
	.spin { display: flex; flex-direction: column; }
	.spin span { flex: 1; background: #f0f0f0; border: 1px solid #a0a0a0;
		display: flex; align-items: center; justify-content: center; font-size: 6pt; }
	.scroll { background: #d4d0c8; border: 1px solid #a0a0a0; position: relative; }
	.scroll::after { content: ''; position: absolute; left: 1px; right: 1px; top: 15%; height: 30%;
		background: #f0f0f0; border: 1px solid #808080; }
	.frame { border: 1px solid #a0a0a0; box-shadow: inset 0 0 0 1px #fff;
		overflow: visible; }
	.frame .legend { position: absolute; top: 0; left: 6pt; transform: translateY(-55%);
		background: ${formBack}; padding: 0 3px; line-height: 1.1; }
	.frame .surface { position: absolute; inset: 0; overflow: hidden; }
	.tabs { display: flex; gap: 1px; padding: 0 2px; height: 14pt; align-items: flex-end;
		position: relative; z-index: 1; }
	.tab { background: #f0f0f0; border: 1px solid #a0a0a0; border-bottom: none;
		padding: 0 8px 1px; border-radius: 3px 3px 0 0; cursor: pointer; }
	.tab.active { background: #fff; position: relative; top: 1px; }
	.tabstrip, .multipage { background: #f0f0f0; overflow: visible; }
	.tabbody, .pagearea { position: absolute; inset: 14pt 0 0; border: 1px solid #a0a0a0;
		box-shadow: inset 0 0 0 1px #fff; background: #f0f0f0; overflow: hidden; }
	.page { position: absolute; inset: 0; }
	.handle { position: absolute; width: 6px; height: 6px; background: #fff;
		border: 1px solid #0e639c; z-index: 5; }
	.guide { position: absolute; background: #e51400; z-index: 4; pointer-events: none; }
	.guide.v { width: 1px; top: 0; bottom: 0; }
	.guide.h { height: 1px; left: 0; right: 0; }
	.placing, .placing .ctl { cursor: crosshair !important; }
</style>
</head>
<body>
${interactive ? `	<div class="toolbar" id="toolbar">
		<span>Add:</span>
		${toolbox.map((k) => `<button class="tool" data-kind="${k}">${k}</button>`).join('')}
		<label><input type="checkbox" id="snapGrid" checked>Grid 6pt</label>
		<label><input type="checkbox" id="snapNeighbors" checked>Snap to neighbors</label>
	</div>` : ''}
	<div class="stage">
	<div class="dialog">
		<div class="titlebar"><span>${esc(caption)}</span><span>&#10005;</span></div>
		<div class="client" data-surface="">
${renderSurface(pkg, 'c', options.selected)}
		</div>
	</div>
	</div>
	<script>
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
		const PX_PER_PT = 96 / 72;
		const GRID = 6;
		const SNAP_TOL = 3;
		const px2pt = (px) => px / PX_PER_PT;
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
		};
		if (selected) { layHandles(); }

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

		let drag = null;
		document.addEventListener('pointerdown', (e) => {
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
			} else if (!e.target.closest('.toolbar')) {
				select(null);
			}
		});

		document.addEventListener('pointermove', (e) => {
			if (!drag) { return; }
			const dx = px2pt(e.clientX - drag.x);
			const dy = px2pt(e.clientY - drag.y);
			if (Math.abs(dx) + Math.abs(dy) > 0.5) { drag.moved = true; document.body.dataset.dragging = '1'; }
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

		document.addEventListener('pointerup', () => {
			if (!drag) { return; }
			clearGuides();
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
		document.querySelectorAll('.tool').forEach((tool) => {
			tool.addEventListener('click', () => {
				if (armedKind === tool.dataset.kind) { disarm(); return; }
				disarm();
				armedKind = tool.dataset.kind;
				tool.classList.add('armed');
				document.body.classList.add('placing');
			});
		});
		document.addEventListener('keyup', (e) => { if (e.key === 'Escape') { disarm(); } });
	})();
	</script>` : ''}
</body>
</html>`;
}
