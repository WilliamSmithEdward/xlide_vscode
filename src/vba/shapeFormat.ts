// How a shape looks: fill, outline, rotation and the font of its text.
//
// Excel, PowerPoint and Word write all of it in the same DrawingML, inside
// each host's own properties element (xdr:spPr, p:spPr, wps:spPr), and each
// was measured by having the application set one property and save:
//
//   - fill is the spPr's a:solidFill, with transparency as an a:alpha child
//     of the color (a:alpha 75000 is 25% transparent); a:noFill is none;
//   - outline is the spPr's a:ln: width `w` in EMU, the same fill children,
//     and an a:prstDash for its dashes;
//   - rotation is `rot` on the a:xfrm, in 60,000ths of a degree;
//   - the font is the a:rPr of every run: `sz` in hundredths of a point,
//     `b`, `i`, `u="sng"`, a fill child for its color and a:latin.
//
// What a shape does not set comes from its style (xdr:style, p:style,
// wps:style): a reference into the theme's fill and line styles, colored by
// a theme color. That is how a new AutoShape is "Accent 1", and it is read
// here so the editor shows the color the application shows. A shape's own
// setting always wins, so every write here sets the shape's own properties
// and leaves the style alone.
//
// Word's text is WordprocessingML, not DrawingML, so its font is a w:rPr on
// every run and paragraph mark instead; that is here too, since the editor
// treats a font as one thing whichever host it is in.

import { attr, children, emuToPoints, encodeXml, findElement, insertInOrder, pointsToEmu, splice, withAttr, type Span } from './ooxml';

/** A DrawingML preset dash, as a:prstDash names it. */
export type ShapeDash =
	| 'solid' | 'dot' | 'dash' | 'lgDash' | 'dashDot' | 'lgDashDot' | 'lgDashDotDot'
	| 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot';

export const SHAPE_DASHES: readonly ShapeDash[] = [
	'solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot',
	'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot',
];

/**
 * A shape's fill. `other` is a gradient, picture or pattern, which is shown
 * but not edited. `automatic` means the shape sets nothing itself and the
 * value comes from its style and the theme.
 */
export interface ShapeFill {
	type: 'none' | 'solid' | 'other';
	/** #RRGGBB. */
	color?: string;
	/** 0 to 100 percent. */
	transparency?: number;
	/** The theme color the color comes from, such as accent1. */
	themeColor?: string;
	automatic?: boolean;
}

export interface ShapeLine {
	type: 'none' | 'solid' | 'other';
	color?: string;
	themeColor?: string;
	/** Points. */
	weight?: number;
	dash?: ShapeDash;
	automatic?: boolean;
}

export interface ShapeFont {
	name?: string;
	/** Points. */
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	/** #RRGGBB. */
	color?: string;
}

export type FillEdit = { type: 'none' } | { type: 'solid'; color: string; transparency?: number };
export type LineEdit = { type: 'none' } | { type: 'solid'; color?: string; weight?: number; dash?: ShapeDash };
/** A font change: what is given is set, and a color of '' goes back to the style's. */
export type FontEdit = Partial<ShapeFont>;

export class ShapeFormatError extends Error {}

// -------------------------------------------------------------------- colors

/** #RRGGBB from six hex digits, or undefined when it is not a color. */
export function normalizeColor(value: string | undefined): string | undefined {
	const m = /^#?([0-9a-fA-F]{6})$/.exec((value ?? '').trim());
	return m ? `#${m[1].toUpperCase()}` : undefined;
}

function requireColor(value: string, what: string): string {
	const color = normalizeColor(value);
	if (!color) {
		throw new ShapeFormatError(`'${value}' is not a ${what} color; give one as #RRGGBB, such as #FF0000.`);
	}
	return color;
}

/** What a shape's theme holds that its appearance reads. */
export interface ShapeTheme {
	/** RRGGBB by scheme name: dk1, lt1, dk2, lt2, accent1 to accent6, hlink, folHlink. */
	colors: Record<string, string>;
	majorFont?: string;
	minorFont?: string;
	/** Line widths in EMU for style references 1, 2 and 3. */
	lineWidths: number[];
	/** Whether style fill 1, 2 and 3 is a solid color. */
	solidFills: boolean[];
}

/** How the style names map to the scheme, as every host's default writes it. */
const COLOR_MAP: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

/** A theme part read, or an empty theme when the package has none. */
export function readTheme(xml: string | undefined): ShapeTheme {
	const theme: ShapeTheme = { colors: {}, lineWidths: [], solidFills: [] };
	if (!xml) { return theme; }
	const scheme = findElement(xml, 'a:clrScheme');
	if (scheme) {
		for (const entry of children(xml, scheme.openEnd, scheme.end)) {
			const inner = xml.slice(entry.openEnd, entry.end);
			const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner)?.[1];
			const sys = /<a:sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(inner)?.[1];
			const value = srgb ?? sys;
			if (value) { theme.colors[entry.name.replace(/^a:/, '')] = value.toUpperCase(); }
		}
	}
	const major = /<a:majorFont>\s*<a:latin\b[^>]*\btypeface="([^"]*)"/.exec(xml)?.[1];
	const minor = /<a:minorFont>\s*<a:latin\b[^>]*\btypeface="([^"]*)"/.exec(xml)?.[1];
	if (major) { theme.majorFont = major; }
	if (minor) { theme.minorFont = minor; }
	const lines = findElement(xml, 'a:lnStyleLst');
	if (lines) {
		theme.lineWidths = children(xml, lines.openEnd, lines.end)
			.map((ln) => Number(attr(xml.slice(ln.start, ln.openEnd), 'w') ?? 0));
	}
	const fills = findElement(xml, 'a:fillStyleLst');
	if (fills) {
		theme.solidFills = children(xml, fills.openEnd, fills.end).map((fill) => fill.name === 'a:solidFill');
	}
	return theme;
}

interface Rgb { r: number; g: number; b: number }

function hexToRgb(hex: string): Rgb {
	const n = parseInt(hex.replace('#', ''), 16);
	return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function rgbToHex({ r, g, b }: Rgb): string {
	const part = (v: number): string => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
	return `#${part(r)}${part(g)}${part(b)}`.toUpperCase();
}

function rgbToHsl({ r, g, b }: Rgb): { h: number; s: number; l: number } {
	const [rr, gg, bb] = [r / 255, g / 255, b / 255];
	const max = Math.max(rr, gg, bb);
	const min = Math.min(rr, gg, bb);
	const l = (max + min) / 2;
	if (max === min) { return { h: 0, s: 0, l }; }
	const d = max - min;
	const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
	const h = max === rr ? (gg - bb) / d + (gg < bb ? 6 : 0) : max === gg ? (bb - rr) / d + 2 : (rr - gg) / d + 4;
	return { h: h / 6, s, l };
}

function hslToRgb({ h, s, l }: { h: number; s: number; l: number }): Rgb {
	if (s === 0) { return { r: l * 255, g: l * 255, b: l * 255 }; }
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	const hue = (t: number): number => {
		const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
		if (x < 1 / 6) { return p + (q - p) * 6 * x; }
		if (x < 1 / 2) { return q; }
		if (x < 2 / 3) { return p + (q - p) * (2 / 3 - x) * 6; }
		return p;
	};
	return { r: hue(h + 1 / 3) * 255, g: hue(h) * 255, b: hue(h - 1 / 3) * 255 };
}

const toLinear = (c: number): number => {
	const v = c / 255;
	return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
};
const fromLinear = (v: number): number => 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055);

/**
 * A color element's transforms applied: lumMod and lumOff in HSL, shade and
 * tint in linear RGB, as DrawingML defines them. The result is for showing;
 * a color the user did not change is never written back from it.
 */
function transformed(hex: string, colorXml: string): { color: string; alpha: number } {
	let rgb = hexToRgb(hex);
	let alpha = 1;
	for (const [, name, value] of colorXml.matchAll(/<a:(lumMod|lumOff|shade|tint|alpha)\b[^>]*\bval="(-?\d+)"/g)) {
		const v = Number(value) / 100000;
		if (name === 'alpha') {
			alpha = v;
		} else if (name === 'lumMod' || name === 'lumOff') {
			const hsl = rgbToHsl(rgb);
			hsl.l = Math.max(0, Math.min(1, name === 'lumMod' ? hsl.l * v : hsl.l + v));
			rgb = hslToRgb(hsl);
		} else if (name === 'shade') {
			rgb = { r: fromLinear(toLinear(rgb.r) * v), g: fromLinear(toLinear(rgb.g) * v), b: fromLinear(toLinear(rgb.b) * v) };
		} else {
			const tint = (c: number): number => fromLinear(1 - (1 - toLinear(c)) * v);
			rgb = { r: tint(rgb.r), g: tint(rgb.g), b: tint(rgb.b) };
		}
	}
	return { color: rgbToHex(rgb), alpha };
}

/** The handful of preset color names a shape's markup can carry. */
const PRESET_COLORS: Record<string, string> = {
	black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
	gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', navy: '000080', purple: '800080',
	teal: '008080', olive: '808000', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF', orange: 'FFA500',
};

interface ResolvedColor { color?: string; alpha: number; themeColor?: string }

/**
 * The color a fill, line or style reference holds: its first color element,
 * with a theme color looked up and its transforms applied. `placeholder` is
 * what phClr means: the color of the style reference being expanded.
 */
function resolveColor(xml: string, holder: Span, theme: ShapeTheme | undefined, placeholder?: ResolvedColor): ResolvedColor {
	const element = children(xml, holder.openEnd, holder.end)
		.find((child) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(child.name));
	if (!element) { return { alpha: 1 }; }
	const colorXml = xml.slice(element.start, element.end);
	const tag = xml.slice(element.start, element.openEnd);
	let base: string | undefined;
	let themeColor: string | undefined;
	switch (element.name) {
		case 'a:srgbClr': base = attr(tag, 'val'); break;
		case 'a:sysClr': base = attr(tag, 'lastClr'); break;
		case 'a:prstClr': base = PRESET_COLORS[(attr(tag, 'val') ?? '').toLowerCase()]; break;
		case 'a:scrgbClr': {
			const channel = (name: string): number => fromLinear(Number(attr(tag, name) ?? 0) / 100000);
			base = rgbToHex({ r: channel('r'), g: channel('g'), b: channel('b') }).slice(1);
			break;
		}
		case 'a:hslClr': {
			const hsl = { h: Number(attr(tag, 'hue') ?? 0) / 21600000, s: Number(attr(tag, 'sat') ?? 0) / 100000, l: Number(attr(tag, 'lum') ?? 0) / 100000 };
			base = rgbToHex(hslToRgb(hsl)).slice(1);
			break;
		}
		default: {
			const name = attr(tag, 'val') ?? '';
			if (name === 'phClr') {
				if (!placeholder?.color) { return { alpha: 1 }; }
				base = placeholder.color.slice(1);
				themeColor = placeholder.themeColor;
			} else {
				themeColor = COLOR_MAP[name] ?? name;
				base = theme?.colors[themeColor];
			}
		}
	}
	if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) { return { alpha: 1, ...(themeColor ? { themeColor } : {}) }; }
	const { color, alpha } = transformed(base, colorXml);
	return { color, alpha, ...(themeColor ? { themeColor } : {}) };
}

// ---------------------------------------------------------------- reading

const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

/** The style reference of a shape's style element: fillRef, lnRef or fontRef. */
function styleRef(xml: string, style: Span | undefined, name: string): Span | undefined {
	return style ? findElement(xml, name, style.openEnd, style.end) : undefined;
}

function transparencyOf(alpha: number): number | undefined {
	const percent = Math.round((1 - alpha) * 100);
	return percent > 0 ? percent : undefined;
}

/** A resolved color as a fill or line carries it. */
function colorFields(resolved: ResolvedColor): { color?: string; themeColor?: string } {
	return {
		...(resolved.color ? { color: resolved.color } : {}),
		...(resolved.themeColor ? { themeColor: resolved.themeColor } : {}),
	};
}

/** A solid fill from a resolved color, its alpha read as transparency. */
function solidFill(resolved: ResolvedColor): ShapeFill {
	const transparency = transparencyOf(resolved.alpha);
	return { type: 'solid', ...colorFields(resolved), ...(transparency !== undefined ? { transparency } : {}) };
}

/**
 * A shape's fill: its own when its properties set one, else what its style
 * takes from the theme. Undefined when neither says anything, as for a
 * picture, whose image is not a fill.
 */
export function readFill(xml: string, spPr: Span | undefined, style: Span | undefined, theme: ShapeTheme | undefined): ShapeFill | undefined {
	const own = spPr ? children(xml, spPr.openEnd, spPr.end).find((child) => FILLS.includes(child.name)) : undefined;
	if (own) {
		if (own.name === 'a:noFill') { return { type: 'none' }; }
		return own.name === 'a:solidFill' ? solidFill(resolveColor(xml, own, theme)) : { type: 'other' };
	}
	const ref = styleRef(xml, style, 'a:fillRef');
	if (!ref) { return undefined; }
	const idx = Number(attr(xml.slice(ref.start, ref.openEnd), 'idx') ?? 0);
	if (idx === 0) { return { type: 'none', automatic: true }; }
	// 1 to 3 name the theme's fill styles; 1001 and up its background fills.
	const solid = idx > 1000 ? true : theme?.solidFills[idx - 1] ?? true;
	return solid ? { ...solidFill(resolveColor(xml, ref, theme)), automatic: true } : { type: 'other', automatic: true };
}

/**
 * A shape's outline: what its style's line reference gives, with whatever
 * its own a:ln sets laid over it - the width can be the shape's own while
 * the color still comes from the style.
 */
export function readLine(xml: string, spPr: Span | undefined, style: Span | undefined, theme: ShapeTheme | undefined): ShapeLine | undefined {
	const ln = spPr ? children(xml, spPr.openEnd, spPr.end).find((child) => child.name === 'a:ln') : undefined;
	const ref = styleRef(xml, style, 'a:lnRef');
	if (!ln && !ref) { return undefined; }
	const refIdx = ref ? Number(attr(xml.slice(ref.start, ref.openEnd), 'idx') ?? 0) : 0;
	const styleWeight = refIdx > 0 && theme?.lineWidths[refIdx - 1] ? emuToPoints(theme.lineWidths[refIdx - 1]) : undefined;
	let line: ShapeLine = refIdx > 0 && ref
		? { type: 'solid', ...colorFields(resolveColor(xml, ref, theme)), ...(styleWeight !== undefined ? { weight: styleWeight } : {}), automatic: true }
		: { type: 'none', automatic: true };
	if (!ln) { return line; }
	const fill = children(xml, ln.openEnd, ln.end).find((child) => FILLS.includes(child.name));
	if (fill) {
		if (fill.name === 'a:noFill') { return { type: 'none' }; }
		line = fill.name === 'a:solidFill'
			? { type: 'solid', ...colorFields(resolveColor(xml, fill, theme)), ...(styleWeight !== undefined ? { weight: styleWeight } : {}) }
			: { type: 'other' };
	}
	if (line.type === 'none') { return line; }
	const w = attr(xml.slice(ln.start, ln.openEnd), 'w');
	if (w !== undefined) { line.weight = emuToPoints(Number(w)); }
	const dash = /<a:prstDash\b[^>]*\bval="([^"]*)"/.exec(xml.slice(ln.openEnd, ln.end))?.[1];
	if (dash && (SHAPE_DASHES as readonly string[]).includes(dash)) { line.dash = dash as ShapeDash; }
	return line;
}

/** Degrees clockwise from an a:xfrm's rot, or undefined when it is not rotated. */
export function readRotation(xfrmTag: string | undefined): number | undefined {
	const rot = xfrmTag ? Number(attr(xfrmTag, 'rot') ?? 0) : 0;
	if (!rot) { return undefined; }
	const degrees = Math.round((rot / 60000) * 100) / 100;
	return ((degrees % 360) + 360) % 360 || undefined;
}

/**
 * A list style that text inherits its format from: a shape's own
 * a:lstStyle, a layout or master placeholder's, a master's p:titleStyle,
 * p:bodyStyle or p:otherStyle, or a presentation's p:defaultTextStyle. Each
 * holds a:lvl1pPr to a:lvl9pPr, and the a:defRPr in a level is the format a
 * paragraph at that level starts from.
 */
export interface TextStyleSource {
	xml: string;
	element: Span;
}

/** What one run-properties element sets: an a:rPr, a:endParaRPr or a:defRPr. */
function runFontOf(rPr: string, theme: ShapeTheme | undefined): ShapeFont {
	const font: ShapeFont = {};
	const openEnd = rPr.indexOf('>') + 1;
	const tag = rPr.slice(0, openEnd);
	const sz = attr(tag, 'sz');
	if (sz) { font.size = Number(sz) / 100; }
	const b = attr(tag, 'b');
	if (b !== undefined) { font.bold = b === '1' || b === 'true'; }
	const i = attr(tag, 'i');
	if (i !== undefined) { font.italic = i === '1' || i === 'true'; }
	const u = attr(tag, 'u');
	if (u !== undefined) { font.underline = u !== 'none'; }
	// Direct children only: an a:ln (the text's own outline) comes first
	// and holds a fill of its own.
	const own = children(rPr, openEnd, rPr.length);
	const latin = own.find((child) => child.name === 'a:latin');
	const typeface = latin ? attr(rPr.slice(latin.start, latin.openEnd), 'typeface') : undefined;
	if (typeface) { font.name = themeFont(typeface, theme) ?? typeface; }
	const fill = own.find((child) => child.name === 'a:solidFill');
	if (fill) {
		const color = resolveColor(rPr, fill, theme).color;
		if (color) { font.color = color; }
	}
	return font;
}

/** The a:defRPr a list style gives a paragraph level (0 is the first). */
function levelFont(source: TextStyleSource, level: number, theme: ShapeTheme | undefined): ShapeFont {
	const lvl = findElement(source.xml, `a:lvl${level + 1}pPr`, source.element.openEnd, source.element.end);
	const defRPr = lvl ? findElement(source.xml, 'a:defRPr', lvl.openEnd, lvl.end) : undefined;
	return defRPr ? runFontOf(source.xml.slice(defRPr.start, defRPr.end), theme) : {};
}

/** `font` with every property it lacks taken from `from`. */
function inheritFont(font: ShapeFont, from: ShapeFont): void {
	for (const key of Object.keys(from) as Array<keyof ShapeFont>) {
		if (font[key] === undefined) { (font as Record<string, unknown>)[key] = from[key]; }
	}
}

/**
 * The font of a DrawingML text body, from its first run - which is what the
 * application's font box shows for a shape whose text is all one format.
 * What the run does not set comes, in order, from the body's own list
 * style, the shape style's font reference (its theme font and color), and
 * `inherited`, nearest first: PowerPoint measured it so, with a shape's
 * style color outranking the presentation's default text color.
 */
export function readDrawingFont(
	xml: string,
	txBody: Span | undefined,
	style: Span | undefined,
	theme: ShapeTheme | undefined,
	inherited: readonly TextStyleSource[] = [],
): ShapeFont | undefined {
	if (!txBody) { return undefined; }
	const inner = xml.slice(txBody.openEnd, txBody.end);
	const paragraphs = [...inner.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((m) => m[0]);
	const paragraph = paragraphs.find((p) => /<a:r>/.test(p)) ?? paragraphs[0] ?? '';
	const level = Number(/<a:pPr\b[^>]*\blvl="(\d)"/.exec(paragraph)?.[1] ?? 0);
	const rPr = /<a:r>\s*(<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>)/.exec(paragraph)?.[1]
		?? /<a:endParaRPr\b[^>]*\/>|<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/.exec(paragraph)?.[0];
	const font: ShapeFont = rPr ? runFontOf(rPr, theme) : {};
	const own = findElement(xml, 'a:lstStyle', txBody.openEnd, txBody.end);
	if (own) { inheritFont(font, levelFont({ xml, element: own }, level, theme)); }
	const ref = styleRef(xml, style, 'a:fontRef');
	if (ref) {
		const idx = attr(xml.slice(ref.start, ref.openEnd), 'idx');
		const name = idx === 'major' ? theme?.majorFont : idx === 'minor' ? theme?.minorFont : undefined;
		const color = resolveColor(xml, ref, theme).color;
		inheritFont(font, { ...(name ? { name } : {}), ...(color ? { color } : {}) });
	}
	for (const source of inherited) {
		inheritFont(font, levelFont(source, level, theme));
	}
	return Object.keys(font).length > 0 ? font : undefined;
}

function themeFont(typeface: string, theme: ShapeTheme | undefined): string | undefined {
	if (typeface === '+mn-lt' || typeface === '+mn-ea' || typeface === '+mn-cs') { return theme?.minorFont; }
	if (typeface === '+mj-lt' || typeface === '+mj-ea' || typeface === '+mj-cs') { return theme?.majorFont; }
	return undefined;
}

/** What a Word document's styles part gives text that sets nothing itself. */
export interface WordTextStyles {
	/** Each style's own w:rPr, and the style it is based on, by style id. */
	styles: Map<string, { rPr?: string; basedOn?: string }>;
	/** The style of a paragraph that names none: Normal, as a rule. */
	defaultParagraph?: string;
	/** The document defaults' w:rPr, which every style starts from. */
	defaults?: string;
}

/** A Word styles part read for text formatting, or nothing when there is none. */
export function readWordStyles(xml: string | undefined): WordTextStyles {
	const out: WordTextStyles = { styles: new Map() };
	if (!xml) { return out; }
	const defaults = /<w:rPrDefault>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(xml)?.[1];
	if (defaults) { out.defaults = defaults; }
	for (const [style] of xml.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)) {
		const tag = style.slice(0, style.indexOf('>') + 1);
		const id = attr(tag, 'w:styleId');
		if (!id) { continue; }
		// The style's own w:rPr: a table style nests more inside its parts.
		const own = children(style, tag.length, style.length).find((child) => child.name === 'w:rPr');
		const basedOn = /<w:basedOn\b[^>]*\bw:val="([^"]*)"/.exec(style)?.[1];
		out.styles.set(id, {
			...(own ? { rPr: style.slice(own.start, own.end) } : {}),
			...(basedOn ? { basedOn } : {}),
		});
		if (attr(tag, 'w:type') === 'paragraph' && attr(tag, 'w:default') === '1') { out.defaultParagraph = id; }
	}
	return out;
}

/** What one w:rPr sets. A color of "auto" is left unset: it is the absence of one. */
function wordRunFontOf(rPr: string | undefined, theme: ShapeTheme | undefined): ShapeFont {
	const font: ShapeFont = {};
	if (!rPr) { return font; }
	const fonts = /<w:rFonts\b[^>]*>/.exec(rPr)?.[0];
	const ascii = fonts ? attr(fonts, 'w:ascii') : undefined;
	if (ascii) { font.name = ascii; } else if (fonts && attr(fonts, 'w:asciiTheme')) {
		const themeFace = attr(fonts, 'w:asciiTheme')!.toLowerCase().startsWith('major') ? theme?.majorFont : theme?.minorFont;
		if (themeFace) { font.name = themeFace; }
	}
	const sz = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(rPr)?.[1];
	if (sz) { font.size = Number(sz) / 2; }
	const toggle = (name: string): boolean | undefined => {
		const tag = new RegExp(`<w:${name}(?=[\\s/>])[^>]*>`).exec(rPr)?.[0];
		if (!tag) { return undefined; }
		const val = attr(tag, 'w:val');
		return val === undefined || !['0', 'false', 'off'].includes(val);
	};
	const bold = toggle('b');
	if (bold !== undefined) { font.bold = bold; }
	const italic = toggle('i');
	if (italic !== undefined) { font.italic = italic; }
	const u = /<w:u\b[^>]*\bw:val="([^"]*)"/.exec(rPr)?.[1];
	if (u !== undefined) { font.underline = u !== 'none'; }
	const color = normalizeColor(/<w:color\b[^>]*\bw:val="([^"]*)"/.exec(rPr)?.[1]);
	if (color) { font.color = color; }
	return font;
}

/** `font` with what a style and every style it is based on set filled in. */
function inheritWordStyle(font: ShapeFont, styles: WordTextStyles, id: string | undefined, theme: ShapeTheme | undefined): void {
	const seen = new Set<string>();
	for (let at = id; at && !seen.has(at); at = styles.styles.get(at)?.basedOn) {
		seen.add(at);
		inheritFont(font, wordRunFontOf(styles.styles.get(at)?.rPr, theme));
	}
}

/**
 * The font of a Word text box, from its first run - or, in an empty one,
 * its first paragraph mark. What the run does not set comes from its
 * character style, its paragraph's style, the document defaults, and last
 * the shape style's font reference. Word showed that order: a Normal style
 * that sets a font and a color wins over the shape's white theme text, and
 * text whose color is automatic in a document whose styles set none is
 * drawn in the shape style's color, white on an Accent 1 shape.
 */
export function readWordFont(
	xml: string,
	txbxContent: Span | undefined,
	theme: ShapeTheme | undefined,
	style?: Span,
	styles?: WordTextStyles,
): ShapeFont | undefined {
	if (!txbxContent) { return undefined; }
	const inner = xml.slice(txbxContent.openEnd, txbxContent.end);
	const paragraphs = [...inner.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
	const paragraph = paragraphs.find((p) => /<w:r(?=[\s>])/.test(p)) ?? paragraphs[0] ?? '';
	const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(paragraph)?.[0] ?? '';
	const run = /<w:r(?=[\s>])[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)?/.exec(paragraph);
	const rPr = run ? run[1] : /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPr)?.[0];
	const font = wordRunFontOf(rPr, theme);
	if (styles) {
		inheritWordStyle(font, styles, rPr ? /<w:rStyle\b[^>]*\bw:val="([^"]*)"/.exec(rPr)?.[1] : undefined, theme);
		inheritWordStyle(font, styles, /<w:pStyle\b[^>]*\bw:val="([^"]*)"/.exec(pPr)?.[1] ?? styles.defaultParagraph, theme);
		inheritFont(font, wordRunFontOf(styles.defaults, theme));
	}
	const ref = styleRef(xml, style, 'a:fontRef');
	if (ref) {
		const idx = attr(xml.slice(ref.start, ref.openEnd), 'idx');
		const name = idx === 'major' ? theme?.majorFont : idx === 'minor' ? theme?.minorFont : undefined;
		const color = resolveColor(xml, ref, theme).color;
		inheritFont(font, { ...(name ? { name } : {}), ...(color ? { color } : {}) });
	}
	return Object.keys(font).length > 0 ? font : undefined;
}

// ---------------------------------------------------------------- writing

/** The schema order of a shape's properties element, for the children written here. */
const SPPR_ORDER = [
	'a:xfrm', 'a:custGeom', 'a:prstGeom',
	'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill',
	'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst',
];
const LN_ORDER = [
	'a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:prstDash', 'a:custDash',
	'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd', 'a:extLst',
];
const RPR_ORDER = [
	'a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill',
	'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill',
	'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst',
];

/** An element with its self-closing form opened, so children can go in. */
function opened(elementXml: string, name: string): string {
	return /\/>$/.test(elementXml) && !elementXml.includes(`</${name}>`)
		? `${elementXml.slice(0, -2)}></${name}>`
		: elementXml;
}

/** `elementXml` (one whole element) with every direct child named in `names` removed. */
function withoutChildren(elementXml: string, names: readonly string[]): string {
	const openEnd = elementXml.indexOf('>') + 1;
	let out = elementXml;
	for (const child of children(elementXml, openEnd, elementXml.length).reverse()) {
		if (names.includes(child.name)) { out = splice(out, child, ''); }
	}
	return out;
}

/** `elementXml` (one whole element, opened) with `text` inserted as a child in schema order. */
function withChild(elementXml: string, name: string, order: readonly string[], text: string): string {
	const openEnd = elementXml.indexOf('>') + 1;
	const close = elementXml.lastIndexOf('</');
	return insertInOrder(elementXml, openEnd, close, order, name, text);
}

function solidFillXml(color: string, transparency?: number): string {
	const hex = color.slice(1);
	if (transparency === undefined || transparency <= 0) { return `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`; }
	const alpha = Math.round((100 - Math.min(100, transparency)) * 1000);
	return `<a:solidFill><a:srgbClr val="${hex}"><a:alpha val="${alpha}"/></a:srgbClr></a:solidFill>`;
}

/** A shape's properties element (the whole element) with its fill set. */
export function withFill(spPrXml: string, spPrName: string, fill: FillEdit): string {
	let out = withoutChildren(opened(spPrXml, spPrName), FILLS);
	if (fill.type === 'none') {
		return withChild(out, 'a:noFill', SPPR_ORDER, '<a:noFill/>');
	}
	const color = requireColor(fill.color, 'fill');
	if (fill.transparency !== undefined && (fill.transparency < 0 || fill.transparency > 100)) {
		throw new ShapeFormatError(`Transparency is a percentage from 0 to 100; ${fill.transparency} is not one.`);
	}
	out = withChild(out, 'a:solidFill', SPPR_ORDER, solidFillXml(color, fill.transparency));
	return out;
}

/** A shape's properties element (the whole element) with its outline set. */
export function withLine(spPrXml: string, spPrName: string, line: LineEdit): string {
	let props = opened(spPrXml, spPrName);
	const openEnd = props.indexOf('>') + 1;
	let ln = children(props, openEnd, props.length).find((child) => child.name === 'a:ln');
	if (!ln) {
		props = withChild(props, 'a:ln', SPPR_ORDER, '<a:ln></a:ln>');
		ln = children(props, props.indexOf('>') + 1, props.length).find((child) => child.name === 'a:ln')!;
	}
	let lnXml = opened(props.slice(ln.start, ln.end), 'a:ln');
	const existingFill = children(lnXml, lnXml.indexOf('>') + 1, lnXml.length).find((child) => FILLS.includes(child.name));
	if (line.type === 'none') {
		lnXml = withChild(withoutChildren(lnXml, FILLS), 'a:noFill', LN_ORDER, '<a:noFill/>');
		return splice(props, ln, lnXml);
	}
	if (line.weight !== undefined) {
		if (!(line.weight > 0) || line.weight > 1584) {
			throw new ShapeFormatError(`An outline weight is in points, above 0 and up to 1584; ${line.weight} is not one.`);
		}
		const tag = lnXml.slice(0, lnXml.indexOf('>') + 1);
		lnXml = withAttr(tag, 'w', String(pointsToEmu(line.weight))) + lnXml.slice(tag.length);
	}
	if (line.color !== undefined) {
		const color = requireColor(line.color, 'outline');
		lnXml = withChild(withoutChildren(lnXml, FILLS), 'a:solidFill', LN_ORDER, solidFillXml(color));
	} else if (!existingFill || existingFill.name !== 'a:solidFill') {
		// Turning an outline on without saying its color: black, which is
		// what a plain line is, rather than nothing a reader can see.
		lnXml = withChild(withoutChildren(lnXml, FILLS), 'a:solidFill', LN_ORDER, solidFillXml('#000000'));
	}
	if (line.dash !== undefined) {
		if (!SHAPE_DASHES.includes(line.dash)) {
			throw new ShapeFormatError(`'${line.dash}' is not a dash style; use one of ${SHAPE_DASHES.join(', ')}.`);
		}
		lnXml = withChild(withoutChildren(lnXml, ['a:prstDash', 'a:custDash']), 'a:prstDash', LN_ORDER, `<a:prstDash val="${line.dash}"/>`);
	}
	return splice(props, ln, lnXml);
}

/** Degrees normalized to [0, 360), checked. */
export function checkedRotation(degrees: number): number {
	if (!Number.isFinite(degrees)) {
		throw new ShapeFormatError(`A rotation is a number of degrees; ${degrees} is not one.`);
	}
	return ((degrees % 360) + 360) % 360;
}

/** An a:xfrm start tag with its rotation set; 0 removes it, as the applications do. */
export function withRotation(xfrmTag: string, degrees: number): string {
	const rot = Math.round(checkedRotation(degrees) * 60000);
	return withAttr(xfrmTag, 'rot', rot ? String(rot) : undefined);
}

/**
 * Whether a rotation lies where Office treats a box as turned a quarter:
 * from 45 up to 135 degrees, and from 225 up to 315. Excel keeps the cell
 * anchor of such a shape as its box rotated 90 degrees about the center,
 * and Word measures the room a floating shape's ink takes (wp:effectExtent)
 * from that turned box. Word's wp:extent and its VML twin's box stay as
 * they were: a box turned 90 degrees kept both, measured.
 */
export function isQuarterTurned(degrees: number | undefined): boolean {
	const d = checkedRotation(degrees ?? 0);
	return (d >= 45 && d < 135) || (d >= 225 && d < 315);
}

function checkedFont(font: FontEdit): void {
	if (font.size !== undefined && (!(font.size >= 1) || font.size > 409)) {
		throw new ShapeFormatError(`A font size is in points, from 1 to 409; ${font.size} is not one.`);
	}
	if (font.name !== undefined && !font.name.trim()) {
		throw new ShapeFormatError('A font needs a name.');
	}
	if (font.color) { requireColor(font.color, 'font'); }
}

/** One a:rPr or a:endParaRPr (the whole element) with a font change applied. */
function withRunFont(rPrXml: string, name: string, font: FontEdit): string {
	let out = opened(rPrXml, name);
	let tag = out.slice(0, out.indexOf('>') + 1);
	const body = out.slice(tag.length);
	if (font.size !== undefined) { tag = withAttr(tag, 'sz', String(Math.round(font.size * 100))); }
	if (font.bold !== undefined) { tag = withAttr(tag, 'b', font.bold ? '1' : '0'); }
	if (font.italic !== undefined) { tag = withAttr(tag, 'i', font.italic ? '1' : '0'); }
	if (font.underline !== undefined) { tag = withAttr(tag, 'u', font.underline ? 'sng' : 'none'); }
	out = tag + body;
	if (font.color !== undefined) {
		out = withoutChildren(out, FILLS);
		if (font.color) { out = withChild(out, 'a:solidFill', RPR_ORDER, solidFillXml(normalizeColor(font.color)!)); }
	}
	if (font.name !== undefined) {
		out = withChild(withoutChildren(out, ['a:latin']), 'a:latin', RPR_ORDER, `<a:latin typeface="${encodeXml(font.name.trim())}"/>`);
	}
	// A run property element left with no children collapses again.
	return out.replace(new RegExp(`^(<${name}\\b[^>]*)></${name}>$`), '$1/>');
}

/**
 * A DrawingML text body (the whole element) with a font change applied to
 * every run and every paragraph's end mark, which is what setting the font
 * of a shape does in each application. A run with no a:rPr gets one.
 */
export function withDrawingFont(txBodyXml: string, font: FontEdit): string {
	checkedFont(font);
	let out = txBodyXml.replace(/<a:r>(?!\s*<a:rPr\b)/g, '<a:r><a:rPr lang="en-US"/>');
	out = out.replace(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/g, (m) => withRunFont(m, 'a:rPr', font));
	out = out.replace(/<a:endParaRPr\b[^>]*\/>|<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/g, (m) => withRunFont(m, 'a:endParaRPr', font));
	return out;
}

/** The schema order of w:rPr, as far as a font touches it. */
const WORD_RPR_ORDER = [
	'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike',
	'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden',
	'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect',
	'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout',
	'w:specVanish', 'w:oMath',
];

/** One w:rPr (the whole element) with a font change applied, as Word writes one. */
function withWordRunFont(rPrXml: string, font: FontEdit): string {
	let out = opened(rPrXml, 'w:rPr');
	if (font.name !== undefined) {
		const face = encodeXml(font.name.trim());
		out = withChild(withoutChildren(out, ['w:rFonts']), 'w:rFonts', WORD_RPR_ORDER, `<w:rFonts w:ascii="${face}" w:hAnsi="${face}" w:cs="${face}"/>`);
	}
	if (font.bold !== undefined) {
		out = withoutChildren(out, ['w:b', 'w:bCs']);
		if (font.bold) { out = withChild(out, 'w:b', WORD_RPR_ORDER, '<w:b/>'); }
	}
	if (font.italic !== undefined) {
		out = withoutChildren(out, ['w:i', 'w:iCs']);
		if (font.italic) { out = withChild(out, 'w:i', WORD_RPR_ORDER, '<w:i/>'); }
	}
	if (font.color !== undefined) {
		out = withoutChildren(out, ['w:color']);
		if (font.color) { out = withChild(out, 'w:color', WORD_RPR_ORDER, `<w:color w:val="${normalizeColor(font.color)!.slice(1)}"/>`); }
	}
	if (font.size !== undefined) {
		out = withChild(withoutChildren(out, ['w:sz', 'w:szCs']), 'w:sz', WORD_RPR_ORDER, `<w:sz w:val="${Math.round(font.size * 2)}"/>`);
	}
	if (font.underline !== undefined) {
		out = withoutChildren(out, ['w:u']);
		if (font.underline) { out = withChild(out, 'w:u', WORD_RPR_ORDER, '<w:u w:val="single"/>'); }
	}
	return out.replace(/^<w:rPr><\/w:rPr>$/, '');
}

/**
 * A w:txbxContent (the whole element) with a font change applied to every
 * run and every paragraph mark, as Word writes it when the font of a text
 * box is set: the same w:rPr on each run and in each paragraph's w:pPr.
 */
export function withWordFont(txbxContentXml: string, font: FontEdit): string {
	checkedFont(font);
	// Every run gets a w:rPr to carry the change. A self-closing run or
	// paragraph has no content to format and is left alone.
	let out = txbxContentXml.replace(/(<w:r(?:\s[^>]*[^/])?>)(?!\s*<w:rPr\b)/g, '$1<w:rPr></w:rPr>');
	// Every paragraph gets a w:pPr, and every w:pPr a w:rPr for its mark.
	out = out.replace(/(<w:p(?:\s[^>]*[^/])?>)(?!\s*<w:pPr\b)/g, '$1<w:pPr></w:pPr>');
	out = out.replace(/<w:pPr\/>/g, '<w:pPr></w:pPr>');
	out = out.replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/g, (pPr) => (/<w:rPr\b/.test(pPr)
		? pPr
		: pPr.replace(/<\/w:pPr>$/, '<w:rPr></w:rPr></w:pPr>')));
	out = out.replace(/<w:rPr\/>|<w:rPr>[\s\S]*?<\/w:rPr>/g, (rPr) => withWordRunFont(rPr, font));
	return out.replace(/<w:pPr><\/w:pPr>/g, '');
}

// ------------------------------------------------------------------- VML

/** A VML color attribute value: #RRGGBB, or a name or "[index]" form Office writes. */
export function vmlColorOf(value: string | undefined): string | undefined {
	if (!value) { return undefined; }
	const hex = /^#([0-9a-fA-F]{6})\b/.exec(value.trim())?.[1];
	if (hex) { return `#${hex.toUpperCase()}`; }
	const named = PRESET_COLORS[value.trim().split(/\s/)[0].toLowerCase()];
	return named ? `#${named}` : undefined;
}

/** DrawingML's dash names as VML spells them. */
export const VML_DASH: Record<ShapeDash, string> = {
	solid: 'solid', dot: 'dot', dash: 'dash', lgDash: 'longDash', dashDot: 'dashDot', lgDashDot: 'longDashDot',
	lgDashDotDot: 'longDashDotDot', sysDash: 'shortDash', sysDot: 'shortDot', sysDashDot: 'shortDashDot',
	sysDashDotDot: 'shortDashDotDot',
};

/** A VML style attribute's value with one property set, or removed when `value` is undefined. */
export function withVmlStyle(style: string, name: string, value: string | undefined): string {
	const parts = style.split(';').map((part) => part.trim()).filter(Boolean)
		.filter((part) => part.split(':')[0].trim().toLowerCase() !== name.toLowerCase());
	if (value !== undefined) { parts.push(`${name}:${value}`); }
	return parts.join(';');
}

/**
 * A VML element split into its start tag and body, with a self-closing one
 * opened when `open` is set, so a child can go in.
 */
function vmlParts(shapeXml: string, open: boolean): { tag: string; body: string } {
	const openEnd = shapeXml.indexOf('>') + 1;
	const tag = shapeXml.slice(0, openEnd);
	if (!open || !tag.endsWith('/>')) { return { tag, body: shapeXml.slice(openEnd) }; }
	const name = /^<([\w:]+)/.exec(tag)![1];
	return { tag: `${tag.slice(0, -2)}>`, body: `</${name}>` };
}

/**
 * A VML shape with a fill set, as Word writes a shape's twin and Excel a
 * check box: `fillcolor`, `filled="f"` for none, and the opacity on a v:fill
 * child in 65,536ths. Word takes the opacity from the alpha as a byte and
 * widens it to 16 bits by 257: 25% transparent is byte 191, written
 * 49087f. This writes the same.
 */
export function withVmlFill(shapeXml: string, fill: FillEdit): string {
	if (fill.type === 'none') {
		const { tag, body } = vmlParts(shapeXml, false);
		return withAttr(tag, 'filled', 'f') + body;
	}
	const opacity = fill.transparency
		? `${Math.round(((100 - fill.transparency) / 100) * 255) * 257}f`
		: undefined;
	const vFill = /<v:fill\b[^>]*\/>|<v:fill\b[^>]*>[\s\S]*?<\/v:fill>/.exec(shapeXml);
	let { tag, body } = vmlParts(shapeXml, opacity !== undefined && !vFill);
	tag = withAttr(withAttr(tag, 'filled', undefined), 'fillcolor', normalizeColor(fill.color)!);
	if (vFill) {
		const fillTag = vFill[0].slice(0, vFill[0].indexOf('>') + 1);
		body = body.replace(fillTag, withAttr(fillTag, 'opacity', opacity));
	} else if (opacity) {
		body = `<v:fill opacity="${opacity}"/>${body}`;
	}
	return tag + body;
}

/** A VML shape with an outline set: `strokecolor`, `strokeweight`, `stroked="f"` and a v:stroke dash. */
export function withVmlLine(shapeXml: string, line: LineEdit): string {
	if (line.type === 'none') {
		const { tag, body } = vmlParts(shapeXml, false);
		return withAttr(tag, 'stroked', 'f') + body;
	}
	const dash = line.dash === undefined || line.dash === 'solid' ? undefined : VML_DASH[line.dash];
	const stroke = /<v:stroke\b[^>]*\/>|<v:stroke\b[^>]*>[\s\S]*?<\/v:stroke>/.exec(shapeXml);
	let { tag, body } = vmlParts(shapeXml, dash !== undefined && !stroke);
	tag = withAttr(tag, 'stroked', undefined);
	if (line.color) { tag = withAttr(tag, 'strokecolor', normalizeColor(line.color)!); }
	if (line.weight !== undefined) { tag = withAttr(tag, 'strokeweight', `${line.weight}pt`); }
	if (line.dash !== undefined) {
		if (stroke) {
			const strokeTag = stroke[0].slice(0, stroke[0].indexOf('>') + 1);
			body = body.replace(strokeTag, withAttr(strokeTag, 'dashstyle', dash));
		} else if (dash) {
			body = `<v:stroke dashstyle="${dash}"/>${body}`;
		}
	}
	return tag + body;
}
