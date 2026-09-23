// Document shapes: what agents can see and change on a Word surface.
//
// Word is the odd host of the three. A shape is not a child of one drawing
// part but a run inside a paragraph, and it sits in whichever STORY the text
// belongs to: the document body, a header, a footer, the footnotes. Each
// story is its own package part, so a reader that knows only
// word/document.xml misses the logo in the header - and the part numbering
// says nothing about which is which, since Word writes the even-page header
// as header1.xml and the one every other page uses as header2.xml. The
// w:sectPr references are the map.
//
// A shape is written twice. mc:Choice holds the modern DrawingML
// (wp:inline or wp:anchor, then wps:wsp), and mc:Fallback holds a VML twin
// for Word 2007 carrying the same name, text and geometry. This module
// keeps both in step, because a file whose two copies disagree renders
// differently depending on who opens it.
//
// A Word shape CANNOT run a macro. Its Shape has neither OnAction nor
// ActionSettings - checked against the live type library, hidden members
// included - and a document Word saved with a shape carries no macro link
// in any form. Word attaches macros to form fields instead, so this module
// refuses `macro` with that explanation rather than writing an attribute no
// host would read.
//
// Word also refuses to GROUP free-floating shapes at all: every wrap style
// and a shared anchor were tried and each was refused with "Grouping is
// disabled for the selected shapes". A group in Word lives inside a drawing
// canvas, so a canvas is listed as a container and its members are read
// from the wpg:wgp inside it.
//
// `placement` is read from the markup - wp:inline or wp:anchor - and not
// from Word's collections, which do not draw the line there: Word reports
// InlineShapes.Count as 0 for a document of its own making whose AutoShape
// it converted to inline, and lists that shape under Shapes with the
// floating ones. The markup is what decides whether there is a position to
// set, so the markup is what is reported.
//
// Every layout here was read from files Word 16 saved, and the writer is
// checked by having Word open and re-save the result.

import { ZipArchive } from './zip';
import {
	Package,
	attr,
	children,
	decodeXml,
	emuToPoints,
	encodeXml,
	findElement,
	pointsToEmu,
	splice,
	withAttr,
	xfrmBox,
	type Span,
} from './ooxml';
import {
	PRESET_GEOMETRY,
	PRESET_LABELS,
	ShapeError,
	lookOf,
	restackedIndex,
	type ShapeEdit,
	type ShapeInfo,
	type ShapeKind,
	type PresetShapeType,
	type ZOrderCommand,
} from './shapes';
import {
	checkedRotation,
	isQuarterTurned,
	readFill,
	readLine,
	readRotation,
	readTheme,
	readWordFont,
	readWordStyles,
	withFill,
	withLine,
	withRotation,
	withVmlFill,
	withVmlLine,
	withVmlStyle,
	withWordFont,
	type ShapeTheme,
	type WordTextStyles,
} from './shapeFormat';
import { STORY_FORMATTABLE } from './shapeCapabilities';

const DOCUMENT = 'word/document.xml';
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

/** What a document's shapes are styled from: its theme, and its text styles. */
interface WordLook {
	theme: ShapeTheme;
	styles: WordTextStyles;
}

/** The document's theme and styles; a header or footer uses the same ones. */
function documentLook(pkg: Package): WordLook {
	const rels = pkg.has(DOCUMENT) ? pkg.relationships(DOCUMENT) : [];
	const partOf = (type: string): string | undefined => {
		const path = rels.find((rel) => rel.type === type)?.path;
		return path && pkg.has(path) ? pkg.read(path) : undefined;
	};
	return { theme: readTheme(partOf(THEME_REL)), styles: readWordStyles(partOf(STYLES_REL)) };
}

/**
 * Why a Word shape takes no macro. One copy, because the caller layer has to
 * refuse before it validates a macro against the project, and two wordings
 * for one fact is how a user ends up with the unhelpful one.
 */
export const WORD_SHAPE_MACRO_REFUSAL =
	'Word cannot run a macro from a shape: its Shape has neither OnAction nor '
	+ 'ActionSettings, and the file format carries no macro link for one. Word runs a '
	+ "macro from a form field's EntryMacro or ExitMacro, or from a control on the ribbon.";

/** What a:graphicData's uri says is inside it. */
const GRAPHIC = {
	shape: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
	group: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
	canvas: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
	picture: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
	chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
};

// ----------------------------------------------------------------- surfaces

export type StoryKind = 'document' | 'header' | 'footer' | 'footnotes' | 'endnotes' | 'comments';

/** One surface a Word shape can sit on: a story, which is one package part. */
export interface StoryRef {
	name: string;
	kind: StoryKind;
	path: string;
}

/** What Word's own Header and Footer menu calls each reference type. */
const REFERENCE_LABEL: Record<string, string> = {
	default: '',
	first: 'First Page ',
	even: 'Even Page ',
};

/**
 * Every surface of a document: the body first, then each header and footer
 * in section order, then the note and comment stories. A part referenced
 * from more than one section keeps the first name it was given, since it is
 * one surface either way.
 */
export function documentStories(zip: ZipArchive): StoryRef[] {
	const pkg = new Package(zip);
	if (!pkg.has(DOCUMENT)) { return []; }
	const out: StoryRef[] = [{ name: 'Document', kind: 'document', path: DOCUMENT }];
	const seen = new Set([DOCUMENT]);
	const rels = new Map(pkg.relationships(DOCUMENT).map((rel) => [rel.id, rel.path]));
	const xml = pkg.read(DOCUMENT);
	const sections = [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g)];
	for (const [index, [section]] of sections.entries()) {
		const suffix = sections.length > 1 ? ` (Section ${index + 1})` : '';
		for (const match of section.matchAll(/<w:(header|footer)Reference\b[^>]*>/g)) {
			const [tag, which] = match;
			const path = rels.get(attr(tag, 'r:id') ?? '');
			if (!path || !pkg.has(path) || seen.has(path)) { continue; }
			seen.add(path);
			const label = REFERENCE_LABEL[attr(tag, 'w:type') ?? 'default'] ?? '';
			const noun = which === 'header' ? 'Header' : 'Footer';
			out.push({ name: `${label}${noun}${suffix}`, kind: which as StoryKind, path });
		}
	}
	const NOTE_STORIES: Array<[StoryKind, string]> = [
		['footnotes', 'Footnotes'], ['endnotes', 'Endnotes'], ['comments', 'Comments'],
	];
	for (const [kind, name] of NOTE_STORIES) {
		const path = `word/${kind}.xml`;
		if (pkg.has(path) && !seen.has(path)) {
			seen.add(path);
			out.push({ name, kind, path });
		}
	}
	return out;
}

/** The surface a caller named, by its name or by its part file name. */
export function requireStory(stories: readonly StoryRef[], surface: string): StoryRef {
	const wanted = surface.trim().toLowerCase();
	const byName = stories.find((story) => story.name.toLowerCase() === wanted);
	if (byName) { return byName; }
	const byPart = stories.find((story) => story.path.toLowerCase() === `word/${wanted}.xml`
		|| story.path.toLowerCase() === wanted);
	if (byPart) { return byPart; }
	throw new ShapeError(`No surface named '${surface}'. The document has ${stories.map((s) => `'${s.name}'`).join(', ')}.`);
}

/** The body surface, which is what a caller who names none means. */
export function defaultStory(stories: readonly StoryRef[]): StoryRef {
	const body = stories.find((story) => story.kind === 'document');
	if (!body) {
		throw new ShapeError('The package has no word/document.xml; it is not a Word document.');
	}
	return body;
}

// ------------------------------------------------------------------ reading

/** A shape found in a story, with where its markup sits in the part. */
interface DocShape {
	info: ShapeInfo;
	/** The mc:AlternateContent, or the bare w:drawing when there is none. */
	entry: Span;
	/** The wps:wsp, wpg:wgp, pic:pic or wpc:wpc that is the shape itself. */
	element: Span & { name: string };
	id: number;
	inGroup: boolean;
	/**
	 * How a floating shape stacks: behind the text or in front of it, and
	 * its wp:anchor relativeHeight, the higher on top. A shape inline with
	 * the text does not stack.
	 */
	stack?: { behind: boolean; height: number };
}

/** What a kind of Word shape can be given: which of fill, line, font and rotation apply. */
const FORMATTABLE = STORY_FORMATTABLE;

/** The element holding a shape's fill, outline and transform. */
const PROPS_OF: Record<string, string> = { 'wps:wsp': 'wps:spPr', 'pic:pic': 'pic:spPr', 'wpg:wgp': 'wpg:grpSpPr' };

/** A shape's fill, outline, font and rotation, as far as its kind has them. */
function formatOf(xml: string, element: Span & { name: string }, kind: ShapeKind, look: WordLook): Partial<ShapeInfo> {
	const out: Partial<ShapeInfo> = {};
	const propsName = PROPS_OF[element.name];
	if (!propsName) { return out; }
	const direct = children(xml, element.openEnd, element.end);
	const props = direct.find((child) => child.name === propsName);
	const style = direct.find((child) => child.name === 'wps:style');
	if (FORMATTABLE.rotation.includes(kind) && props) {
		const xfrm = findElement(xml, 'a:xfrm', props.openEnd, props.end);
		const rotation = readRotation(xfrm ? xml.slice(xfrm.start, xfrm.openEnd) : undefined);
		if (rotation) { out.rotation = rotation; }
	}
	if (FORMATTABLE.fill.includes(kind)) {
		const fill = readFill(xml, props, style, look.theme);
		if (fill) { out.fill = fill; }
	}
	if (FORMATTABLE.line.includes(kind)) {
		const line = readLine(xml, props, style, look.theme);
		if (line) { out.line = line; }
	}
	if (FORMATTABLE.font.includes(kind)) {
		const box = direct.find((child) => child.name === 'wps:txbx');
		const content = box ? findElement(xml, 'w:txbxContent', box.openEnd, box.end) : undefined;
		const font = readWordFont(xml, content, look.theme, style, look.styles);
		if (font) { out.font = font; }
	}
	return out;
}

/** Floating shapes from the back up: behind the text first, then by height, then in document order. */
function stackOrder(shapes: readonly DocShape[]): DocShape[] {
	return shapes.filter((s) => s.stack)
		.sort((a, b) => Number(!a.stack!.behind) - Number(!b.stack!.behind) || a.stack!.height - b.stack!.height);
}

/** The element naming each kind of nested member. */
const MEMBER_NAME_ELEMENT: Record<string, string> = {
	'wps:wsp': 'wps:cNvPr',
	'wpg:wgp': 'wpg:cNvPr',
	'pic:pic': 'pic:cNvPr',
};

function kindOfGraphic(uri: string | undefined, inner: string): ShapeKind {
	switch (uri) {
		case GRAPHIC.shape: return /<wps:cNvSpPr\b[^>]*\btxBox="1"/.test(inner) ? 'textBox' : 'shape';
		case GRAPHIC.group: return 'group';
		case GRAPHIC.canvas: return 'canvas';
		case GRAPHIC.picture: return 'picture';
		case GRAPHIC.chart: return 'chart';
		default: return 'other';
	}
}

/** The text of a w:txbxContent: one line per paragraph, as Word shows it. */
function wordTextOf(xml: string, element: Span): string | undefined {
	const box = findElement(xml, 'w:txbxContent', element.openEnd, element.end);
	if (!box) { return undefined; }
	return [...xml.slice(box.openEnd, box.end).matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>|<w:p\b[^>]*\/>/g)]
		.map((p) => [...(p[1] ?? '').matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:t\b[^>]*\/>/g)]
			.map((t) => decodeXml(t[1] ?? '')).join(''))
		.join('\n');
}

function offsetOf(xml: string, anchor: Span, which: string): number | undefined {
	const position = findElement(xml, which, anchor.openEnd, anchor.end);
	if (!position) { return undefined; }
	const offset = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(xml.slice(position.start, position.end));
	return offset ? emuToPoints(Number(offset[1])) : undefined;
}

/** A top-level shape's size from its wp:extent, and its offset from a wp:anchor. */
function anchorBox(xml: string, anchor: Span & { name: string }): Partial<ShapeInfo> {
	const box: Partial<ShapeInfo> = { placement: anchor.name === 'wp:inline' ? 'inline' : 'anchored' };
	const extent = /<wp:extent\b[^>]*>/.exec(xml.slice(anchor.start, anchor.end))?.[0];
	if (extent) {
		box.width = emuToPoints(Number(attr(extent, 'cx') ?? 0));
		box.height = emuToPoints(Number(attr(extent, 'cy') ?? 0));
	}
	if (anchor.name === 'wp:anchor') {
		// A shape placed by wp:align ("centered", "right") carries no offset,
		// and reporting one would be inventing a number.
		const left = offsetOf(xml, anchor, 'wp:positionH');
		const top = offsetOf(xml, anchor, 'wp:positionV');
		if (left !== undefined) { box.left = left; }
		if (top !== undefined) { box.top = top; }
	}
	return box;
}

/** A nested member's box, from the a:xfrm of its own properties element. */
function memberBox(xml: string, element: Span & { name: string }): Partial<ShapeInfo> {
	const wanted = element.name === 'wpg:wgp' ? 'wpg:grpSpPr'
		: element.name === 'pic:pic' ? 'pic:spPr' : 'wps:spPr';
	const props = children(xml, element.openEnd, element.end).find((child) => child.name === wanted);
	if (!props) { return {}; }
	const xfrm = findElement(xml, 'a:xfrm', props.start, props.end);
	return xfrm ? xfrmBox(xml, xfrm) : {};
}

/** The members of a canvas or group: its direct children, and all descendants. */
function readMembers(xml: string, container: Span, entry: Span, look: WordLook): { direct: DocShape[]; all: DocShape[] } {
	const direct: DocShape[] = [];
	const all: DocShape[] = [];
	for (const child of children(xml, container.openEnd, container.end)) {
		const nameElementName = MEMBER_NAME_ELEMENT[child.name];
		if (!nameElementName) { continue; }
		const nameElement = findElement(xml, nameElementName, child.openEnd, child.end);
		if (!nameElement) { continue; }
		const tag = xml.slice(nameElement.start, nameElement.openEnd);
		const inner = xml.slice(child.start, child.end);
		const kind: ShapeKind = child.name === 'wpg:wgp' ? 'group'
			: child.name === 'pic:pic' ? 'picture'
				: /<wps:cNvSpPr\b[^>]*\btxBox="1"/.test(inner) ? 'textBox' : 'shape';
		const info: ShapeInfo = { name: attr(tag, 'name') ?? '', kind };
		const geometry = /<a:prstGeom\b[^>]*\bprst="([^"]*)"/.exec(inner)?.[1];
		if (geometry && (kind === 'shape' || kind === 'textBox')) { info.geometry = geometry; }
		Object.assign(info, memberBox(xml, child));
		if (kind === 'shape' || kind === 'textBox') {
			const text = wordTextOf(xml, child);
			if (text) { info.text = text; }
		}
		const descr = attr(tag, 'descr');
		if (descr) { info.altText = descr; }
		if (attr(tag, 'hidden') === '1') { info.hidden = true; }
		Object.assign(info, formatOf(xml, child, kind, look));
		const found: DocShape = { info, entry, element: child, id: Number(attr(tag, 'id') ?? 0), inGroup: true };
		direct.push(found);
		all.push(found);
		if (child.name === 'wpg:wgp') {
			const nested = readMembers(xml, child, entry, look);
			info.shapes = nested.direct.map((member) => member.info);
			all.push(...nested.all);
		}
	}
	return { direct, all };
}

function readTopLevel(xml: string, entry: Span, anchor: Span & { name: string }, look: WordLook): DocShape[] {
	const docPr = findElement(xml, 'wp:docPr', anchor.openEnd, anchor.end);
	if (!docPr) { return []; }
	const tag = xml.slice(docPr.start, docPr.openEnd);
	const data = findElement(xml, 'a:graphicData', anchor.openEnd, anchor.end);
	const uri = data ? attr(xml.slice(data.start, data.openEnd), 'uri') : undefined;
	const body = data ? children(xml, data.openEnd, data.end)[0] : undefined;
	const inner = xml.slice(anchor.start, anchor.end);
	const kind = kindOfGraphic(uri, inner);
	const info: ShapeInfo = { name: attr(tag, 'name') ?? '', kind };
	const geometry = /<a:prstGeom\b[^>]*\bprst="([^"]*)"/.exec(inner)?.[1];
	if (geometry && (kind === 'shape' || kind === 'textBox')) { info.geometry = geometry; }
	Object.assign(info, anchorBox(xml, anchor));
	if (kind === 'shape' || kind === 'textBox') {
		const text = wordTextOf(xml, anchor);
		if (text) { info.text = text; }
	}
	const descr = attr(tag, 'descr');
	if (descr) { info.altText = descr; }
	if (attr(tag, 'hidden') === '1') { info.hidden = true; }
	const element = body ?? { ...anchor, name: anchor.name };
	Object.assign(info, formatOf(xml, element, kind, look));
	const found: DocShape = { info, entry, element, id: Number(attr(tag, 'id') ?? 0), inGroup: false };
	if (anchor.name === 'wp:anchor') {
		const anchorTag = xml.slice(anchor.start, anchor.openEnd);
		found.stack = { behind: attr(anchorTag, 'behindDoc') === '1', height: Number(attr(anchorTag, 'relativeHeight') ?? 0) };
	}
	const out = [found];
	if (kind === 'canvas' || kind === 'group') {
		const members = readMembers(xml, element, entry, look);
		info.shapes = members.direct.map((member) => member.info);
		out.push(...members.all);
	}
	return out;
}

/** The mc:AlternateContent a span sits in, when it sits in one. */
function enclosingAlternateContent(xml: string, span: Span): Span | undefined {
	for (let at = 0; ;) {
		const found = findElement(xml, 'mc:AlternateContent', at);
		if (!found || found.start > span.start) { return undefined; }
		if (found.end >= span.end) { return found; }
		at = found.start + 1;
	}
}

/**
 * Every drawing in a story, with the VML twins inside mc:Fallback skipped.
 * A floating shape's zOrder counts the story's floating shapes from the
 * back; Word's own ZOrderPosition counts an inline shape as well, which
 * never overlaps anything.
 */
function readStory(pkg: Package, story: StoryRef): { xml: string; shapes: DocShape[] } {
	const xml = pkg.read(story.path);
	const look = documentLook(pkg);
	const fallbacks: Span[] = [];
	for (let at = 0; ;) {
		const span = findElement(xml, 'mc:Fallback', at);
		if (!span) { break; }
		fallbacks.push(span);
		at = span.end;
	}
	const shapes: DocShape[] = [];
	for (let at = 0; ;) {
		const drawing = findElement(xml, 'w:drawing', at);
		if (!drawing) { break; }
		at = drawing.end;
		if (fallbacks.some((fallback) => drawing.start >= fallback.start && drawing.end <= fallback.end)) { continue; }
		const anchor = children(xml, drawing.openEnd, drawing.end)
			.find((child) => child.name === 'wp:inline' || child.name === 'wp:anchor');
		if (!anchor) { continue; }
		// The entry is the AlternateContent when there is one, so a delete
		// takes the VML twin with it and an edit can reach both copies.
		shapes.push(...readTopLevel(xml, enclosingAlternateContent(xml, drawing) ?? drawing, anchor, look));
	}
	stackOrder(shapes).forEach((shape, index) => { shape.info.zOrder = index + 1; });
	return { xml, shapes };
}

/** Every shape on a surface, top level only; a group or canvas lists its members. */
export function listStoryShapes(zip: ZipArchive, story: StoryRef): ShapeInfo[] {
	return readStory(new Package(zip), story).shapes.filter((s) => !s.inGroup).map((s) => s.info);
}

// ------------------------------------------------------------------ editing

function findTarget(shapes: readonly DocShape[], name: string, surface: string): DocShape {
	const lower = name.toLowerCase();
	const matches = shapes.filter((s) => s.info.name.toLowerCase() === lower);
	if (matches.length === 0) {
		throw new ShapeError(`No shape named '${name}' on ${surface}.`);
	}
	if (matches.length > 1) {
		throw new ShapeError(`${matches.length} shapes on ${surface} are named '${name}'; rename one in Word first.`);
	}
	return matches[0];
}

function assertNameFree(shapes: readonly DocShape[], name: string, surface: string): void {
	if (shapes.some((s) => s.info.name.toLowerCase() === name.toLowerCase())) {
		throw new ShapeError(`${surface} already has a shape named '${name}'.`);
	}
}

/**
 * The element of `elementName` whose name element says `shapeName`. Splices
 * move everything after them, so each step of an edit finds its target
 * again by name rather than by an offset taken before the last one.
 */
function relocate(xml: string, elementName: string, nameElementName: string, shapeName: string): Span | undefined {
	for (let at = 0; ;) {
		const span = findElement(xml, elementName, at);
		if (!span) { return undefined; }
		const nameElement = elementName === nameElementName
			? span
			: findElement(xml, nameElementName, span.openEnd, span.end);
		if (nameElement && attr(xml.slice(nameElement.start, nameElement.openEnd), 'name') === shapeName) {
			return span;
		}
		at = span.start + 1;
	}
}

/** Where a shape's own markup and its name element sit inside its entry. */
function locate(entryXml: string, shape: DocShape): { element: Span; nameElement: Span } {
	const nameElementName = shape.inGroup ? MEMBER_NAME_ELEMENT[shape.element.name] ?? 'wps:cNvPr' : 'wp:docPr';
	const elementName = shape.inGroup ? shape.element.name : 'wp:docPr';
	const found = relocate(entryXml, elementName, nameElementName, shape.info.name);
	if (!found) {
		throw new ShapeError(`'${shape.info.name}' could not be found again in its own markup; the part is not one Word wrote.`);
	}
	if (!shape.inGroup) {
		return { element: { start: 0, openEnd: 0, end: entryXml.length }, nameElement: found };
	}
	const nameElement = findElement(entryXml, nameElementName, found.openEnd, found.end)!;
	return { element: found, nameElement };
}

/** A w:txbxContent's paragraphs replaced, keeping the first one's formatting. */
function withWordText(boxXml: string, text: string): string {
	const pPr = /<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(boxXml)?.[0] ?? '';
	const rPr = /<w:rPr\b[^>]*\/>|<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(boxXml)?.[0] ?? '';
	const paragraphs = text.split(/\r?\n/).map((line) => (line
		? `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${encodeXml(line)}</w:t></w:r></w:p>`
		: `<w:p>${pPr}</w:p>`)).join('');
	const openEnd = boxXml.indexOf('>') + 1;
	return `${boxXml.slice(0, openEnd)}${paragraphs}</w:txbxContent>`;
}

/** Every w:txbxContent in a range replaced, so the DrawingML and VML twins agree. */
function setAllText(xml: string, within: Span, text: string): { xml: string; found: boolean } {
	let out = xml;
	let found = false;
	for (let at = within.start; at < within.end;) {
		const box = findElement(out, 'w:txbxContent', at, within.end);
		if (!box) { break; }
		found = true;
		const replaced = withWordText(out.slice(box.start, box.end), text);
		const shift = replaced.length - (box.end - box.start);
		out = splice(out, box, replaced);
		at = box.start + replaced.length;
		within = { ...within, end: within.end + shift };
	}
	return { xml: out, found };
}

/** A VML style attribute with one property set, added when it is absent. */
function withStyleProperty(style: string, name: string, value: string): string {
	const re = new RegExp(`(^|;)\\s*${name}\\s*:[^;]*`);
	return re.test(style) ? style.replace(re, `$1${name}:${value}`) : `${style};${name}:${value}`;
}

/** The VML twin of a top-level shape: the shape element in its entry's mc:Fallback. */
function vmlTwin(entryXml: string): (Span & { name: string }) | undefined {
	const fallback = findElement(entryXml, 'mc:Fallback');
	const pict = fallback ? findElement(entryXml, 'w:pict', fallback.openEnd, fallback.end) : undefined;
	return pict
		? children(entryXml, pict.openEnd, pict.end).find((child) => child.name.startsWith('v:') && child.name !== 'v:shapetype')
		: undefined;
}

/** An entry with its VML twin rewritten by `change`; an entry with no twin is left as it is. */
function withTwin(entryXml: string, change: (twinXml: string) => string): string {
	const twin = vmlTwin(entryXml);
	return twin ? splice(entryXml, twin, change(entryXml.slice(twin.start, twin.end))) : entryXml;
}

/** A VML element with its style attribute rewritten by `change`. */
function withVmlElementStyle(elementXml: string, change: (style: string) => string): string {
	const openEnd = elementXml.indexOf('>') + 1;
	const tag = elementXml.slice(0, openEnd);
	return withAttr(tag, 'style', change(attr(tag, 'style') ?? '')) + elementXml.slice(openEnd);
}

/**
 * The VML twin's geometry brought in line with the DrawingML one. Only the
 * twin itself: a canvas's members keep their own coordinates inside it.
 */
function movedVml(entryXml: string, box: { left?: number; top?: number; width: number; height: number }): string {
	return withTwin(entryXml, (twin) => withVmlElementStyle(twin, (style) => {
		let next = withStyleProperty(style, 'width', `${box.width}pt`);
		next = withStyleProperty(next, 'height', `${box.height}pt`);
		if (box.left !== undefined) { next = withStyleProperty(next, 'margin-left', `${box.left}pt`); }
		if (box.top !== undefined) { next = withStyleProperty(next, 'margin-top', `${box.top}pt`); }
		return next;
	}));
}

function withPosOffset(xml: string, which: string, points: number): string {
	const position = findElement(xml, which);
	if (!position) { return xml; }
	const inner = xml.slice(position.start, position.end);
	const next = /<wp:posOffset>/.test(inner)
		? inner.replace(/<wp:posOffset>-?\d+<\/wp:posOffset>/, `<wp:posOffset>${pointsToEmu(points)}</wp:posOffset>`)
		// The shape was placed by wp:align; an offset replaces it.
		: inner.replace(/<wp:align>[\s\S]*?<\/wp:align>/, `<wp:posOffset>${pointsToEmu(points)}</wp:posOffset>`);
	return splice(xml, position, next);
}

function movedShape(entryXml: string, shape: DocShape, edit: ShapeEdit): string {
	const width = edit.width ?? shape.info.width ?? 0;
	const height = edit.height ?? shape.info.height ?? 0;
	if (width <= 0 || height <= 0) {
		throw new ShapeError(`A shape needs a width and a height above zero; '${shape.info.name}' would be ${width} by ${height} points.`);
	}
	const left = edit.left ?? shape.info.left;
	const top = edit.top ?? shape.info.top;
	if (shape.inGroup) {
		const { element } = locate(entryXml, shape);
		const inner = entryXml.slice(element.start, element.end)
			.replace(/<a:off\b[^>]*\/>/, `<a:off x="${pointsToEmu(left ?? 0)}" y="${pointsToEmu(top ?? 0)}"/>`)
			.replace(/<a:ext\b[^>]*\/>/, `<a:ext cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/>`);
		return splice(entryXml, element, inner);
	}
	if (shape.info.placement === 'inline' && (edit.left !== undefined || edit.top !== undefined)) {
		throw new ShapeError(`'${shape.info.name}' is inline with the text, which Word positions by the text; set its size instead, or make it a floating shape in Word first.`);
	}
	// wp:effectExtent describes the ink outside the box, not the box itself.
	let out = entryXml.replace(/<wp:extent\b[^>]*\/>/, `<wp:extent cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/>`);
	if (left !== undefined) { out = withPosOffset(out, 'wp:positionH', left); }
	if (top !== undefined) { out = withPosOffset(out, 'wp:positionV', top); }
	// The shape's own transform, which carries rotation and a canvas's
	// interior, is sized to match the extent.
	out = out.replace(/(<a:ext\b[^>]*\bcx=")[^"]*("[^>]*\bcy=")[^"]*(")/,
		`$1${pointsToEmu(width)}$2${pointsToEmu(height)}$3`);
	return movedVml(out, { left, top, width, height });
}

/** The VML twin names the shape with its id attribute. */
function renamedVml(xml: string, from: string, to: string): string {
	return xml.replace(/(<v:(?:rect|oval|roundrect|shape|group)\b[^>]*\bid=")([^"]*)(")/g,
		(match, head: string, id: string, tail: string) => (id === from ? head + encodeXml(to) + tail : match));
}

/** A top-level shape's own element (wps:wsp, pic:pic, wpg:wgp, wpc:wpc) inside its entry. */
function drawingElement(entryXml: string): (Span & { name: string }) | undefined {
	const data = findElement(entryXml, 'a:graphicData');
	return data ? children(entryXml, data.openEnd, data.end)[0] : undefined;
}

/** Every w:txbxContent in `xml` rewritten by `change`. */
function withAllTextBoxes(xml: string, change: (boxXml: string) => string): string {
	let out = xml;
	for (let at = 0; ;) {
		const box = findElement(out, 'w:txbxContent', at);
		if (!box) { return out; }
		const replaced = change(out.slice(box.start, box.end));
		out = splice(out, box, replaced);
		at = box.start + replaced.length;
	}
}

/** A VML element with `child` added last, ahead of any w10 wrap settings. */
function withVmlChild(elementXml: string, child: string): string {
	const openEnd = elementXml.indexOf('>') + 1;
	if (openEnd === elementXml.length && elementXml.endsWith('/>')) {
		const name = /^<([\w:]+)/.exec(elementXml)![1];
		return `${elementXml.slice(0, -2)}>${child}</${name}>`;
	}
	const wrap = children(elementXml, openEnd, elementXml.length).find((c) => c.name.startsWith('w10:'));
	const at = wrap ? wrap.start : elementXml.lastIndexOf('</');
	return elementXml.slice(0, at) + child + elementXml.slice(at);
}

/**
 * The paragraph an AutoShape's text goes in: centered, as in the shapes
 * Word drew for the fixtures. A text box's text is left-aligned.
 */
const SHAPE_TEXT = '<w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:txbxContent>';

/**
 * An entry with `text` set on its shape. A shape with no text yet is given
 * a text box in each copy it has, the wps:txbx before its wps:bodyPr as the
 * schema orders them and a v:textbox in the VML twin.
 */
function withShapeText(entryXml: string, shape: DocShape, text: string): string {
	const within = shape.inGroup ? locate(entryXml, shape).element : { start: 0, openEnd: 0, end: entryXml.length };
	const result = setAllText(entryXml, within, text);
	if (result.found) { return result.xml; }
	const element = shape.inGroup ? locate(entryXml, shape).element : drawingElement(entryXml);
	if (!element || shape.element.name !== 'wps:wsp') {
		throw new ShapeError(`'${shape.info.name}' has no text box; give it text in Word first.`);
	}
	const content = withWordText(SHAPE_TEXT, text);
	const bodyPr = findElement(entryXml, 'wps:bodyPr', element.openEnd, element.end);
	const at = bodyPr ? bodyPr.start : element.end - '</wps:wsp>'.length;
	const out = `${entryXml.slice(0, at)}<wps:txbx>${content}</wps:txbx>${entryXml.slice(at)}`;
	return shape.inGroup ? out : withTwin(out, (twin) => withVmlChild(twin, `<v:textbox>${content}</v:textbox>`));
}

/** The refusal for a property a kind of shape does not have. */
function notFor(shape: ShapeInfo, what: string): ShapeError {
	return new ShapeError(`'${shape.name}' is a ${shape.kind}, which has no ${what} XLIDE sets.`);
}

/** A shape element with its properties element rewritten by `change`. */
function withProperties(elementXml: string, propsName: string, change: (propsXml: string) => string): string {
	const props = children(elementXml, elementXml.indexOf('>') + 1, elementXml.length).find((child) => child.name === propsName);
	if (!props) {
		throw new ShapeError(`The shape has no ${propsName}; the part is not one Word wrote.`);
	}
	return splice(elementXml, props, change(elementXml.slice(props.start, props.end)));
}

/** Degrees as VML writes them in a style: a whole number when it is one. */
const vmlDegrees = (degrees: number): string => String(Math.round(degrees * 100) / 100);

/**
 * The room a shape's ink takes past its box, as wp:effectExtent: the turned
 * shape with half its outline on every side, measured from the box Word
 * wraps text around. Word measured that box from the shape's own box, and
 * from the box turned a quarter for a shape turned 45 to 135 degrees (or
 * 225 to 315): at 90 degrees Word wrote only the outline's width, at 30 the
 * turned corners. Word's own figures come from its renderer and are not
 * symmetric; this is the geometry they approximate.
 */
function effectExtentXml(width: number, height: number, degrees: number, outline: number): string {
	const turn = (checkedRotation(degrees) * Math.PI) / 180;
	const cos = Math.abs(Math.cos(turn));
	const sin = Math.abs(Math.sin(turn));
	const w = width + outline;
	const h = height + outline;
	const [boxWidth, boxHeight] = isQuarterTurned(degrees) ? [height, width] : [width, height];
	const x = pointsToEmu(Math.max(0, (w * cos + h * sin - boxWidth) / 2));
	const y = pointsToEmu(Math.max(0, (w * sin + h * cos - boxHeight) / 2));
	return `<wp:effectExtent l="${x}" t="${y}" r="${x}" b="${y}"/>`;
}

/** A top-level entry's wp:effectExtent recomputed for the shape's rotation and outline after `edit`. */
function withEffectExtent(entryXml: string, shape: DocShape, edit: ShapeEdit): string {
	const extent = /<wp:extent\b[^>]*>/.exec(entryXml)?.[0];
	if (!extent) { return entryXml; }
	const width = emuToPoints(Number(attr(extent, 'cx') ?? 0));
	const height = emuToPoints(Number(attr(extent, 'cy') ?? 0));
	const degrees = edit.rotation ?? shape.info.rotation ?? 0;
	const before = shape.info.line?.type === 'solid' ? shape.info.line.weight ?? 0 : 0;
	const outline = edit.line === undefined ? before
		: edit.line.type === 'none' ? 0
			: edit.line.weight ?? before;
	return entryXml.replace(/<wp:effectExtent\b[^>]*\/>/, () => effectExtentXml(width, height, degrees, outline));
}

/** An entry with its shape's fill, outline, font and rotation set, in each copy Word keeps in step. */
function formattedEntry(entryXml: string, shape: DocShape, edit: ShapeEdit): string {
	const { info } = shape;
	if (edit.fill === undefined && edit.line === undefined && edit.font === undefined && edit.rotation === undefined) {
		return entryXml;
	}
	if (edit.fill !== undefined && !FORMATTABLE.fill.includes(info.kind)) { throw notFor(info, 'fill'); }
	if (edit.line !== undefined && !FORMATTABLE.line.includes(info.kind)) { throw notFor(info, 'outline'); }
	if (edit.font !== undefined && !FORMATTABLE.font.includes(info.kind)) { throw notFor(info, 'text to give a font'); }
	if (edit.rotation !== undefined && !FORMATTABLE.rotation.includes(info.kind)) { throw notFor(info, 'rotation'); }
	const span = shape.inGroup ? locate(entryXml, shape).element : drawingElement(entryXml);
	const propsName = PROPS_OF[shape.element.name];
	if (!span || !propsName) {
		throw new ShapeError(`'${info.name}' has no drawing properties; the part is not one Word wrote.`);
	}
	let element = entryXml.slice(span.start, span.end);
	if (edit.fill !== undefined) { element = withProperties(element, propsName, (props) => withFill(props, propsName, edit.fill!)); }
	if (edit.line !== undefined) { element = withProperties(element, propsName, (props) => withLine(props, propsName, edit.line!)); }
	if (edit.rotation !== undefined) {
		element = withProperties(element, propsName, (props) => {
			const xfrm = findElement(props, 'a:xfrm');
			if (!xfrm) {
				throw new ShapeError(`'${info.name}' has no a:xfrm to carry a rotation; the part is not one Word wrote.`);
			}
			return props.slice(0, xfrm.start) + withRotation(props.slice(xfrm.start, xfrm.openEnd), edit.rotation!) + props.slice(xfrm.openEnd);
		});
	}
	if (edit.font !== undefined) {
		if (!/<w:txbxContent\b/.test(element)) {
			throw new ShapeError(`'${info.name}' has no text yet; give it text, and the font with it.`);
		}
		element = withAllTextBoxes(element, (box) => withWordFont(box, edit.font!));
	}
	let out = splice(entryXml, span, element);
	if (shape.inGroup) {
		// Word changed only the DrawingML of a shape in a canvas, and left
		// its VML twin as it was; so does this.
		return out;
	}
	out = withTwin(out, (twin) => {
		let next = twin;
		// Outline before fill: Word writes the v:fill ahead of the v:stroke.
		if (edit.line !== undefined) { next = withVmlLine(next, edit.line); }
		if (edit.fill !== undefined) { next = withVmlFill(next, edit.fill); }
		if (edit.rotation !== undefined) {
			const degrees = checkedRotation(edit.rotation);
			next = withVmlElementStyle(next, (style) => (degrees
				? withStyleProperty(style, 'rotation', vmlDegrees(degrees))
				: withVmlStyle(style, 'rotation', undefined)));
		}
		if (edit.font !== undefined) { next = withAllTextBoxes(next, (box) => withWordFont(box, edit.font!)); }
		return next;
	});
	return edit.rotation !== undefined || edit.line !== undefined ? withEffectExtent(out, shape, edit) : out;
}

/**
 * A VML twin shown or hidden, as Word writes it: visibility in its style,
 * and in a canvas's own background shape too. The named shapes in a canvas
 * keep their own visibility.
 */
function withVmlVisibility(twinXml: string, hidden: boolean): string {
	const set = (style: string): string => {
		if (hidden) { return withStyleProperty(style, 'visibility', 'hidden'); }
		return /(?:^|;)\s*visibility\s*:/.test(style) ? withStyleProperty(style, 'visibility', 'visible') : style;
	};
	let out = withVmlElementStyle(twinXml, set);
	if (!/^<v:group\b[^>]*\beditas="canvas"/.test(out)) { return out; }
	for (const child of children(out, out.indexOf('>') + 1, out.length).reverse()) {
		if (child.name === 'v:shape' && /^_x0000_/.test(attr(out.slice(child.start, child.openEnd), 'id') ?? '')) {
			out = splice(out, child, withVmlElementStyle(out.slice(child.start, child.end), set));
		}
	}
	return out;
}

/**
 * A floating shape's entry at a new stacking height: its wp:anchor
 * relativeHeight, and the z-index of its VML twin, which Word keeps equal
 * to it (negative for a shape behind the text, whose sign is kept).
 */
function withStackHeight(entryXml: string, height: number): string {
	const out = entryXml.replace(/(<wp:anchor\b[^>]*\brelativeHeight=")\d+(")/, `$1${height}$2`);
	return withTwin(out, (twin) => withVmlElementStyle(twin, (style) => {
		const current = /(?:^|;)\s*z-index\s*:\s*(-?)\d+/.exec(style);
		return current ? withStyleProperty(style, 'z-index', `${current[1]}${height}`) : style;
	}));
}

/**
 * The stacking heights a z-order command gives, as [shape, height] pairs;
 * a shape stacks among the story's floating shapes on its side of the text.
 * Word brought a shape to the front by giving it its layer's highest
 * height plus 1024, its own step between new shapes, and changed no other
 * shape. Back is the same at the bottom, and forward and backward trade
 * heights with the neighbor; when there is no room, the layer is numbered
 * afresh in its new order.
 */
function restackHeights(shapes: readonly DocShape[], shape: DocShape, command: ZOrderCommand): Array<[DocShape, number]> {
	const layer = stackOrder(shapes).filter((s) => s.stack!.behind === shape.stack!.behind);
	const index = layer.indexOf(shape);
	const target = restackedIndex(index, layer.length, command);
	if (target === index) { return []; }
	const heights = layer.map((s) => s.stack!.height);
	if (command === 'front') { return [[shape, Math.max(...heights) + 1024]]; }
	if (command === 'back' && Math.min(...heights) >= 1024) { return [[shape, Math.min(...heights) - 1024]]; }
	const neighbor = layer[target];
	if ((command === 'forward' || command === 'backward') && neighbor.stack!.height !== shape.stack!.height) {
		return [[shape, neighbor.stack!.height], [neighbor, shape.stack!.height]];
	}
	const order = layer.filter((s) => s !== shape);
	order.splice(target, 0, shape);
	const base = Math.min(...heights);
	return order.map((s, i): [DocShape, number] => [s, base + i * 1024]).filter(([s, height]) => s.stack!.height !== height);
}

function updateShape(pkg: Package, story: StoryRef, read: { xml: string; shapes: DocShape[] }, shape: DocShape, edit: ShapeEdit): string {
	const { info } = shape;
	let entryXml = read.xml.slice(shape.entry.start, shape.entry.end);
	if (edit.left !== undefined || edit.top !== undefined || edit.width !== undefined || edit.height !== undefined) {
		entryXml = movedShape(entryXml, shape, edit);
	}
	if (edit.text !== undefined) {
		if (info.kind !== 'shape' && info.kind !== 'textBox') {
			throw new ShapeError(`'${info.name}' is a ${info.kind}, which holds no text.`);
		}
		entryXml = withShapeText(entryXml, shape, edit.text);
	}
	entryXml = formattedEntry(entryXml, shape, edit);
	if (edit.zOrder !== undefined) {
		if (shape.inGroup) {
			const container = read.shapes.find((s) => !s.inGroup && s.entry.start === shape.entry.start);
			throw new ShapeError(`'${info.name}' is inside '${container?.info.name ?? 'a canvas'}' and stacks with it; restack '${container?.info.name ?? 'the canvas'}' instead.`);
		}
		if (!shape.stack) {
			throw new ShapeError(`'${info.name}' is inline with the text, which does not stack; make it a floating shape in Word first.`);
		}
	}
	if (edit.altText !== undefined || edit.newName !== undefined || edit.hidden !== undefined) {
		const { nameElement } = locate(entryXml, shape);
		let startTag = entryXml.slice(nameElement.start, nameElement.openEnd);
		if (edit.altText !== undefined) {
			startTag = withAttr(startTag, 'descr', edit.altText || undefined);
		}
		if (edit.newName !== undefined) {
			assertNameFree(read.shapes.filter((s) => s !== shape), edit.newName, story.name);
			startTag = withAttr(startTag, 'name', edit.newName);
		}
		if (edit.hidden !== undefined) {
			startTag = withAttr(startTag, 'hidden', edit.hidden ? '1' : undefined);
		}
		entryXml = entryXml.slice(0, nameElement.start) + startTag + entryXml.slice(nameElement.openEnd);
		if (edit.newName !== undefined && !shape.inGroup) {
			entryXml = renamedVml(entryXml, info.name, edit.newName);
		}
		if (edit.hidden !== undefined && !shape.inGroup) {
			entryXml = withTwin(entryXml, (twin) => withVmlVisibility(twin, edit.hidden!));
		}
	}
	// A restack can move a neighbor too; each entry is written back from the
	// last to the first, so the offsets read before stay good.
	const writes: Array<{ span: Span; xml: string }> = [{ span: shape.entry, xml: entryXml }];
	for (const [moved, height] of edit.zOrder !== undefined ? restackHeights(read.shapes, shape, edit.zOrder) : []) {
		if (moved === shape) {
			writes[0].xml = withStackHeight(writes[0].xml, height);
		} else {
			writes.push({ span: moved.entry, xml: withStackHeight(read.xml.slice(moved.entry.start, moved.entry.end), height) });
		}
	}
	let written = read.xml;
	for (const write of writes.sort((a, b) => b.span.start - a.span.start)) {
		written = splice(written, write.span, write.xml);
	}
	pkg.write(story.path, written);
	return edit.newName ?? info.name;
}

function enclosingRun(xml: string, span: Span): Span | undefined {
	for (let at = 0; ;) {
		const run = findElement(xml, 'w:r', at);
		if (!run || run.start > span.start) { return undefined; }
		if (run.end >= span.end) { return run; }
		at = run.start + 1;
	}
}

/** Whether a run holds the drawing and nothing else that deleting would lose. */
function onlyDrawingIn(xml: string, run: Span, entry: Span): boolean {
	return children(xml, run.openEnd, run.end)
		.every((child) => child.name === 'w:rPr' || (child.start >= entry.start && child.end <= entry.end));
}

function deleteShape(pkg: Package, story: StoryRef, read: { xml: string; shapes: DocShape[] }, shape: DocShape): void {
	if (shape.inGroup) {
		const container = read.shapes.find((s) => !s.inGroup && s.entry.start === shape.entry.start);
		throw new ShapeError(`'${shape.info.name}' is inside '${container?.info.name ?? 'a container'}'; ungroup it in Word, or delete the container.`);
	}
	const run = enclosingRun(read.xml, shape.entry);
	// A run holding nothing but this drawing goes with it; one that also
	// holds text keeps its text.
	const span = run && onlyDrawingIn(read.xml, run, shape.entry) ? run : shape.entry;
	pkg.write(story.path, splice(read.xml, span, ''));
}

/** The style Word 16 gives a new AutoShape: the theme's first accent. */
const AUTOSHAPE_STYLE = '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
	+ '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
	+ '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
	+ '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>';

const BODY_PR = '<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow"'
	+ ' vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0"'
	+ ' rtlCol="0" fromWordArt="0" anchor="ctr" anchorCtr="0" forceAA="0" compatLnSpc="1">'
	+ '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr>';

/** The size a shape is added at when the caller gives none, in points. */
const DEFAULT_SIZE = { width: 120, height: 60 };

/** A drawing id free in this story. Word writes large ids; any free one works. */
function nextDrawingId(shapes: readonly DocShape[]): number {
	return Math.max(0, ...shapes.map((s) => s.id)) + 1;
}

function addShape(pkg: Package, story: StoryRef, read: { xml: string; shapes: DocShape[] }, edit: ShapeEdit): string {
	if (edit.type === 'button') {
		throw new ShapeError('A button is an Excel form control; Word has no shape that runs a macro.');
	}
	if (!edit.type) {
		throw new ShapeError('Adding a shape needs a type: rectangle, roundedRectangle, oval or textBox.');
	}
	if (story.kind !== 'document' && story.kind !== 'header' && story.kind !== 'footer') {
		// A note story holds many notes, and nothing here says which one a
		// shape would belong to. Listing, changing and deleting still work.
		throw new ShapeError(`${story.name} holds one story per note, so there is nowhere here to add a shape; add it in Word, then change it with this tool.`);
	}
	const width = edit.width ?? DEFAULT_SIZE.width;
	const height = edit.height ?? DEFAULT_SIZE.height;
	if (width <= 0 || height <= 0) {
		throw new ShapeError(`A shape needs a width and a height above zero; ${width} by ${height} points is not one.`);
	}
	const textBox = edit.type === 'textBox';
	const label = textBox ? 'Text Box' : PRESET_LABELS[edit.type as PresetShapeType];
	const prst = textBox ? 'rect' : PRESET_GEOMETRY[edit.type as PresetShapeType];
	const id = nextDrawingId(read.shapes);
	const name = edit.name ?? `${label} ${read.shapes.filter((s) => !s.inGroup).length + 1}`;
	assertNameFree(read.shapes, name, story.name);
	const descr = edit.altText ? ` descr="${encodeXml(edit.altText)}"` : '';
	const floating = edit.left !== undefined || edit.top !== undefined;
	const text = edit.text ?? '';
	const txbx = textBox || text ? `<wps:txbx>${withWordText(textBox ? '<w:txbxContent>' : SHAPE_TEXT, text)}</wps:txbx>` : '';
	const outline = textBox
		? '<a:noFill/><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>'
		: '';
	const wsp = `<wps:wsp><wps:cNvSpPr${textBox ? ' txBox="1"' : ''}/>`
		+ `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/></a:xfrm>`
		+ `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${outline}</wps:spPr>`
		+ `${textBox ? '' : AUTOSHAPE_STYLE}${txbx}${BODY_PR}</wps:wsp>`;
	const graphic = '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
		+ `<a:graphicData uri="${GRAPHIC.shape}">${wsp}</a:graphicData></a:graphic>`;
	const docPr = `<wp:docPr id="${id}" name="${encodeXml(name)}"${descr}/><wp:cNvGraphicFramePr/>`;
	// On top of the story's other floating shapes, a step of 1024 above the
	// highest, as Word stacked the fixture's shapes it drew one after another.
	const heights = read.shapes.filter((s) => s.stack && !s.stack.behind).map((s) => s.stack!.height);
	const stackHeight = heights.length > 0 ? Math.max(...heights) + 1024 : 251659264;
	// A bare w:drawing, with no mc:AlternateContent: the VML fallback exists
	// only for Word 2007, and a twin this writer could not keep truthful is
	// worse than none. Word 2010 and later read the drawing itself.
	const drawing = floating
		? '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"'
			+ ` relativeHeight="${stackHeight}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
			+ '<wp:simplePos x="0" y="0"/>'
			+ `<wp:positionH relativeFrom="column"><wp:posOffset>${pointsToEmu(edit.left ?? 0)}</wp:posOffset></wp:positionH>`
			+ `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${pointsToEmu(edit.top ?? 0)}</wp:posOffset></wp:positionV>`
			+ `<wp:extent cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/>`
			+ `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>${docPr}${graphic}</wp:anchor></w:drawing>`
		: '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
			+ `<wp:extent cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/>`
			+ `<wp:effectExtent l="0" t="0" r="0" b="0"/>${docPr}${graphic}</wp:inline></w:drawing>`;
	pkg.write(story.path, placedRun(read.xml, story, `<w:r><w:rPr><w:noProof/></w:rPr>${drawing}</w:r>`, floating));
	return name;
}

/**
 * The element that holds a story's paragraphs. Only the document wraps them
 * in a w:body; a header's root is w:hdr and a footer's is w:ftr, with the
 * paragraphs directly inside.
 */
function storyBody(xml: string, story: StoryRef): Span & { name: string } {
	const rootName = story.kind === 'document' ? 'w:body' : story.kind === 'header' ? 'w:hdr' : 'w:ftr';
	const root = findElement(xml, rootName);
	if (!root) {
		throw new ShapeError(`The part has no ${rootName}; it is not a surface Word wrote.`);
	}
	return { ...root, name: rootName };
}

/**
 * Where a new shape's run goes. A floating shape's anchor takes no room, so
 * it joins the last paragraph and nothing on the page moves; an inline
 * shape is part of the text and gets a paragraph of its own.
 */
function placedRun(xml: string, story: StoryRef, run: string, floating: boolean): string {
	const body = storyBody(xml, story);
	const bodyChildren = children(xml, body.openEnd, body.end);
	const paragraphs = bodyChildren.filter((child) => child.name === 'w:p');
	const last = paragraphs[paragraphs.length - 1];
	if (floating && last) {
		const closing = last.end - '</w:p>'.length;
		return splice(xml, { start: closing, end: closing }, run);
	}
	// Before any trailing w:sectPr, which has to stay last in the body.
	const sectPr = bodyChildren.find((child) => child.name === 'w:sectPr');
	const at = sectPr ? sectPr.start : body.end - `</${body.name}>`.length;
	return splice(xml, { start: at, end: at }, `<w:p>${run}</w:p>`);
}

/** Add, change or remove one shape on a surface; gives the shape's name after the edit. */
export function editStoryShape(zip: ZipArchive, story: StoryRef, edit: ShapeEdit): string {
	const pkg = new Package(zip);
	const read = readStory(pkg, story);
	if (edit.macro !== undefined) {
		throw new ShapeError(WORD_SHAPE_MACRO_REFUSAL);
	}
	if (edit.range !== undefined) {
		throw new ShapeError('A document has no cells; place a shape with left, top, width and height, in points.');
	}
	if (edit.linkedCell !== undefined || edit.inputRange !== undefined) {
		throw new ShapeError('linkedCell and inputRange are Excel form-control properties; a document has neither.');
	}
	if (edit.action === 'add') {
		const added = addShape(pkg, story, read, edit);
		// How the new shape looks is set on it once it is there, the same way
		// as on any other shape.
		const look = lookOf(edit);
		return look ? editStoryShape(zip, story, { action: 'update', name: added, ...look }) : added;
	}
	if (!edit.name) {
		throw new ShapeError(`A shape to ${edit.action} needs a name; call the list tool for the names on ${story.name}.`);
	}
	const shape = findTarget(read.shapes, edit.name, story.name);
	if (edit.action === 'delete') {
		deleteShape(pkg, story, read, shape);
		return shape.info.name;
	}
	return updateShape(pkg, story, read, shape, edit);
}
