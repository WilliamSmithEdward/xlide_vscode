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
	type ShapeEdit,
	type ShapeInfo,
	type ShapeKind,
	type PresetShapeType,
} from './shapes';

const DOCUMENT = 'word/document.xml';

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
function readMembers(xml: string, container: Span, entry: Span): { direct: DocShape[]; all: DocShape[] } {
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
		const found: DocShape = { info, entry, element: child, id: Number(attr(tag, 'id') ?? 0), inGroup: true };
		direct.push(found);
		all.push(found);
		if (child.name === 'wpg:wgp') {
			const nested = readMembers(xml, child, entry);
			info.shapes = nested.direct.map((member) => member.info);
			all.push(...nested.all);
		}
	}
	return { direct, all };
}

function readTopLevel(xml: string, entry: Span, anchor: Span & { name: string }): DocShape[] {
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
	const found: DocShape = { info, entry, element, id: Number(attr(tag, 'id') ?? 0), inGroup: false };
	const out = [found];
	if (kind === 'canvas' || kind === 'group') {
		const members = readMembers(xml, element, entry);
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

/** Every drawing in a story, with the VML twins inside mc:Fallback skipped. */
function readStory(pkg: Package, story: StoryRef): { xml: string; shapes: DocShape[] } {
	const xml = pkg.read(story.path);
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
		shapes.push(...readTopLevel(xml, enclosingAlternateContent(xml, drawing) ?? drawing, anchor));
	}
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

/** The VML twin's geometry brought in line with the DrawingML one. */
function movedVml(xml: string, box: { left?: number; top?: number; width: number; height: number }): string {
	const shapes = /(<v:(?:rect|oval|roundrect|shape|group)\b[^>]*\bstyle=")([^"]*)(")/g;
	return xml.replace(shapes, (_m, head: string, style: string, tail: string) => {
		let next = withStyleProperty(style, 'width', `${box.width}pt`);
		next = withStyleProperty(next, 'height', `${box.height}pt`);
		if (box.left !== undefined) { next = withStyleProperty(next, 'margin-left', `${box.left}pt`); }
		if (box.top !== undefined) { next = withStyleProperty(next, 'margin-top', `${box.top}pt`); }
		return head + next + tail;
	});
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

function updateShape(pkg: Package, story: StoryRef, read: { xml: string; shapes: DocShape[] }, shape: DocShape, edit: ShapeEdit): string {
	let entryXml = read.xml.slice(shape.entry.start, shape.entry.end);
	if (edit.left !== undefined || edit.top !== undefined || edit.width !== undefined || edit.height !== undefined) {
		entryXml = movedShape(entryXml, shape, edit);
	}
	if (edit.text !== undefined) {
		if (shape.info.kind !== 'shape' && shape.info.kind !== 'textBox') {
			throw new ShapeError(`'${shape.info.name}' is a ${shape.info.kind}, which holds no text.`);
		}
		const within = shape.inGroup
			? locate(entryXml, shape).element
			: { start: 0, openEnd: 0, end: entryXml.length };
		const result = setAllText(entryXml, within, edit.text);
		if (!result.found) {
			throw new ShapeError(`'${shape.info.name}' has no text box; give it text in Word first.`);
		}
		entryXml = result.xml;
	}
	if (edit.altText !== undefined || edit.newName !== undefined) {
		const { nameElement } = locate(entryXml, shape);
		let startTag = entryXml.slice(nameElement.start, nameElement.openEnd);
		if (edit.altText !== undefined) {
			startTag = withAttr(startTag, 'descr', edit.altText || undefined);
		}
		if (edit.newName !== undefined) {
			assertNameFree(read.shapes.filter((s) => s !== shape), edit.newName, story.name);
			startTag = withAttr(startTag, 'name', edit.newName);
		}
		entryXml = entryXml.slice(0, nameElement.start) + startTag + entryXml.slice(nameElement.openEnd);
		if (edit.newName !== undefined && !shape.inGroup) {
			entryXml = renamedVml(entryXml, shape.info.name, edit.newName);
		}
	}
	pkg.write(story.path, splice(read.xml, shape.entry, entryXml));
	return edit.newName ?? shape.info.name;
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
	const txbx = textBox || text ? `<wps:txbx>${withWordText('<w:txbxContent>', text)}</wps:txbx>` : '';
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
	// A bare w:drawing, with no mc:AlternateContent: the VML fallback exists
	// only for Word 2007, and a twin this writer could not keep truthful is
	// worse than none. Word 2010 and later read the drawing itself.
	const drawing = floating
		? '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"'
			+ ' relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
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
		return addShape(pkg, story, read, edit);
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
