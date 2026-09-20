// Slide shapes: what agents can see and change on a PowerPoint slide.
//
// A slide is one part (ppt/slides/slideN.xml) and its shapes are the
// children of its p:spTree, positioned by the a:xfrm inside each - there is
// no anchor grid as on a worksheet, so a shape's place is EMU from the
// slide's top-left corner.
//
// A shape runs a macro through an a:hlinkClick inside its p:cNvPr, whose
// action is the URI ppaction://macro?name=Proc. That is a child element
// carrying a URI, where Excel uses a `macro` attribute, and its r:id is
// deliberately empty: the hyperlink relationship a normal link would need
// does not exist for a macro. PowerPoint accepts a bare Sub name as well as
// Module.Sub, and writes back whatever it was given.
//
// Slide ORDER is p:sldIdLst in ppt/presentation.xml, not the slideN file
// numbering, which does not renumber when slides are reordered or deleted.
//
// Every layout here was read from files PowerPoint 16 saved, and the writer
// is checked by having PowerPoint open and re-save the result.

import { ZipArchive } from './zip';
import {
	Package,
	attr,
	children,
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
	drawingTextOf,
	withDrawingText,
	type ShapeEdit,
	type ShapeInfo,
	type ShapeKind,
	type PresetShapeType,
} from './shapes';

const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const PRESENTATION = 'ppt/presentation.xml';

/** The URI scheme PowerPoint runs a macro through. */
const MACRO_ACTION = 'ppaction://macro?name=';

/** One slide, in presentation order. */
export interface SlideRef {
	/** 1-based position in p:sldIdLst. */
	index: number;
	/** The slide's own name when it has one, else `Slide N`. */
	name: string;
	path: string;
}

/**
 * The slides of a presentation, in the order p:sldIdLst gives. A slide
 * carries a name only once someone sets one; otherwise it is addressed by
 * position, the way the thumbnail pane shows it.
 */
export function presentationSlides(zip: ZipArchive): SlideRef[] {
	const pkg = new Package(zip);
	if (!pkg.has(PRESENTATION)) { return []; }
	const xml = pkg.read(PRESENTATION);
	const list = findElement(xml, 'p:sldIdLst');
	if (!list) { return []; }
	const rels = new Map(pkg.relationships(PRESENTATION).filter((rel) => rel.type === SLIDE_REL).map((rel) => [rel.id, rel.path]));
	const out: SlideRef[] = [];
	for (const [tag] of xml.slice(list.openEnd, list.end).matchAll(/<p:sldId\b[^>]*>/g)) {
		const path = rels.get(attr(tag, 'r:id') ?? '');
		if (!path || !pkg.has(path)) { continue; }
		const index = out.length + 1;
		out.push({ index, name: slideName(pkg.read(path)) ?? `Slide ${index}`, path });
	}
	return out;
}

function slideName(xml: string): string | undefined {
	const tag = /<p:cSld\b[^>]*>/.exec(xml)?.[0];
	return tag ? attr(tag, 'name') || undefined : undefined;
}

/**
 * The slide a caller named. A name someone gave a slide wins over the
 * positional spelling, so a deck that names its slides is addressed the way
 * its author addresses it. `Slide 3`, `Slide3` and `3` all mean the third.
 */
export function requireSlide(slides: readonly SlideRef[], surface: string): SlideRef {
	const wanted = surface.trim();
	const lower = wanted.toLowerCase();
	const named = slides.filter((slide) => slide.name.toLowerCase() === lower);
	if (named.length === 1) { return named[0]; }
	if (named.length > 1) {
		throw new ShapeError(`${named.length} slides are named '${surface}'; use its position instead, as 'Slide ${named[0].index}'.`);
	}
	const position = /^(?:slide\s*)?(\d+)$/i.exec(lower);
	if (position) {
		const slide = slides[Number(position[1]) - 1];
		if (slide) { return slide; }
		throw new ShapeError(`The presentation has ${slides.length} slide${slides.length === 1 ? '' : 's'}; there is no slide ${position[1]}.`);
	}
	throw new ShapeError(`No slide named '${surface}'. The presentation has ${slides.map((s) => `'${s.name}'`).join(', ') || 'no slides'}.`);
}

// ------------------------------------------------------------------ reading

/** A shape found on a slide, with where its markup sits in the part. */
interface SlideShape {
	info: ShapeInfo;
	element: Span & { name: string };
	id: number;
	inGroup: boolean;
}

const SHAPE_ELEMENTS = ['p:sp', 'p:grpSp', 'p:graphicFrame', 'p:cxnSp', 'p:pic'];

function shapeKind(element: string, inner: string): ShapeKind {
	switch (element) {
		case 'p:sp':
			if (/<p:nvPr\b[^>]*>[\s\S]*?<p:ph\b/.test(inner)) { return 'placeholder'; }
			return /<p:cNvSpPr\b[^>]*\btxBox="1"/.test(inner) ? 'textBox' : 'shape';
		case 'p:cxnSp': return 'line';
		case 'p:pic': return 'picture';
		case 'p:grpSp': return 'group';
		case 'p:graphicFrame':
			if (/drawingml\/2006\/chart"/.test(inner)) { return 'chart'; }
			return /drawingml\/2006\/table"/.test(inner) ? 'table' : 'other';
		default: return 'other';
	}
}

/** The macro a shape's hyperlink runs, or undefined when it runs something else. */
function macroOf(cNvPrInner: string): string | undefined {
	const tag = /<a:hlinkClick\b[^>]*>/.exec(cNvPrInner)?.[0];
	const action = tag ? attr(tag, 'action') : undefined;
	if (!action?.startsWith(MACRO_ACTION)) { return undefined; }
	return decodeURIComponent(action.slice(MACRO_ACTION.length)) || undefined;
}

function readShape(
	xml: string,
	element: Span & { name: string },
	inGroup: boolean,
	inherited?: ReadonlyMap<string, Partial<ShapeInfo>>,
): SlideShape[] {
	const inner = xml.slice(element.start, element.end);
	const cNvPr = findElement(xml, 'p:cNvPr', element.openEnd, element.end);
	if (!cNvPr) { return []; }
	const cNvPrTag = xml.slice(cNvPr.start, cNvPr.openEnd);
	const kind = shapeKind(element.name, inner);
	const info: ShapeInfo = { name: attr(cNvPrTag, 'name') ?? '', kind };
	const geometry = /<a:prstGeom\b[^>]*\bprst="([^"]*)"/.exec(inner)?.[1];
	if (geometry && (kind === 'shape' || kind === 'placeholder' || kind === 'line')) { info.geometry = geometry; }
	const box = boxOf(xml, element);
	if (kind === 'placeholder' && box.width === undefined) {
		// PowerPoint reports the box this placeholder inherits, so a reader
		// that stopped at the slide would report none at all.
		const key = placeholderKey(inner);
		Object.assign(info, (key !== undefined && inherited?.get(key)) || box);
	} else {
		Object.assign(info, box);
	}
	const macro = macroOf(xml.slice(cNvPr.openEnd, cNvPr.end));
	if (macro) { info.macro = macro; }
	if (kind === 'shape' || kind === 'textBox' || kind === 'placeholder') {
		const text = drawingTextOf(xml, element, 'p:txBody');
		if (text) { info.text = text; }
	}
	const descr = attr(cNvPrTag, 'descr');
	if (descr) { info.altText = descr; }
	if (attr(cNvPrTag, 'hidden') === '1') { info.hidden = true; }
	const out: SlideShape[] = [{ info, element, id: Number(attr(cNvPrTag, 'id') ?? 0), inGroup }];
	if (element.name === 'p:grpSp') {
		info.shapes = [];
		for (const child of children(xml, element.openEnd, element.end)) {
			if (!SHAPE_ELEMENTS.includes(child.name)) { continue; }
			const members = readShape(xml, child, true);
			if (members[0]) { info.shapes.push(members[0].info); }
			out.push(...members);
		}
	}
	return out;
}

/**
 * Which child of a shape element holds its a:xfrm. A group's is p:grpSpPr
 * and a graphic frame keeps a p:xfrm of its own; everything else uses
 * p:spPr.
 */
function propsNameOf(element: string): string {
	if (element === 'p:grpSp') { return 'p:grpSpPr'; }
	return element === 'p:graphicFrame' ? 'p:xfrm' : 'p:spPr';
}

/**
 * A shape's position and size in points. The properties element is looked
 * up among the element's DIRECT children: a group's members carry a p:spPr
 * each, and a search into the subtree would report the first member's box
 * as the group's.
 */
function boxOf(xml: string, element: Span & { name: string }): Partial<ShapeInfo> {
	const wanted = propsNameOf(element.name);
	const props = children(xml, element.openEnd, element.end).find((child) => child.name === wanted);
	if (!props) { return {}; }
	// A graphic frame's p:xfrm is the transform itself, not a wrapper.
	const xfrm = wanted === 'p:xfrm' ? props : findElement(xml, 'a:xfrm', props.start, props.end);
	return xfrm ? xfrmBox(xml, xfrm) : {};
}

/**
 * What a placeholder is, as PowerPoint matches one to its layout: the type
 * it carries, and the index that tells two of a type apart. A p:ph with no
 * type is a body placeholder, which is the schema's default.
 */
function placeholderKey(markup: string): string | undefined {
	const tag = /<p:ph\b[^>]*>/.exec(markup)?.[0];
	if (!tag) { return undefined; }
	return `${attr(tag, 'type') ?? 'body'}:${attr(tag, 'idx') ?? ''}`;
}

/**
 * Where a slide's layout, and behind it the master, put each placeholder.
 *
 * A placeholder on a slide usually carries no a:xfrm of its own, and
 * PowerPoint reports the box it inherits. Reading only the slide reports no
 * position at all for the title of every deck that is not blank-layout.
 * The master is read first and the layout written over it, which is the
 * order the inheritance runs in.
 */
function inheritedBoxes(pkg: Package, slidePath: string): Map<string, Partial<ShapeInfo>> {
	const out = new Map<string, Partial<ShapeInfo>>();
	const layout = pkg.relationships(slidePath).find((rel) => rel.type === LAYOUT_REL)?.path;
	if (!layout || !pkg.has(layout)) { return out; }
	const master = pkg.relationships(layout).find((rel) => rel.type === MASTER_REL)?.path;
	for (const part of [master, layout]) {
		if (!part || !pkg.has(part)) { continue; }
		const xml = pkg.read(part);
		const tree = findElement(xml, 'p:spTree');
		if (!tree) { continue; }
		for (const child of children(xml, tree.openEnd, tree.end)) {
			if (!SHAPE_ELEMENTS.includes(child.name)) { continue; }
			const key = placeholderKey(xml.slice(child.start, child.end));
			if (key === undefined) { continue; }
			const box = boxOf(xml, child);
			if (box.width || box.height) { out.set(key, box); }
		}
	}
	return out;
}

function readSlide(pkg: Package, slide: SlideRef): { xml: string; tree: Span; shapes: SlideShape[] } {
	const xml = pkg.read(slide.path);
	const tree = findElement(xml, 'p:spTree');
	if (!tree) {
		throw new ShapeError(`${slide.name} has no shape tree; the slide part is not one PowerPoint wrote.`);
	}
	const inherited = inheritedBoxes(pkg, slide.path);
	const shapes: SlideShape[] = [];
	for (const child of children(xml, tree.openEnd, tree.end)) {
		if (!SHAPE_ELEMENTS.includes(child.name)) { continue; }
		shapes.push(...readShape(xml, child, false, inherited));
	}
	return { xml, tree, shapes };
}

/** Every shape on a slide, top level only; a group lists its members. */
export function listSlideShapes(zip: ZipArchive, slide: SlideRef): ShapeInfo[] {
	return readSlide(new Package(zip), slide).shapes.filter((s) => !s.inGroup).map((s) => s.info);
}

// ------------------------------------------------------------------ editing

/** The schema order of p:cNvPr's children; an out-of-order child fails to load. */
const CNVPR_ORDER = ['a:hlinkClick', 'a:hlinkHover', 'a:extLst'];

/** The schema order of p:spPr's children, as far as this writer touches them. */
const SPPR_ORDER = ['a:xfrm', 'a:custGeom', 'a:prstGeom'];

/**
 * The size a shape is added at when the caller gives none, in points: a
 * default this tool picks, not one read from PowerPoint.
 */
const DEFAULT_SIZE = { width: 120, height: 60 };

function findTarget(shapes: readonly SlideShape[], name: string, slide: string): SlideShape {
	const lower = name.toLowerCase();
	const matches = shapes.filter((s) => s.info.name.toLowerCase() === lower);
	if (matches.length === 0) {
		throw new ShapeError(`No shape named '${name}' on ${slide}.`);
	}
	if (matches.length > 1) {
		throw new ShapeError(`${matches.length} shapes on ${slide} are named '${name}'; rename one in PowerPoint first.`);
	}
	return matches[0];
}

function assertNameFree(shapes: readonly SlideShape[], name: string, slide: string): void {
	if (shapes.some((s) => s.info.name.toLowerCase() === name.toLowerCase())) {
		throw new ShapeError(`${slide} already has a shape named '${name}'.`);
	}
}

/**
 * `cNvPr` with a child set or removed, expanding a self-closing tag when it
 * gains its first child and collapsing it again when it loses its last.
 */
function withCNvPrChild(cNvPrXml: string, name: string, replacement: string | undefined): string {
	const selfClosing = cNvPrXml.endsWith('/>') && !cNvPrXml.includes(`</p:cNvPr>`);
	const openEnd = cNvPrXml.indexOf('>') + 1;
	const startTag = selfClosing ? `${cNvPrXml.slice(0, openEnd - 2)}>` : cNvPrXml.slice(0, openEnd);
	let inner = selfClosing ? '' : cNvPrXml.slice(openEnd, cNvPrXml.lastIndexOf('</p:cNvPr>'));
	const existing = findElement(inner, name);
	if (existing) {
		inner = splice(inner, existing, replacement ?? '');
	} else if (replacement) {
		const rank = (child: string): number => {
			const at = CNVPR_ORDER.indexOf(child);
			return at < 0 ? CNVPR_ORDER.length : at;
		};
		const mine = rank(name);
		const after = children(inner, 0, inner.length).find((child) => rank(child.name) > mine);
		const at = after?.start ?? inner.length;
		inner = inner.slice(0, at) + replacement + inner.slice(at);
	}
	return inner ? `${startTag}${inner}</p:cNvPr>` : `${startTag.slice(0, -1)}/>`;
}

/** The a:xfrm a shape is given when it has none: PowerPoint writes off then ext. */
function xfrmXml(left: number, top: number, width: number, height: number): string {
	return `<a:xfrm><a:off x="${pointsToEmu(left)}" y="${pointsToEmu(top)}"/>`
		+ `<a:ext cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/></a:xfrm>`;
}

function movedShape(elementXml: string, shape: SlideShape, edit: ShapeEdit): string {
	const left = edit.left ?? shape.info.left ?? 0;
	const top = edit.top ?? shape.info.top ?? 0;
	const width = edit.width ?? shape.info.width ?? 0;
	const height = edit.height ?? shape.info.height ?? 0;
	if (width <= 0 || height <= 0) {
		throw new ShapeError(`A shape needs a width and a height above zero; '${shape.info.name}' would be ${width} by ${height} points.`);
	}
	const propsName = propsNameOf(shape.element.name);
	const open = elementXml.indexOf('>') + 1;
	const props = children(elementXml, open, elementXml.length).find((child) => child.name === propsName);
	if (!props) {
		throw new ShapeError(`'${shape.info.name}' has no ${propsName}; the slide part is not one PowerPoint wrote.`);
	}
	const xfrm = findElement(elementXml, 'a:xfrm', props.start, props.end);
	if (xfrm) {
		// Keep any rotation or flip the shape already carries.
		const old = elementXml.slice(xfrm.start, xfrm.end);
		const body = old
			.replace(/<a:off\b[^>]*\/>/, `<a:off x="${pointsToEmu(left)}" y="${pointsToEmu(top)}"/>`)
			.replace(/<a:ext\b[^>]*\/>/, `<a:ext cx="${pointsToEmu(width)}" cy="${pointsToEmu(height)}"/>`);
		return splice(elementXml, xfrm, body);
	}
	if (props.openEnd === props.end) {
		// A self-closing <p:spPr/>: give it a body holding just the xfrm.
		return splice(elementXml, props, `<${propsName}>${xfrmXml(left, top, width, height)}</${propsName}>`);
	}
	const rank = (child: string): number => {
		const at = SPPR_ORDER.indexOf(child);
		return at < 0 ? SPPR_ORDER.length : at;
	};
	const after = children(elementXml, props.openEnd, props.end).find((child) => rank(child.name) > rank('a:xfrm'));
	const at = after?.start ?? props.end - `</${propsName}>`.length;
	return elementXml.slice(0, at) + xfrmXml(left, top, width, height) + elementXml.slice(at);
}

function updateShape(pkg: Package, slide: SlideRef, read: ReturnType<typeof readSlide>, shape: SlideShape, edit: ShapeEdit): string {
	let elementXml = read.xml.slice(shape.element.start, shape.element.end);
	if (edit.text !== undefined) {
		const body = findElement(elementXml, 'p:txBody');
		if (!body || (shape.info.kind !== 'shape' && shape.info.kind !== 'textBox' && shape.info.kind !== 'placeholder')) {
			throw new ShapeError(`'${shape.info.name}' is a ${shape.info.kind}, which holds no text.`);
		}
		const old = elementXml.slice(body.start, body.end);
		elementXml = splice(elementXml, body, withDrawingText(old, edit.text, 'p:txBody'));
	}
	if (edit.left !== undefined || edit.top !== undefined || edit.width !== undefined || edit.height !== undefined) {
		elementXml = movedShape(elementXml, shape, edit);
	}
	// The cNvPr edits come last: the splices above move its offsets.
	const cNvPr = findElement(elementXml, 'p:cNvPr');
	if (!cNvPr) {
		throw new ShapeError(`'${shape.info.name}' has no p:cNvPr; the slide part is not one PowerPoint wrote.`);
	}
	let cNvPrXml = elementXml.slice(cNvPr.start, cNvPr.end);
	if (edit.macro !== undefined) {
		if (shape.info.kind === 'group') {
			throw new ShapeError(`'${shape.info.name}' is a group, which PowerPoint runs no macro for; assign the macro to a shape in it.`);
		}
		cNvPrXml = withCNvPrChild(cNvPrXml, 'a:hlinkClick', edit.macro
			? `<a:hlinkClick r:id="" action="${encodeXml(MACRO_ACTION + encodeURIComponent(edit.macro))}"/>`
			: undefined);
	}
	if (edit.altText !== undefined || edit.newName !== undefined) {
		const openEnd = cNvPrXml.indexOf('>') + 1;
		let startTag = cNvPrXml.slice(0, openEnd);
		if (edit.altText !== undefined) {
			startTag = withAttr(startTag, 'descr', edit.altText || undefined);
		}
		if (edit.newName !== undefined) {
			assertNameFree(read.shapes.filter((s) => s !== shape), edit.newName, slide.name);
			startTag = withAttr(startTag, 'name', edit.newName);
		}
		cNvPrXml = startTag + cNvPrXml.slice(openEnd);
	}
	elementXml = splice(elementXml, cNvPr, cNvPrXml);
	pkg.write(slide.path, splice(read.xml, shape.element, elementXml));
	return edit.newName ?? shape.info.name;
}

/** The style PowerPoint 16 gives a new AutoShape: the theme's first accent. */
const AUTOSHAPE_STYLE = '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
	+ '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
	+ '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
	+ '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

function addShape(pkg: Package, slide: SlideRef, read: ReturnType<typeof readSlide>, edit: ShapeEdit): string {
	if (edit.type === 'button') {
		throw new ShapeError('A button is an Excel form control; on a slide, add a shape and give it a macro.');
	}
	if (!edit.type) {
		throw new ShapeError('Adding a shape needs a type: rectangle, roundedRectangle, oval or textBox.');
	}
	if (edit.left === undefined || edit.top === undefined) {
		throw new ShapeError('Adding a shape to a slide needs left and top, in points from the slide\'s top-left corner.');
	}
	const width = edit.width ?? DEFAULT_SIZE.width;
	const height = edit.height ?? DEFAULT_SIZE.height;
	if (width <= 0 || height <= 0) {
		throw new ShapeError(`A shape needs a width and a height above zero; ${width} by ${height} points is not one.`);
	}
	const textBox = edit.type === 'textBox';
	const id = nextShapeId(read.shapes);
	const label = textBox ? 'TextBox' : PRESET_LABELS[edit.type as PresetShapeType];
	const prst = textBox ? 'rect' : PRESET_GEOMETRY[edit.type as PresetShapeType];
	const name = edit.name ?? `${label} ${id - 1}`;
	assertNameFree(read.shapes, name, slide.name);
	const descr = edit.altText ? ` descr="${encodeXml(edit.altText)}"` : '';
	const link = edit.macro
		? `<a:hlinkClick r:id="" action="${encodeXml(MACRO_ACTION + encodeURIComponent(edit.macro))}"/>`
		: '';
	const cNvPr = link
		? `<p:cNvPr id="${id}" name="${encodeXml(name)}"${descr}>${link}</p:cNvPr>`
		: `<p:cNvPr id="${id}" name="${encodeXml(name)}"${descr}/>`;
	const body = withDrawingText(
		textBox
			? '<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/></p:txBody>'
			: '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/></p:txBody>',
		edit.text ?? '',
		'p:txBody',
	);
	const shape = `<p:sp><p:nvSpPr>${cNvPr}<p:cNvSpPr${textBox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>`
		+ `<p:spPr>${xfrmXml(edit.left, edit.top, width, height)}`
		+ `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${textBox ? '<a:noFill/>' : ''}</p:spPr>`
		+ `${textBox ? '' : AUTOSHAPE_STYLE}${body}</p:sp>`;
	// Last in the tree is topmost, which is where PowerPoint puts a new shape.
	const closing = read.xml.lastIndexOf('</p:spTree>', read.tree.end);
	pkg.write(slide.path, splice(read.xml, { start: closing, end: closing }, shape));
	return name;
}

/** A drawing id free on this slide. Ids start at 2; the tree itself is 1. */
function nextShapeId(shapes: readonly SlideShape[]): number {
	return Math.max(1, ...shapes.map((s) => s.id)) + 1;
}

function deleteShape(pkg: Package, slide: SlideRef, read: ReturnType<typeof readSlide>, shape: SlideShape): void {
	if (shape.inGroup) {
		throw new ShapeError(`'${shape.info.name}' is inside a group; ungroup it in PowerPoint, or delete the group.`);
	}
	pkg.write(slide.path, splice(read.xml, shape.element, ''));
}

/** Add, change or remove one shape on a slide; gives the shape's name after the edit. */
export function editSlideShape(zip: ZipArchive, slide: SlideRef, edit: ShapeEdit): string {
	const pkg = new Package(zip);
	const read = readSlide(pkg, slide);
	if (edit.action === 'add') {
		return addShape(pkg, slide, read, edit);
	}
	if (!edit.name) {
		throw new ShapeError(`A shape to ${edit.action} needs a name; call the list tool for the names on ${slide.name}.`);
	}
	const shape = findTarget(read.shapes, edit.name, slide.name);
	if (edit.action === 'delete') {
		deleteShape(pkg, slide, read, shape);
		return shape.info.name;
	}
	if (edit.range !== undefined) {
		throw new ShapeError('A slide has no cells; place a shape with left, top, width and height, in points.');
	}
	if (edit.linkedCell !== undefined || edit.inputRange !== undefined) {
		throw new ShapeError('linkedCell and inputRange are Excel form-control properties; a slide has neither.');
	}
	return updateShape(pkg, slide, read, shape, edit);
}
