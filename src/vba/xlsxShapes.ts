// Worksheet shapes: what agents can see and change on a sheet's drawing layer.
//
// A sheet's shapes live in its drawing part (xl/drawings/drawingN.xml,
// DrawingML), each in a cell anchor that Excel positions it by - the a:xfrm
// inside is ignored. A shape's OnAction is its `macro` attribute. Form
// controls are older, and Excel keeps each in four places that have to
// agree: a VML shape (xl/drawings/vmlDrawingN.vml, which cell notes share),
// an entry in the sheet's <controls>, a ctrlProp part, and a hidden DrawingML
// twin with the same id in the drawing part; Excel 2010 and later place a
// control by the anchors of the last two. ActiveX controls are listed, not
// edited: they run event procedures in the sheet's module, not a macro.
// Every layout here was read from files Excel 16 saved, and each edit was
// checked by having Excel open the result.

import { ZipArchive } from './zip';
import { columnToIndex, indexToColumn } from './xlsxFormula';
import {
	Package,
	attr,
	children,
	decodeXml,
	encodeXml,
	findElement,
	splice,
	withAttr,
	type Relationship,
	type Span,
} from './ooxml';
import {
	PRESET_GEOMETRY,
	PRESET_LABELS,
	ShapeError,
	drawingTextOf,
	withDrawingText as withDrawingTextBody,
	type ShapeEdit,
	type ShapeInfo,
	type ShapeKind,
	type PresetShapeType,
} from './shapes';

export { ShapeError } from './shapes';
export type { NewShapeType, ShapeEdit, ShapeInfo, ShapeKind } from './shapes';

const REL = {
	drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
	vmlDrawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
	ctrlProp: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
};
const CONTENT_TYPE = {
	drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
	vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
	ctrlProp: 'application/vnd.ms-excel.controlproperties+xml',
};
const NS = {
	xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
};
// ------------------------------------------------------------------- anchors

interface Marker {
	col: number;
	row: number;
	colOff: number;
	rowOff: number;
}

function readMarker(xml: string, name: string, prefix: string): Marker | undefined {
	const span = findElement(xml, `${prefix}${name}`);
	if (!span) { return undefined; }
	const inner = xml.slice(span.openEnd, span.end);
	const value = (field: string): number => Number(new RegExp(`<xdr:${field}>(-?\\d+)</xdr:${field}>`).exec(inner)?.[1] ?? 0);
	return { col: value('col'), row: value('row'), colOff: value('colOff'), rowOff: value('rowOff') };
}

function markerCell(marker: Marker): string {
	return `${indexToColumn(marker.col + 1)}${marker.row + 1}`;
}

function markerXml(name: string, prefix: string, col: number, row: number): string {
	return `<${prefix}${name}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row>`
		+ `<xdr:rowOff>0</xdr:rowOff></${prefix}${name}>`;
}

/** Cells as 0-based first and last column and row. */
interface CellBox {
	c1: number;
	r1: number;
	c2: number;
	r2: number;
}

function parseRange(range: string): CellBox {
	const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(range);
	if (!m) {
		throw new ShapeError(`'${range}' is not a range of cells; give one such as B2:D4.`);
	}
	const c1 = columnToIndex(m[1]) - 1;
	const r1 = Number(m[2]) - 1;
	const c2 = m[3] ? columnToIndex(m[3]) - 1 : c1;
	const r2 = m[4] ? Number(m[4]) - 1 : r1;
	const box = { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
	if (box.c2 >= 16384 || box.r2 >= 1048576 || box.r1 < 0) {
		throw new ShapeError(`'${range}' runs past XFD1048576, the last cell of a worksheet.`);
	}
	return box;
}

/**
 * The from and to markers covering a box of cells: from its first cell's
 * top-left corner to the top-left corner past its last.
 */
function anchorMarkers(box: CellBox, prefix: string): string {
	return markerXml('from', prefix, box.c1, box.r1) + markerXml('to', prefix, box.c2 + 1, box.r2 + 1);
}

/** A sheet name as a formula writes it, quoted unless it could not be read as anything else. */
function sheetPrefix(name: string): string {
	return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$|^[Rr]\d*(?:[Cc]\d*)?$|^[Cc]\d*$/.test(name)
		? `${name}!`
		: `'${name.replace(/'/g, "''")}'!`;
}

/**
 * A form control's cell link or input range as Excel stores it: absolute,
 * and naming a sheet only when it is not the control's own.
 */
function controlReference(value: string, what: 'cell link' | 'input range', sheet: string, workbookSheets: readonly string[]): string {
	const m = /^\s*(?:('(?:[^']|'')+'|[^\s'!:$]+)!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(value);
	const example = what === 'cell link' ? '$H$10' : '$J$1:$J$5';
	if (!m) {
		throw new ShapeError(`'${value}' is not a ${what}; give cells such as ${example}, or Sheet2!${example} on another sheet.`);
	}
	const [, quoted, c1, r1, c2 = c1, r2 = r1] = m;
	const cells = [[c1, r1], [c2, r2]].map(([column, row]) => {
		const col = columnToIndex(column);
		if (col > 16384 || Number(row) < 1 || Number(row) > 1048576) {
			throw new ShapeError(`'${value}' runs past XFD1048576, the last cell of a worksheet.`);
		}
		return `$${indexToColumn(col)}$${Number(row)}`;
	});
	if (what === 'cell link' && cells[0] !== cells[1]) {
		throw new ShapeError(`'${value}' is a range; a cell link is one cell, such as ${example}.`);
	}
	let prefix = '';
	if (quoted !== undefined) {
		const name = quoted.startsWith("'") ? quoted.slice(1, -1).replace(/''/g, "'") : quoted;
		const found = workbookSheets.find((s) => s.toLowerCase() === name.toLowerCase());
		if (found === undefined) {
			throw new ShapeError(`The workbook has no sheet named '${name}'.`);
		}
		if (found.toLowerCase() !== sheet.toLowerCase()) { prefix = sheetPrefix(found); }
	}
	return prefix + (cells[0] === cells[1] ? cells[0] : `${cells[0]}:${cells[1]}`);
}

function rangeOf(from: Marker | undefined, to: Marker | undefined): string | undefined {
	if (!from) { return undefined; }
	if (!to) { return markerCell(from); }
	// A shape ending exactly on a cell's top-left corner does not reach into it.
	const last: Marker = {
		col: to.colOff === 0 && to.col > from.col ? to.col - 1 : to.col,
		row: to.rowOff === 0 && to.row > from.row ? to.row - 1 : to.row,
		colOff: 0,
		rowOff: 0,
	};
	const a = markerCell(from);
	const b = markerCell(last);
	return a === b ? a : `${a}:${b}`;
}

// ------------------------------------------------------------ reading a sheet

/** A shape found in the drawing part, with where it is. */
interface DrawingShape {
	info: ShapeInfo;
	/** The top-level entry: the anchor, or the mc:AlternateContent around it. */
	entry: Span;
	anchor: Span;
	/** The shape element: xdr:sp, xdr:pic, xdr:grpSp, ... */
	element: Span & { name: string };
	/** The cNvPr id. */
	id: number;
	/** Inside a group, not at the top level. */
	inGroup: boolean;
}

/** A form control, from its VML shape and <controls> entry. */
interface FormControl {
	info: ShapeInfo;
	shapeId: number;
	/** The VML shape, when there is one. */
	vml?: Span;
	/** The <controls> entry, and the mc:AlternateContent around it when there is one. */
	control?: Span;
	controlEntry?: Span;
	/** The ctrlProp part and its relationship. */
	ctrlPropPath?: string;
	ctrlPropRel?: string;
}

const CONTROL_KINDS: Record<string, ShapeKind> = {
	button: 'button', checkbox: 'checkBox', radio: 'optionButton', drop: 'dropDown', list: 'listBox',
	scroll: 'scrollBar', spin: 'spinner', label: 'label', gbox: 'groupBox', edit: 'editBox',
};

function shapeKind(element: string, xml: string, span: Span): ShapeKind {
	switch (element) {
		case 'xdr:sp': return /<xdr:cNvSpPr\b[^>]*\btxBox="1"/.test(xml.slice(span.start, span.end)) ? 'textBox' : 'shape';
		case 'xdr:cxnSp': return 'line';
		case 'xdr:pic': return 'picture';
		case 'xdr:grpSp': return 'group';
		case 'xdr:graphicFrame': return /drawingml\/2006\/chart"/.test(xml.slice(span.start, span.end)) ? 'chart' : 'other';
		default: return 'other';
	}
}

/** A macro as Excel shows it: [0]! names this workbook, which it leaves out. */
function displayMacro(stored: string | undefined): string | undefined {
	if (!stored) { return undefined; }
	return stored.replace(/^\[0\]!/, '');
}

function readDrawing(pkg: Package, drawingPath: string, controlIds: ReadonlySet<number>): DrawingShape[] {
	const xml = pkg.read(drawingPath);
	const root = findElement(xml, 'xdr:wsDr');
	if (!root) { return []; }
	const out: DrawingShape[] = [];
	for (const entry of children(xml, root.openEnd, root.end - '</xdr:wsDr>'.length)) {
		let anchor: Span | undefined = entry;
		if (entry.name === 'mc:AlternateContent') {
			const choice = findElement(xml, 'mc:Choice', entry.openEnd, entry.end);
			anchor = choice
				? children(xml, choice.openEnd, choice.end).find((c) => /^xdr:\w+Anchor$/.test(c.name))
				: undefined;
		} else if (!/^xdr:\w+Anchor$/.test(entry.name)) {
			anchor = undefined;
		}
		if (!anchor) { continue; }
		const element = children(xml, anchor.openEnd, anchor.end)
			.find((c) => ['xdr:sp', 'xdr:grpSp', 'xdr:graphicFrame', 'xdr:cxnSp', 'xdr:pic', 'xdr:contentPart'].includes(c.name));
		if (!element) { continue; }
		const anchorXml = xml.slice(anchor.start, anchor.end);
		const range = rangeOf(readMarker(anchorXml, 'from', 'xdr:'), readMarker(anchorXml, 'to', 'xdr:'));
		out.push(...readShapeElement(xml, entry, anchor, element, range, controlIds, false));
	}
	return out;
}

function readShapeElement(
	xml: string,
	entry: Span,
	anchor: Span,
	element: Span & { name: string },
	range: string | undefined,
	controlIds: ReadonlySet<number>,
	inGroup: boolean,
): DrawingShape[] {
	const cNvPr = /<xdr:cNvPr\b[^>]*>/.exec(xml.slice(element.start, element.end))?.[0] ?? '';
	const id = Number(attr(cNvPr, 'id') ?? 0);
	// A form control's hidden DrawingML twin is listed as the control.
	if (controlIds.has(id)) { return []; }
	const kind = shapeKind(element.name, xml, element);
	const startTag = xml.slice(element.start, element.openEnd);
	const info: ShapeInfo = { name: attr(cNvPr, 'name') ?? '', kind };
	const geometry = kind === 'shape' ? /<a:prstGeom\b[^>]*\bprst="([^"]*)"/.exec(xml.slice(element.start, element.end))?.[1] : undefined;
	if (geometry) { info.geometry = geometry; }
	if (range) { info.range = range; }
	const macro = displayMacro(attr(startTag, 'macro'));
	if (macro) { info.macro = macro; }
	const text = kind === 'shape' || kind === 'textBox' ? drawingTextOf(xml, element, 'xdr:txBody') : undefined;
	if (text) { info.text = text; }
	const descr = attr(cNvPr, 'descr');
	if (descr) { info.altText = descr; }
	if (attr(cNvPr, 'hidden') === '1') { info.hidden = true; }
	const found: DrawingShape = { info, entry, anchor, element, id, inGroup };
	const out = [found];
	if (element.name === 'xdr:grpSp') {
		info.shapes = [];
		for (const child of children(xml, element.openEnd, element.end)) {
			if (!['xdr:sp', 'xdr:grpSp', 'xdr:graphicFrame', 'xdr:cxnSp', 'xdr:pic'].includes(child.name)) { continue; }
			const members = readShapeElement(xml, entry, anchor, child, undefined, controlIds, true);
			if (members[0]) { info.shapes.push(members[0].info); }
			out.push(...members);
		}
	}
	return out;
}

/** The caption of a VML shape's text box, as Excel shows it. */
function vmlText(inner: string): string | undefined {
	const box = /<v:textbox\b[^>]*>([\s\S]*?)<\/v:textbox>/.exec(inner)?.[1];
	if (box === undefined) { return undefined; }
	// VML wraps its markup across lines; only <br> breaks the caption's.
	const text = decodeXml(box.replace(/\s+/g, ' ').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ''))
		.split('\n').map((line) => line.trim()).join('\n').trim();
	return text || undefined;
}

function clientDataValue(inner: string, name: string): string | undefined {
	const m = new RegExp(`<x:${name}>([\\s\\S]*?)</x:${name}>`).exec(inner);
	return m ? decodeXml(m[1]).trim() : undefined;
}

/** The anchor of a VML shape: column, x offset, row, y offset, twice over. */
function vmlRange(inner: string): string | undefined {
	const numbers = clientDataValue(inner, 'Anchor')?.split(',').map((n) => Number(n.trim()));
	if (!numbers || numbers.length !== 8) { return undefined; }
	return rangeOf(
		{ col: numbers[0], colOff: numbers[1], row: numbers[2], rowOff: numbers[3] },
		{ col: numbers[4], colOff: numbers[5], row: numbers[6], rowOff: numbers[7] },
	);
}

interface SheetParts {
	name: string;
	path: string;
	drawingPath?: string;
	drawingRel?: string;
	vmlPath?: string;
	vmlRel?: string;
}

function sheetParts(pkg: Package, sheet: { name: string; path: string }): SheetParts {
	const xml = pkg.read(sheet.path);
	const rels = pkg.relationships(sheet.path);
	const byId = (tag: string | undefined): Relationship | undefined => {
		const id = tag ? attr(tag, 'r:id') : undefined;
		return rels.find((rel) => rel.id === id);
	};
	const drawing = byId(/<drawing\b[^>]*>/.exec(xml)?.[0]);
	const vml = byId(/<legacyDrawing\b[^>]*>/.exec(xml)?.[0]);
	return {
		...sheet,
		...(drawing ? { drawingPath: drawing.path, drawingRel: drawing.id } : {}),
		...(vml ? { vmlPath: vml.path, vmlRel: vml.id } : {}),
	};
}

function readControls(pkg: Package, parts: SheetParts): FormControl[] {
	const sheetXml = pkg.read(parts.path);
	const rels = pkg.relationships(parts.path);
	const controls = new Map<number, FormControl>();
	// The <controls> entries: name, macro, alt text and anchor, and the
	// ctrlProp part or ActiveX part each refers to.
	for (let at = 0; ;) {
		const control = findElement(sheetXml, 'control', at);
		if (!control) { break; }
		at = control.end;
		const tag = sheetXml.slice(control.start, control.openEnd);
		const shapeId = Number(attr(tag, 'shapeId') ?? 0);
		// An ActiveX control repeats its entry in an mc:Fallback, after the one that counts.
		if (controls.has(shapeId)) { continue; }
		const rel = rels.find((r) => r.id === attr(tag, 'r:id'));
		const inner = sheetXml.slice(control.start, control.end);
		const controlPr = /<controlPr\b[^>]*>/.exec(inner)?.[0] ?? '';
		const activeX = rel !== undefined && !rel.type.endsWith('/ctrlProp');
		const info: ShapeInfo = { name: attr(tag, 'name') ?? '', kind: activeX ? 'activeX' : 'formControl' };
		const range = rangeOf(readMarker(inner, 'from', ''), readMarker(inner, 'to', ''));
		if (range) { info.range = range; }
		const macro = displayMacro(attr(controlPr, 'macro'));
		if (macro) { info.macro = macro; }
		const altText = attr(controlPr, 'altText');
		if (altText) { info.altText = altText; }
		// Excel 2010 and later wrap each entry in mc:AlternateContent.
		const wrapper = enclosingAlternateContent(sheetXml, control);
		controls.set(shapeId, {
			info,
			shapeId,
			control,
			controlEntry: wrapper ?? control,
			...(rel && !activeX ? { ctrlPropPath: rel.path, ctrlPropRel: rel.id } : {}),
		});
		if (rel && !activeX && pkg.has(rel.path)) {
			const pr = /<formControlPr\b[^>]*>/.exec(pkg.read(rel.path))?.[0] ?? '';
			const kind = CONTROL_KINDS[(attr(pr, 'objectType') ?? '').toLowerCase()];
			if (kind) { info.kind = kind; }
			const link = attr(pr, 'fmlaLink');
			if (link) { info.linkedCell = link; }
			const input = attr(pr, 'fmlaRange');
			if (input) { info.inputRange = input; }
		}
	}
	// The VML shapes: every form control has one, and Excel 2007 keeps no
	// <controls> entry at all. Cell notes live here too, and are not shapes.
	if (parts.vmlPath && pkg.has(parts.vmlPath)) {
		const vml = pkg.read(parts.vmlPath);
		for (let at = 0; ;) {
			const shape = findElement(vml, 'v:shape', at);
			if (!shape) { break; }
			at = shape.end;
			const inner = vml.slice(shape.start, shape.end);
			const objectType = /<x:ClientData\b[^>]*\bObjectType="([^"]*)"/.exec(inner)?.[1];
			if (!objectType || objectType === 'Note') { continue; }
			// An ActiveX control's VML shape is named for the control and keeps its number in o:spid.
			const shapeId = Number(/\b(?:o:spid|id)="_x0000_s(\d+)"/.exec(inner.slice(0, inner.indexOf('>')))?.[1] ?? 0);
			if (shapeId === 0) { continue; }
			const kind = CONTROL_KINDS[objectType.toLowerCase()] ?? 'formControl';
			const existing = controls.get(shapeId);
			const control: FormControl = existing ?? { info: { name: `${objectType} ${shapeId % 1024}`, kind }, shapeId };
			control.vml = shape;
			const info = control.info;
			if (info.kind === 'formControl') { info.kind = kind; }
			info.range ??= vmlRange(inner);
			const macro = displayMacro(clientDataValue(inner, 'FmlaMacro'));
			if (macro && !info.macro) { info.macro = macro; }
			const link = clientDataValue(inner, 'FmlaLink');
			if (link && !info.linkedCell) { info.linkedCell = link; }
			const input = clientDataValue(inner, 'FmlaRange');
			if (input && !info.inputRange) { info.inputRange = input; }
			const text = vmlText(inner);
			if (text && !['dropDown', 'listBox', 'scrollBar', 'spinner'].includes(info.kind)) { info.text = text; }
			if (/visibility:\s*hidden/.test(inner.slice(0, inner.indexOf('>')))) { info.hidden = true; }
			controls.set(shapeId, control);
		}
	}
	return [...controls.values()];
}

/** The mc:AlternateContent directly around an element, when there is one. */
function enclosingAlternateContent(xml: string, span: Span): Span | undefined {
	const before = xml.lastIndexOf('<mc:AlternateContent', span.start);
	if (before < 0) { return undefined; }
	const wrapper = findElement(xml, 'mc:AlternateContent', before);
	if (!wrapper || wrapper.end < span.end) { return undefined; }
	// Only the element and its mc:Choice, nothing else, inside.
	const inside = xml.slice(wrapper.openEnd, wrapper.end).replace(xml.slice(span.start, span.end), '');
	return /^\s*<mc:Choice\b[^>]*>\s*<\/mc:Choice>\s*(?:<mc:Fallback\/>|<mc:Fallback>\s*<\/mc:Fallback>)?\s*<\/mc:AlternateContent>$/.test(inside)
		? wrapper
		: undefined;
}

interface SheetShapes {
	parts: SheetParts;
	drawing: DrawingShape[];
	controls: FormControl[];
}

function readSheet(pkg: Package, sheet: { name: string; path: string }): SheetShapes {
	const parts = sheetParts(pkg, sheet);
	const controls = readControls(pkg, parts);
	const ids = new Set(controls.map((c) => c.shapeId));
	const drawing = parts.drawingPath && pkg.has(parts.drawingPath) ? readDrawing(pkg, parts.drawingPath, ids) : [];
	return { parts, drawing, controls };
}

/** Every shape on a sheet, top level only; a group lists its members. */
export function listSheetShapes(zip: ZipArchive, sheet: { name: string; path: string }): ShapeInfo[] {
	const { drawing, controls } = readSheet(new Package(zip), sheet);
	return [...drawing.filter((s) => !s.inGroup).map((s) => s.info), ...controls.map((c) => c.info)];
}

// ------------------------------------------------------------------ editing

type Target =
	| { kind: 'drawing'; shape: DrawingShape }
	| { kind: 'control'; control: FormControl };

function findTarget(shapes: SheetShapes, name: string, sheet: string): Target {
	const lower = name.toLowerCase();
	const matches: Target[] = [
		...shapes.drawing.filter((s) => s.info.name.toLowerCase() === lower).map((shape) => ({ kind: 'drawing' as const, shape })),
		...shapes.controls.filter((c) => c.info.name.toLowerCase() === lower).map((control) => ({ kind: 'control' as const, control })),
	];
	if (matches.length === 0) {
		throw new ShapeError(`No shape named '${name}' on sheet '${sheet}'.`);
	}
	if (matches.length > 1) {
		throw new ShapeError(`${matches.length} shapes on sheet '${sheet}' are named '${name}'; rename one in Excel first.`);
	}
	return matches[0];
}

function assertNameFree(shapes: SheetShapes, name: string, sheet: string): void {
	const lower = name.toLowerCase();
	if ([...shapes.drawing, ...shapes.controls].some((s) => s.info.name.toLowerCase() === lower)) {
		throw new ShapeError(`Sheet '${sheet}' already has a shape named '${name}'.`);
	}
}

function updateDrawingShape(pkg: Package, parts: SheetParts, shapes: SheetShapes, shape: DrawingShape, edit: ShapeEdit): void {
	const path = parts.drawingPath!;
	const xml = pkg.read(path);
	const { element, anchor } = shape;
	let elementXml = xml.slice(element.start, element.end);
	if (edit.text !== undefined) {
		const body = /<xdr:txBody>[\s\S]*<\/xdr:txBody>/.exec(elementXml);
		if (!body || (shape.info.kind !== 'shape' && shape.info.kind !== 'textBox')) {
			throw new ShapeError(`'${shape.info.name}' is a ${shape.info.kind}, which holds no text.`);
		}
		elementXml = elementXml.replace(body[0], () => withDrawingTextBody(body[0], edit.text!, 'xdr:txBody'));
	}
	if (edit.macro !== undefined) {
		if (shape.info.kind === 'group') {
			throw new ShapeError(`'${shape.info.name}' is a group, which Excel runs no macro for; assign the macro to a shape in it.`);
		}
		const startTag = elementXml.slice(0, elementXml.indexOf('>') + 1);
		elementXml = withAttr(startTag, 'macro', edit.macro || '') + elementXml.slice(startTag.length);
	}
	if (edit.newName !== undefined || edit.altText !== undefined) {
		const cNvPr = /<xdr:cNvPr\b[^>]*>/.exec(elementXml)![0];
		let updated = cNvPr;
		if (edit.newName !== undefined) {
			assertNameFree(shapes, edit.newName, parts.name);
			updated = withAttr(updated, 'name', edit.newName);
		}
		if (edit.altText !== undefined) { updated = withAttr(updated, 'descr', edit.altText || undefined); }
		elementXml = elementXml.replace(cNvPr, () => updated);
	}
	if (edit.linkedCell !== undefined || edit.inputRange !== undefined) {
		throw new ShapeError(`'${shape.info.name}' is not a form control, so it has no ${edit.linkedCell !== undefined ? 'cell link' : 'input range'}.`);
	}
	// The shape sits inside its anchor, so the anchor is rebuilt around it.
	const before = xml.slice(anchor.start, element.start);
	const after = xml.slice(element.end, anchor.end);
	let anchorXml = before + elementXml + after;
	if (edit.range !== undefined) {
		if (shape.inGroup) {
			throw new ShapeError(`'${shape.info.name}' is in a group; move the group instead.`);
		}
		if (!anchorXml.startsWith('<xdr:twoCellAnchor')) {
			throw new ShapeError(`'${shape.info.name}' is anchored to one cell or to a position, and XLIDE moves only shapes anchored to cells at both corners.`);
		}
		anchorXml = anchorXml.replace(/<xdr:from>[\s\S]*?<\/xdr:to>/, anchorMarkers(parseRange(edit.range), 'xdr:'));
	}
	pkg.write(path, splice(xml, anchor, anchorXml));
}

function deleteDrawingShape(pkg: Package, parts: SheetParts, shape: DrawingShape): void {
	if (shape.inGroup) {
		throw new ShapeError(`'${shape.info.name}' is in a group; delete the group, or ungroup it in Excel first.`);
	}
	const path = parts.drawingPath!;
	const xml = pkg.read(path);
	const removed = xml.slice(shape.entry.start, shape.entry.end);
	const rest = splice(xml, shape.entry, '');
	pkg.write(path, rest);
	// A picture's image and a chart's parts go with it, unless still used.
	for (const id of [...removed.matchAll(/\br:(?:embed|link|id)="([^"]+)"/g)].map((m) => m[1])) {
		if (new RegExp(`\\br:(?:embed|link|id)="${id}"`).test(rest)) { continue; }
		const rel = pkg.relationships(path).find((r) => r.id === id);
		pkg.removeRelationship(path, id);
		if (rel && pkg.has(rel.path)) { pkg.removePart(rel.path); }
	}
	removeDrawingIfEmpty(pkg, parts);
}

/** A drawing part left with no shapes goes, with the sheet's reference to it. */
function removeDrawingIfEmpty(pkg: Package, parts: SheetParts): void {
	const path = parts.drawingPath!;
	const xml = pkg.read(path);
	const root = findElement(xml, 'xdr:wsDr')!;
	if (children(xml, root.openEnd, root.end - '</xdr:wsDr>'.length).length > 0) { return; }
	pkg.write(parts.path, pkg.read(parts.path).replace(/<drawing\b[^>]*\/>/, ''));
	pkg.removeRelationship(parts.path, parts.drawingRel!);
	pkg.removePart(path);
}

// Form controls -------------------------------------------------------------

function setClientData(vmlShape: string, name: string, value: string | undefined, before: string[]): string {
	const element = new RegExp(`\\s*<x:${name}>[\\s\\S]*?</x:${name}>`);
	const stripped = vmlShape.replace(element, '');
	if (value === undefined) { return stripped; }
	const tag = `\n   <x:${name}>${encodeXml(value)}</x:${name}>`;
	for (const next of before) {
		const at = stripped.search(new RegExp(`\\s*<x:${next}[\\s/>]`));
		if (at >= 0) { return stripped.slice(0, at) + tag + stripped.slice(at); }
	}
	return stripped.replace(/\s*<\/x:ClientData>/, (close) => tag + close);
}

function vmlAnchor(box: CellBox): string {
	return `\n    ${box.c1}, 0, ${box.r1}, 0, ${box.c2 + 1}, 0, ${box.r2 + 1}, 0`;
}

/** The DrawingML twin of a form control in the drawing part, found by id. */
function controlTwin(pkg: Package, parts: SheetParts, shapeId: number): { entry: Span; anchor: Span; xml: string } | undefined {
	if (!parts.drawingPath || !pkg.has(parts.drawingPath)) { return undefined; }
	const xml = pkg.read(parts.drawingPath);
	const root = findElement(xml, 'xdr:wsDr');
	if (!root) { return undefined; }
	for (const entry of children(xml, root.openEnd, root.end - '</xdr:wsDr>'.length)) {
		const inner = xml.slice(entry.start, entry.end);
		if (new RegExp(`<xdr:cNvPr\\b[^>]*\\bid="${shapeId}"`).test(inner)) {
			const anchor = findElement(xml, 'xdr:twoCellAnchor', entry.start, entry.end) ?? entry;
			return { entry, anchor, xml };
		}
	}
	return undefined;
}

function updateControl(pkg: Package, parts: SheetParts, shapes: SheetShapes, control: FormControl, edit: ShapeEdit): void {
	const { info } = control;
	if (info.kind === 'activeX') {
		throw new ShapeError(`'${info.name}' is an ActiveX control, which XLIDE does not edit; its code is event procedures in the sheet's module, such as ${info.name}_Click.`);
	}
	const captioned = ['button', 'checkBox', 'optionButton', 'label', 'groupBox'].includes(info.kind);
	const linked = ['checkBox', 'optionButton', 'dropDown', 'listBox', 'scrollBar', 'spinner'].includes(info.kind);
	if (edit.text !== undefined && !captioned) {
		throw new ShapeError(`'${info.name}' is a ${info.kind}, which has no caption.`);
	}
	if (edit.linkedCell !== undefined && !linked) {
		throw new ShapeError(`'${info.name}' is a ${info.kind}, which has no cell link.`);
	}
	if (edit.inputRange !== undefined && info.kind !== 'dropDown' && info.kind !== 'listBox') {
		throw new ShapeError(`'${info.name}' is a ${info.kind}; only a drop-down or list box has an input range.`);
	}
	if (edit.newName !== undefined) { assertNameFree(shapes, edit.newName, parts.name); }
	const box = edit.range !== undefined ? parseRange(edit.range) : undefined;

	// The VML shape.
	if (control.vml && parts.vmlPath) {
		const vml = pkg.read(parts.vmlPath);
		let shape = vml.slice(control.vml.start, control.vml.end);
		if (edit.macro !== undefined) {
			shape = setClientData(shape, 'FmlaMacro', edit.macro || undefined, ['TextHAlign', 'TextVAlign', 'LockText', 'FmlaLink', 'Val']);
		}
		if (edit.linkedCell !== undefined) {
			shape = setClientData(shape, 'FmlaLink', edit.linkedCell || undefined, ['NoThreeD', 'Val', 'FmlaRange']);
		}
		if (edit.inputRange !== undefined) {
			shape = setClientData(shape, 'FmlaRange', edit.inputRange || undefined, ['Sel', 'NoThreeD2', 'SelType', 'LCT', 'DropStyle', 'DropLines']);
		}
		if (box) {
			shape = shape.replace(/(<x:Anchor>)[\s\S]*?(<\/x:Anchor>)/, `$1${vmlAnchor(box)}$2`);
		}
		if (edit.text !== undefined) {
			shape = shape.replace(/(<v:textbox\b[^>]*>\s*<div\b[^>]*>)[\s\S]*?(<\/div>\s*<\/v:textbox>)/, (_m, open: string, close: string) => {
				const font = /<font\b[^>]*>/.exec(shape)?.[0];
				const lines = edit.text!.split(/\r?\n/).map(encodeXml).join('<br>');
				return `${open}${font ? `${font}${lines}</font>` : lines}${close}`;
			});
		}
		pkg.write(parts.vmlPath, splice(vml, control.vml, shape));
	}

	// The <controls> entry.
	if (control.control) {
		const sheet = pkg.read(parts.path);
		let entry = sheet.slice(control.control.start, control.control.end);
		if (edit.newName !== undefined) {
			const tag = entry.slice(0, entry.indexOf('>') + 1);
			entry = withAttr(tag, 'name', edit.newName) + entry.slice(tag.length);
		}
		if (edit.macro !== undefined || edit.altText !== undefined) {
			const controlPr = /<controlPr\b[^>]*>/.exec(entry)![0];
			let updated = controlPr;
			if (edit.macro !== undefined) { updated = withAttr(updated, 'macro', edit.macro || undefined); }
			if (edit.altText !== undefined) { updated = withAttr(updated, 'altText', edit.altText || undefined); }
			entry = entry.replace(controlPr, () => updated);
		}
		if (box) {
			entry = entry.replace(/<from>[\s\S]*?<\/to>/, anchorMarkers(box, ''));
		}
		pkg.write(parts.path, splice(sheet, control.control, entry));
	}

	// The ctrlProp part.
	if (control.ctrlPropPath && pkg.has(control.ctrlPropPath) && (edit.linkedCell !== undefined || edit.inputRange !== undefined)) {
		const xml = pkg.read(control.ctrlPropPath);
		const tag = /<formControlPr\b[^>]*>/.exec(xml)![0];
		let updated = tag;
		if (edit.linkedCell !== undefined) { updated = withAttr(updated, 'fmlaLink', edit.linkedCell || undefined); }
		if (edit.inputRange !== undefined) { updated = withAttr(updated, 'fmlaRange', edit.inputRange || undefined); }
		pkg.write(control.ctrlPropPath, xml.replace(tag, () => updated));
	}

	// The hidden DrawingML twin.
	const twin = controlTwin(pkg, parts, control.shapeId);
	if (twin) {
		let entry = twin.xml.slice(twin.entry.start, twin.entry.end);
		const cNvPr = /<xdr:cNvPr\b[^>]*>/.exec(entry)![0];
		let updated = cNvPr;
		if (edit.newName !== undefined) { updated = withAttr(updated, 'name', edit.newName); }
		if (edit.altText !== undefined) { updated = withAttr(updated, 'descr', edit.altText || undefined); }
		entry = entry.replace(cNvPr, () => updated);
		if (edit.text !== undefined) {
			const body = /<xdr:txBody>[\s\S]*<\/xdr:txBody>/.exec(entry);
			if (body) { entry = entry.replace(body[0], () => withDrawingTextBody(body[0], edit.text!, 'xdr:txBody')); }
		}
		if (box) {
			entry = entry.replace(/<xdr:from>[\s\S]*?<\/xdr:to>/, anchorMarkers(box, 'xdr:'));
		}
		pkg.write(parts.drawingPath!, splice(twin.xml, twin.entry, entry));
	}
}

function deleteControl(pkg: Package, parts: SheetParts, control: FormControl): void {
	if (control.info.kind === 'activeX') {
		throw new ShapeError(`'${control.info.name}' is an ActiveX control, which XLIDE does not remove; delete it in Excel.`);
	}
	// The <controls> entry first: its place was found in the sheet as it was read.
	if (control.controlEntry) {
		let sheet = splice(pkg.read(parts.path), control.controlEntry, '');
		// A <controls> left empty goes, with the mc:AlternateContent Excel puts around it.
		sheet = sheet.replace(/<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<controls>\s*<\/controls>\s*<\/mc:Choice>\s*(?:<mc:Fallback\/>|<mc:Fallback>\s*<\/mc:Fallback>)?\s*<\/mc:AlternateContent>|<controls>\s*<\/controls>/, '');
		pkg.write(parts.path, sheet);
	}
	if (control.vml && parts.vmlPath) {
		const vml = splice(pkg.read(parts.vmlPath), control.vml, '');
		pkg.write(parts.vmlPath, vml);
		// A VML part left with no shapes goes; one that still holds notes stays.
		if (!/<v:shape\b/.test(vml)) {
			pkg.write(parts.path, pkg.read(parts.path).replace(/<legacyDrawing\b[^>]*\/>/, ''));
			pkg.removeRelationship(parts.path, parts.vmlRel!);
			pkg.removePart(parts.vmlPath);
		}
	}
	if (control.ctrlPropRel) {
		pkg.removeRelationship(parts.path, control.ctrlPropRel);
		if (control.ctrlPropPath && pkg.has(control.ctrlPropPath)) { pkg.removePart(control.ctrlPropPath); }
	}
	const twin = controlTwin(pkg, parts, control.shapeId);
	if (twin) {
		pkg.write(parts.drawingPath!, splice(twin.xml, twin.entry, ''));
		removeDrawingIfEmpty(pkg, parts);
	}
}

// Adding ---------------------------------------------------------------------

/** Where an element goes in a worksheet, in the order the schema fixes. */
const SHEET_TAIL_ORDER = [
	'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls',
	'webPublishItems', 'tableParts', 'extLst',
];

/** Insert `text` as the worksheet child `name`, before the children the schema puts after it. */
function insertSheetChild(sheet: string, name: string, text: string): string {
	const later = SHEET_TAIL_ORDER.slice(SHEET_TAIL_ORDER.indexOf(name) + 1);
	// Excel wraps <controls> and <oleObjects> in mc:AlternateContent.
	const candidates = [
		...later.map((next) => sheet.search(new RegExp(`<${next}[\\s/>]`))),
		...(later.includes('controls') || later.includes('oleObjects')
			? [sheet.search(/<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<(?:controls|oleObjects)[\s>]/)]
			: []),
	].filter((at) => at >= 0);
	const at = candidates.length > 0 ? Math.min(...candidates) : sheet.lastIndexOf('</worksheet>');
	return sheet.slice(0, at) + text + sheet.slice(at);
}

/** The sheet's drawing part, made when it has none. */
function ensureDrawing(pkg: Package, parts: SheetParts): string {
	if (parts.drawingPath && pkg.has(parts.drawingPath)) { return parts.drawingPath; }
	const path = pkg.freePath('xl/drawings/drawing', '.xml');
	pkg.write(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
		+ `<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`);
	pkg.addOverride(path, CONTENT_TYPE.drawing);
	const id = pkg.addRelationship(parts.path, REL.drawing, path);
	pkg.write(parts.path, insertSheetChild(pkg.read(parts.path), 'drawing', `<drawing r:id="${id}"/>`));
	parts.drawingPath = path;
	parts.drawingRel = id;
	return path;
}

/** The next cNvPr id free in a drawing, below the range VML controls take. */
function nextDrawingId(xml: string, taken: ReadonlySet<number>): number {
	const ids = [...xml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])).filter((id) => id < 1024);
	let id = Math.max(1, ...ids) + 1;
	while (taken.has(id)) { id++; }
	return id;
}

/** The style Excel 16 gives a new AutoShape: the theme's first accent, white text. */
const AUTOSHAPE_STYLE = '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
	+ '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
	+ '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>';

/** What Excel 16 gives a new text box: white fill, a grey line, dark text. */
const TEXTBOX_FILL = '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525" cmpd="sng"><a:solidFill>'
	+ '<a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>';
const TEXTBOX_STYLE = '<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0">'
	+ '<a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>'
	+ '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style>';

function addDrawingShape(pkg: Package, parts: SheetParts, shapes: SheetShapes, edit: ShapeEdit, box: CellBox): string {
	const path = ensureDrawing(pkg, parts);
	const xml = pkg.read(path);
	const taken = new Set(shapes.controls.map((c) => c.shapeId));
	const id = nextDrawingId(xml, taken);
	const textBox = edit.type === 'textBox';
	const type = edit.type as PresetShapeType;
	const preset = textBox
		? { prst: 'rect', label: 'TextBox' }
		: { prst: PRESET_GEOMETRY[type], label: PRESET_LABELS[type] };
	const name = edit.name ?? `${preset.label} ${id - 1}`;
	assertNameFree(shapes, name, parts.name);
	const descr = edit.altText ? ` descr="${encodeXml(edit.altText)}"` : '';
	const body = withDrawingTextBody(
		`<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"${textBox ? ' vert="horz"' : ''} rtlCol="0" anchor="t"/><a:lstStyle/></xdr:txBody>`,
		edit.text ?? '',
		'xdr:txBody',
	);
	const shape = `<xdr:twoCellAnchor>${anchorMarkers(box, 'xdr:')}`
		+ `<xdr:sp macro="${encodeXml(edit.macro ?? '')}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${encodeXml(name)}"${descr}/>`
		+ `<xdr:cNvSpPr${textBox ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>`
		+ `<a:prstGeom prst="${preset.prst}"><a:avLst/></a:prstGeom>${textBox ? TEXTBOX_FILL : ''}</xdr:spPr>`
		+ `${textBox ? TEXTBOX_STYLE : AUTOSHAPE_STYLE}${body}</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
	pkg.write(path, xml.replace(/<\/xdr:wsDr>\s*$/, () => `${shape}</xdr:wsDr>`));
	return name;
}

/** The VML shape type every form control refers to. */
const CONTROL_SHAPETYPE = '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"\n'
	+ '  path="m,l,21600r21600,l21600,xe">\n  <v:stroke joinstyle="miter"/>\n'
	+ '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\n'
	+ '  <o:lock v:ext="edit" shapetype="t"/>\n </v:shapetype>';

/**
 * The sheet's VML part and the next shape id free in its id block, made when
 * the sheet has none. VML shape ids are 1024 per block, and each VML part of
 * the workbook takes its own block, named by o:idmap.
 */
function ensureVml(pkg: Package, parts: SheetParts): { path: string; shapeId: number } {
	if (parts.vmlPath && pkg.has(parts.vmlPath)) {
		let vml = pkg.read(parts.vmlPath);
		const block = Number(/<o:idmap\b[^>]*\bdata="(\d+)/.exec(vml)?.[1] ?? 1);
		const ids = [...vml.matchAll(/\bid="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
		if (!vml.includes('id="_x0000_t201"')) {
			vml = vml.replace(/(<\/o:shapelayout>)/, `$1${CONTROL_SHAPETYPE}`);
			pkg.write(parts.vmlPath, vml);
		}
		return { path: parts.vmlPath, shapeId: Math.max(block * 1024, ...ids) + 1 };
	}
	const blocks = pkg.names().filter((name) => /\.vml$/i.test(name))
		.flatMap((name) => [...pkg.read(name).matchAll(/<o:idmap\b[^>]*\bdata="([\d,]+)"/g)].flatMap((m) => m[1].split(',').map(Number)));
	const block = Math.max(0, ...blocks) + 1;
	const path = pkg.freePath('xl/drawings/vmlDrawing', '.vml');
	pkg.write(path, '<xml xmlns:v="urn:schemas-microsoft-com:vml"\n xmlns:o="urn:schemas-microsoft-com:office:office"\n'
		+ ` xmlns:x="urn:schemas-microsoft-com:office:excel">\n <o:shapelayout v:ext="edit">\n  <o:idmap v:ext="edit" data="${block}"/>\n`
		+ ` </o:shapelayout>${CONTROL_SHAPETYPE}</xml>`);
	pkg.ensureDefault('vml', CONTENT_TYPE.vml);
	const id = pkg.addRelationship(parts.path, REL.vmlDrawing, path);
	pkg.write(parts.path, insertSheetChild(pkg.read(parts.path), 'legacyDrawing', `<legacyDrawing r:id="${id}"/>`));
	parts.vmlPath = path;
	parts.vmlRel = id;
	return { path, shapeId: block * 1024 + 1 };
}

function addButton(pkg: Package, parts: SheetParts, shapes: SheetShapes, edit: ShapeEdit, box: CellBox): string {
	const { path: vmlPath, shapeId } = ensureVml(pkg, parts);
	const name = edit.name ?? `Button ${shapeId % 1024}`;
	assertNameFree(shapes, name, parts.name);
	const caption = encodeXml(edit.text ?? name).split(/\r?\n/).join('<br>');
	const macro = edit.macro ? `\n   <x:FmlaMacro>${encodeXml(edit.macro)}</x:FmlaMacro>` : '';
	// The VML shape, as Excel 16 writes a new button. Excel places the button
	// by the anchors below; the margins only serve readers of VML alone.
	const vml = pkg.read(vmlPath);
	pkg.write(vmlPath, vml.replace(/<\/xml>\s*$/, () => `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t201" style='position:absolute;\n`
		+ `  margin-left:${box.c1 * 48}pt;margin-top:${box.r1 * 14.5}pt;width:${(box.c2 - box.c1 + 1) * 48}pt;`
		+ `height:${(box.r2 - box.r1 + 1) * 14.5}pt;z-index:1;\n  mso-wrap-style:tight' o:button="t" fillcolor="buttonFace [67]" o:insetmode="auto">\n`
		+ '  <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>\n  <o:lock v:ext="edit" rotation="t"/>\n'
		+ `  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">\n   <div style='text-align:center'><font face="Calibri" size="220"\n`
		+ `   color="#000000">${caption}</font></div>\n  </v:textbox>\n  <x:ClientData ObjectType="Button">\n   <x:Anchor>${vmlAnchor(box)}</x:Anchor>\n`
		+ `   <x:PrintObject>False</x:PrintObject>\n   <x:AutoFill>False</x:AutoFill>${macro}\n`
		+ '   <x:TextHAlign>Center</x:TextHAlign>\n   <x:TextVAlign>Center</x:TextVAlign>\n  </x:ClientData>\n </v:shape></xml>'));

	// The ctrlProp part.
	const ctrlPropPath = pkg.freePath('xl/ctrlProps/ctrlProp', '.xml');
	pkg.write(ctrlPropPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
		+ `<formControlPr xmlns="${NS.x14}" objectType="Button" lockText="1"/>`);
	pkg.addOverride(ctrlPropPath, CONTENT_TYPE.ctrlProp);
	const rel = pkg.addRelationship(parts.path, REL.ctrlProp, ctrlPropPath);

	// The <controls> entry.
	const macroAttr = edit.macro ? ` macro="${encodeXml(edit.macro)}"` : '';
	const altAttr = edit.altText ? ` altText="${encodeXml(edit.altText)}"` : '';
	const entry = `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="x14"><control shapeId="${shapeId}" r:id="${rel}" name="${encodeXml(name)}">`
		+ `<controlPr defaultSize="0" print="0" autoFill="0" autoPict="0"${macroAttr}${altAttr}><anchor moveWithCells="1" sizeWithCells="1">`
		+ `${anchorMarkers(box, '')}</anchor></controlPr></control></mc:Choice></mc:AlternateContent>`;
	let sheet = pkg.read(parts.path);
	const controls = findElement(sheet, 'controls');
	if (controls) {
		sheet = sheet.slice(0, controls.end - '</controls>'.length) + entry + sheet.slice(controls.end - '</controls>'.length);
	} else {
		sheet = insertSheetChild(sheet, 'controls', `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="x14"><controls>${entry}</controls></mc:Choice></mc:AlternateContent>`);
	}
	// The markers are in the drawing namespace, and x14 has to be known where it is required.
	sheet = ensureRootNamespace(sheet, 'xdr', NS.xdr);
	sheet = ensureRootNamespace(sheet, 'x14', NS.x14);
	sheet = ensureRootNamespace(sheet, 'mc', NS.mc);
	pkg.write(parts.path, sheet);

	// The hidden DrawingML twin, which Excel 2010 and later draw the button from.
	const drawingPath = ensureDrawing(pkg, parts);
	const drawing = pkg.read(drawingPath);
	const descr = edit.altText ? ` descr="${encodeXml(edit.altText)}"` : '';
	const twinText = withDrawingTextBody('<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" lIns="36576" tIns="36576" rIns="36576" bIns="36576" anchor="ctr" upright="1"/><a:lstStyle/></xdr:txBody>', edit.text ?? name, 'xdr:txBody')
		.replace(/<a:p>/g, '<a:p><a:pPr algn="ctr" rtl="0"><a:defRPr sz="1000"/></a:pPr>')
		.replace(/<a:rPr lang="en-US" sz="1100"\/>/g, '<a:rPr lang="en-US" sz="1100" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>');
	const twin = `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:a14="${NS.a14}" Requires="a14"><xdr:twoCellAnchor>${anchorMarkers(box, 'xdr:')}`
		+ `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${shapeId}" name="${encodeXml(name)}" hidden="1"${descr}><a:extLst>`
		+ `<a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s${shapeId}"/></a:ext></a:extLst></xdr:cNvPr>`
		+ '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
		+ '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></xdr:spPr>'
		+ `${twinText}</xdr:sp><xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>`;
	pkg.write(drawingPath, drawing.replace(/<\/xdr:wsDr>\s*$/, () => `${twin}</xdr:wsDr>`));
	return name;
}

/** The worksheet root with a namespace prefix declared, when it has none. */
function ensureRootNamespace(sheet: string, prefix: string, uri: string): string {
	const root = /<worksheet\b[^>]*>/.exec(sheet)![0];
	return new RegExp(`\\sxmlns:${prefix}="`).test(root)
		? sheet
		: sheet.replace(root, root.replace(/^<worksheet\b/, `<worksheet xmlns:${prefix}="${uri}"`));
}

/**
 * Apply one edit to the shapes on a sheet, and give the shape's name after it;
 * `workbookSheets` names every sheet a cell link may point at.
 */
export function editSheetShape(
	zip: ZipArchive,
	sheet: { name: string; path: string },
	requested: ShapeEdit,
	workbookSheets: readonly string[],
): string {
	const pkg = new Package(zip);
	const shapes = readSheet(pkg, sheet);
	const { parts } = shapes;
	const edit = { ...requested };
	if (edit.linkedCell) { edit.linkedCell = controlReference(edit.linkedCell, 'cell link', sheet.name, workbookSheets); }
	if (edit.inputRange) { edit.inputRange = controlReference(edit.inputRange, 'input range', sheet.name, workbookSheets); }
	if (edit.action === 'add') {
		if (!edit.type) {
			throw new ShapeError('Adding a shape needs its type: rectangle, roundedRectangle, oval, textBox or button.');
		}
		if (!edit.range) {
			throw new ShapeError('Adding a shape needs the cells it covers, such as B2:D4.');
		}
		if (edit.newName !== undefined || edit.linkedCell !== undefined || edit.inputRange !== undefined) {
			throw new ShapeError('A new shape takes its name from name, and a button has no cell link or input range.');
		}
		const box = parseRange(edit.range);
		if (edit.type === 'button') {
			return addButton(pkg, parts, shapes, edit, box);
		}
		if (edit.type === 'textBox' || Object.hasOwn(PRESET_GEOMETRY, edit.type)) {
			return addDrawingShape(pkg, parts, shapes, edit, box);
		}
		throw new ShapeError(`'${edit.type}' is not a shape XLIDE adds: use rectangle, roundedRectangle, oval, textBox or button.`);
	}
	if (!edit.name) {
		throw new ShapeError(`To ${edit.action} a shape, give its name; xlide_listShapes lists them.`);
	}
	const target = findTarget(shapes, edit.name, parts.name);
	const name = target.kind === 'drawing' ? target.shape.info.name : target.control.info.name;
	if (edit.action === 'delete') {
		if (target.kind === 'drawing') { deleteDrawingShape(pkg, parts, target.shape); } else { deleteControl(pkg, parts, target.control); }
		return name;
	}
	if (target.kind === 'drawing') {
		updateDrawingShape(pkg, parts, shapes, target.shape, edit);
	} else {
		updateControl(pkg, parts, shapes, target.control, edit);
	}
	return edit.newName ?? name;
}
