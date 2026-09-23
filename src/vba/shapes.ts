// What a shape is, in the words all three Office hosts share.
//
// Excel, Word and PowerPoint each keep shapes in a different part and name
// them with a different element, but a shape is the same object in all
// three: it has a name, a geometry, a size, some text, and - in Excel and
// PowerPoint - a macro a click runs. This module holds that vocabulary and
// the DrawingML text handling the three have in common; each host's own
// module (xlsxShapes, pptShapes, docShapes) owns its markup.
//
// Where a host has nothing to put in a field, the field is absent rather
// than defaulted, so "no macro" and "this host cannot hold a macro" stay
// distinguishable to a caller.

import { decodeXml, encodeXml, findElement, type Span } from './ooxml';
import type { FillEdit, FontEdit, LineEdit, ShapeFill, ShapeFont, ShapeLine } from './shapeFormat';

export type { FillEdit, FontEdit, LineEdit, ShapeDash, ShapeFill, ShapeFont, ShapeLine } from './shapeFormat';

export class ShapeError extends Error {}

/**
 * The applications' four Arrange commands. Each moves one shape among the
 * top-level shapes of its surface; a shape inside a group stacks with it.
 */
export type ZOrderCommand = 'front' | 'back' | 'forward' | 'backward';

export type ShapeKind =
	| 'shape' | 'textBox' | 'line' | 'picture' | 'chart' | 'group'
	| 'button' | 'checkBox' | 'optionButton' | 'dropDown' | 'listBox' | 'scrollBar' | 'spinner'
	| 'label' | 'groupBox' | 'editBox' | 'formControl' | 'activeX'
	| 'placeholder' | 'table' | 'canvas' | 'other';

/** Where a Word shape sits relative to the text. */
export type ShapePlacement = 'inline' | 'anchored';

export interface ShapeInfo {
	name: string;
	kind: ShapeKind;
	/** An AutoShape's DrawingML preset geometry, such as rect or ellipse. */
	geometry?: string;
	/** Excel only: the cells the shape covers, top-left to bottom-right. */
	range?: string;
	/**
	 * Word and PowerPoint: position and size in points, as the Office object
	 * model reports them. A Word inline shape flows with the text and has a
	 * size but no position.
	 */
	left?: number;
	top?: number;
	width?: number;
	height?: number;
	/** Word only. */
	placement?: ShapePlacement;
	/** The macro a click runs. Word shapes cannot carry one. */
	macro?: string;
	text?: string;
	/** An Excel form control's cell link. */
	linkedCell?: string;
	/** An Excel drop-down or list box's items. */
	inputRange?: string;
	altText?: string;
	hidden?: boolean;
	/** Degrees clockwise. */
	rotation?: number;
	fill?: ShapeFill;
	line?: ShapeLine;
	/** The font of the shape's text, or of an Excel button's caption. */
	font?: ShapeFont;
	/**
	 * Where the shape stacks among its surface's top-level shapes: 1 is at
	 * the back. A Word shape inline with the text does not stack.
	 */
	zOrder?: number;
	/** The shapes inside a group. */
	shapes?: ShapeInfo[];
}

/**
 * What can be added. `button` is an Excel form control and has no
 * counterpart in the other hosts; pictures and charts are listed and edited
 * but not added, since making one needs image or data bytes a shape tool
 * has no way to supply.
 */
export type NewShapeType =
	| 'rectangle' | 'parallelogram' | 'trapezoid' | 'diamond' | 'roundedRectangle'
	| 'octagon' | 'triangle' | 'rightTriangle' | 'oval' | 'hexagon'
	| 'cross' | 'pentagon' | 'cylinder' | 'cube' | 'bevel'
	| 'foldedCorner' | 'smileyFace' | 'donut' | 'noEntry' | 'blockArc' | 'star'
	| 'textBox' | 'button';

/** One change to one shape. Properties left out are left as they are. */
export interface ShapeEdit {
	action: 'add' | 'update' | 'delete';
	/** The shape to update or delete, or the new shape's name. */
	name?: string;
	/** What to add. */
	type?: NewShapeType;
	/** Excel: the cells the shape covers, such as B2:D4. */
	range?: string;
	/** Word and PowerPoint: position and size in points. */
	left?: number;
	top?: number;
	width?: number;
	height?: number;
	text?: string;
	/** The macro to run on a click; an empty string removes it. */
	macro?: string;
	/** An Excel form control's cell link; an empty string removes it. */
	linkedCell?: string;
	/** The cells an Excel drop-down or list box lists; an empty string removes them. */
	inputRange?: string;
	altText?: string;
	newName?: string;
	hidden?: boolean;
	/** Degrees clockwise; 0 takes the rotation off. */
	rotation?: number;
	fill?: FillEdit;
	line?: LineEdit;
	font?: FontEdit;
	zOrder?: ZOrderCommand;
}

export const Z_ORDER_COMMANDS: readonly ZOrderCommand[] = ['front', 'back', 'forward', 'backward'];

/**
 * Where a z-order command moves the item at `index` in a list of `count`,
 * 0-based, the way the applications' Arrange commands move a shape among
 * the shapes on its surface.
 */
export function restackedIndex(index: number, count: number, command: ZOrderCommand): number {
	switch (command) {
		case 'front': return count - 1;
		case 'back': return 0;
		case 'forward': return Math.min(count - 1, index + 1);
		default: return Math.max(0, index - 1);
	}
}

/**
 * The part of an add that is applied to the shape once it exists - its
 * look, visibility and stacking - or undefined when the add has none. Each
 * host adds a shape in its default look, then sets these the way it sets
 * them on any shape.
 */
export function lookOf(edit: ShapeEdit): Partial<ShapeEdit> | undefined {
	const look: Partial<ShapeEdit> = {};
	if (edit.fill !== undefined) { look.fill = edit.fill; }
	if (edit.line !== undefined) { look.line = edit.line; }
	if (edit.font !== undefined) { look.font = edit.font; }
	if (edit.rotation !== undefined) { look.rotation = edit.rotation; }
	if (edit.hidden !== undefined) { look.hidden = edit.hidden; }
	if (edit.zOrder !== undefined) { look.zOrder = edit.zOrder; }
	return Object.keys(look).length > 0 ? look : undefined;
}

/** An addable AutoShape's DrawingML preset geometry. */
export type PresetShapeType = Exclude<NewShapeType, 'textBox' | 'button'>;

/**
 * The DrawingML geometry each addable AutoShape is written with, taken from
 * a file holding one of each that Excel and PowerPoint saved.
 *
 * A preset's name is not derivable from the MsoAutoShapeType constant, and
 * several belong to a different shape than the constant suggests:
 * msoShapeCross is `plus` and msoShapeRegularPentagon is `pentagon`, while
 * `cross` and `star5` are real presets for other shapes (`star5` is
 * msoShape5pointStar, 92). The table is measured, never derived.
 */
export const PRESET_GEOMETRY: Record<PresetShapeType, string> = {
	rectangle: 'rect',
	parallelogram: 'parallelogram',
	trapezoid: 'trapezoid',
	diamond: 'diamond',
	roundedRectangle: 'roundRect',
	octagon: 'octagon',
	triangle: 'triangle',
	rightTriangle: 'rtTriangle',
	oval: 'ellipse',
	hexagon: 'hexagon',
	cross: 'plus',
	pentagon: 'pentagon',
	cylinder: 'can',
	cube: 'cube',
	bevel: 'bevel',
	foldedCorner: 'foldedCorner',
	smileyFace: 'smileyFace',
	donut: 'donut',
	noEntry: 'noSmoking',
	blockArc: 'blockArc',
	star: 'star5',
};

/**
 * What a host names a new shape before anyone renames it, and the number it
 * appends: the shape's drawing id less one, counting across kinds.
 *
 * All three hosts WRITE the same labels, checked by saving one of each from
 * each. The object model is not the authority: Excel and PowerPoint report
 * a legacy name through `Shape.Name` for 8 of the 21 - "Rounded Rectangle",
 * "Can", "Donut", "5-Point Star" - while the files they save say
 * "Rectangle: Rounded Corners", "Cylinder", "Circle: Hollow" and "Star: 5
 * Points". Word's object model reports the written names. XLIDE reads and
 * writes the file, so the file's spelling is the one used.
 */
export const PRESET_LABELS: Record<PresetShapeType, string> = {
	rectangle: 'Rectangle',
	parallelogram: 'Parallelogram',
	trapezoid: 'Trapezoid',
	diamond: 'Diamond',
	roundedRectangle: 'Rectangle: Rounded Corners',
	octagon: 'Octagon',
	triangle: 'Isosceles Triangle',
	rightTriangle: 'Right Triangle',
	oval: 'Oval',
	hexagon: 'Hexagon',
	cross: 'Cross',
	pentagon: 'Pentagon',
	cylinder: 'Cylinder',
	cube: 'Cube',
	bevel: 'Rectangle: Beveled',
	foldedCorner: 'Rectangle: Folded Corner',
	smileyFace: 'Smiley Face',
	donut: 'Circle: Hollow',
	noEntry: '"Not Allowed" Symbol',
	blockArc: 'Block Arc',
	star: 'Star: 5 Points',
};

// --------------------------------------------------------- DrawingML text

/**
 * The text of a DrawingML text body, one line per paragraph. `bodyName` is
 * the host's element: xdr:txBody in a worksheet drawing, p:txBody on a
 * slide. Word keeps its shape text as WordprocessingML instead.
 */
export function drawingTextOf(xml: string, element: Span, bodyName: string): string | undefined {
	const body = findElement(xml, bodyName, element.openEnd, element.end);
	if (!body) { return undefined; }
	const paragraphs = [...xml.slice(body.openEnd, body.end).matchAll(/<a:p>([\s\S]*?)<\/a:p>|<a:p\/>/g)];
	return paragraphs
		.map((p) => [...(p[1] ?? '').matchAll(/<a:t>([\s\S]*?)<\/a:t>|<a:t\/>/g)].map((t) => decodeXml(t[1] ?? '')).join(''))
		.join('\n');
}

/**
 * A DrawingML text body with its paragraphs replaced by `text`, keeping the
 * first run's character and paragraph formatting so re-typing a caption does
 * not restyle it. Text typed into an empty shape takes the format of its
 * paragraph mark, as it does in the application, and failing that
 * `defaultRun`: Excel writes the 11-point size into every run, while a
 * PowerPoint run leaves its size to the slide. An empty line becomes an
 * empty paragraph carrying the same formatting, which is what Office writes
 * for one.
 */
export function withDrawingText(body: string, text: string, bodyName: string, defaultRun = '<a:rPr lang="en-US" sz="1100"/>'): string {
	const mark = /<a:endParaRPr\b[^>]*\/>|<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/.exec(body)?.[0];
	const rPr = /<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/.exec(body)?.[0]
		?? mark?.replace(/^<a:endParaRPr\b/, '<a:rPr').replace(/<\/a:endParaRPr>$/, '</a:rPr>')
		?? defaultRun;
	const pPr = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/.exec(body)?.[0] ?? '';
	const endRPr = rPr.replace(/^<a:rPr\b/, '<a:endParaRPr').replace(/<\/a:rPr>$/, '</a:endParaRPr>');
	const paragraphs = text.split(/\r?\n/).map((line) => (line
		? `<a:p>${pPr}<a:r>${rPr}<a:t>${encodeXml(line)}</a:t></a:r></a:p>`
		: `<a:p>${pPr}${endRPr}</a:p>`)).join('');
	const lstStyle = /<a:lstStyle\/>|<a:lstStyle>[\s\S]*?<\/a:lstStyle>/.exec(body);
	const keepUntil = lstStyle ? lstStyle.index + lstStyle[0].length : body.indexOf('>') + 1;
	return `${body.slice(0, keepUntil)}${paragraphs}</${bodyName}>`;
}
