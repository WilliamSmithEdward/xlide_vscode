// The shape editor's form: the fields a shape has, the values they start
// from, and the edit a Save makes.
//
// The webview only shows and collects values. The edit is worked out here
// from what changed, so an untouched field is never written back: a color a
// shape takes from its theme stays the theme's, and a position shown rounded
// to hundredths of a point is not rewritten because it was shown. Which
// fields a shape has comes from the same tables its host's writer refuses
// by (vba/shapeCapabilities.ts), so the form offers nothing the writer turns
// down.

import {
	PRESET_LABELS,
	Z_ORDER_COMMANDS,
	type FillEdit,
	type FontEdit,
	type LineEdit,
	type NewShapeType,
	type ShapeDash,
	type ShapeEdit,
	type ShapeInfo,
	type ShapeKind,
	type ZOrderCommand,
} from './vba/shapes';
import { SHAPE_DASHES } from './vba/shapeFormat';
import type { ShapeMacro } from './vba/projectService';
import {
	SHEET_LINKED,
	SHEET_LISTED,
	canFormat,
	canRunMacro,
	holdsText,
	type ShapeHost,
} from './vba/shapeCapabilities';

/** Every value the form holds, as its inputs hold them: text, and checkboxes. */
export interface ShapeFormValues {
	/** Add only: what to add. */
	type: string;
	/** Add in Word only: the story the shape goes on. */
	surface: string;
	name: string;
	altText: string;
	hidden: boolean;
	range: string;
	left: string;
	top: string;
	width: string;
	height: string;
	text: string;
	macro: string;
	linkedCell: string;
	inputRange: string;
	rotation: string;
	zOrder: '' | ZOrderCommand;
	/** auto: the shape's style decides; other: a gradient, picture or pattern, kept as it is. */
	fillType: 'auto' | 'none' | 'solid' | 'other';
	fillColor: string;
	fillTransparency: string;
	lineType: 'auto' | 'none' | 'solid' | 'other';
	lineColor: string;
	lineWeight: string;
	lineDash: string;
	fontName: string;
	fontSize: string;
	fontBold: boolean;
	fontItalic: boolean;
	fontUnderline: boolean;
	fontColor: string;
	/** Takes the text's color back to the one the shape's style gives it. */
	fontColorFromStyle: boolean;
}

/**
 * How a shape is placed: by the cells it covers (Excel), by its box in
 * points, by its size alone (a Word shape inline with the text), or not at
 * all here (a shape in an Excel group moves with the group).
 */
export type ShapePositionField = 'range' | 'box' | 'size' | 'none';

/** Which parts of the form a shape has. */
export interface ShapeFormFields {
	position: ShapePositionField;
	/** Why the position is limited, when it is. */
	positionNote?: string;
	text: boolean;
	macro: boolean;
	/** Why there is no macro, where a user would look for one. */
	macroNote?: string;
	linkedCell: boolean;
	inputRange: boolean;
	rotation: boolean;
	zOrder: boolean;
	zOrderNote?: string;
	fill: boolean;
	line: boolean;
	font: boolean;
	delete: boolean;
}

/** Everything the editor shows, sent to the webview as it opens. */
export interface ShapeEditorModel {
	mode: 'add' | 'edit';
	host: ShapeHost;
	fileName: string;
	surface: string;
	/** Add in Word: the stories a shape can go on. */
	surfaces?: string[];
	/** Edit: the shape as the file lists it. */
	shape?: ShapeInfo;
	/** Edit: its kind, as the rows name it. */
	kindLabel?: string;
	inGroup: boolean;
	macros: ShapeMacro[];
	/** Edit: the fields the shape has. Add: see fieldsByType. */
	fields: ShapeFormFields;
	/** Add: the fields each type of new shape has, by type. */
	fieldsByType?: Record<string, ShapeFormFields>;
	/** Add: what can be added here, in the order offered. */
	types?: Array<{ value: NewShapeType; label: string }>;
	/** The values the form starts from, which a Save is compared with. */
	values: ShapeFormValues;
	/** Where a stacked shape stands: 1 is at the back, of this many. */
	stackCount?: number;
	dashes: readonly ShapeDash[];
}

export interface ShapeEditorTarget {
	host: ShapeHost;
	fileName: string;
	surface: string;
	surfaces?: string[];
	shape?: ShapeInfo;
	inGroup?: boolean;
	/** How many shapes stack on the shape's surface. */
	stackCount?: number;
}

/** The fields of a shape of this kind in this host. */
export function shapeFormFields(host: ShapeHost, kind: ShapeKind, where: { inGroup: boolean; inline?: boolean; adding?: boolean }): ShapeFormFields {
	let position: ShapePositionField;
	let positionNote: string | undefined;
	if (host === 'excel') {
		position = where.inGroup ? 'none' : 'range';
		if (where.inGroup) { positionNote = 'A shape in a group moves with the group.'; }
	} else if (host === 'word' && where.inline && !where.inGroup) {
		position = 'size';
		positionNote = 'Inline with the text, so Word places it with the text; set its size here.';
	} else {
		position = 'box';
		if (where.inGroup) {
			positionNote = host === 'word' ? 'In points inside its canvas.' : 'In points on the slide.';
		} else if (host === 'word' && where.adding) {
			positionNote = 'Give left and top to make it float; leave them empty to put it inline with the text.';
		}
	}
	const macro = canRunMacro(host, kind);
	let macroNote: string | undefined;
	if (host === 'word') {
		macroNote = 'Word cannot run a macro from a shape: its shapes have no OnAction and no action settings.';
	} else if (kind === 'group') {
		macroNote = 'A group runs no macro; give the macro to a shape in it.';
	}
	const stacks = !where.inGroup && !(host === 'word' && where.inline) && !where.adding;
	return {
		position,
		...(positionNote ? { positionNote } : {}),
		text: holdsText(host, kind),
		macro,
		...(macroNote ? { macroNote } : {}),
		linkedCell: host === 'excel' && SHEET_LINKED.includes(kind),
		inputRange: host === 'excel' && SHEET_LISTED.includes(kind),
		rotation: canFormat(host, kind, 'rotation'),
		zOrder: stacks,
		...(!stacks && !where.adding
			? { zOrderNote: where.inGroup ? 'A shape in a group stacks with the group.' : 'A shape inline with the text does not stack.' }
			: {}),
		fill: canFormat(host, kind, 'fill'),
		line: canFormat(host, kind, 'line'),
		font: canFormat(host, kind, 'font'),
		delete: !where.inGroup && !where.adding,
	};
}

/** A number as a field shows it: up to two decimals, and empty for none. */
function shown(value: number | undefined): string {
	return value === undefined ? '' : String(Math.round(value * 100) / 100);
}

/** The values the form starts from: the shape's own, or a new shape's. */
export function shapeFormValues(shape: ShapeInfo | undefined, defaults: { type?: string; surface?: string } = {}): ShapeFormValues {
	const fill = shape?.fill;
	const line = shape?.line;
	const font = shape?.font;
	const typeOf = (look: { type: string; automatic?: boolean } | undefined): 'auto' | 'none' | 'solid' | 'other' => {
		if (!look || look.automatic) { return 'auto'; }
		return look.type === 'none' || look.type === 'solid' ? look.type : 'other';
	};
	return {
		type: defaults.type ?? '',
		surface: defaults.surface ?? '',
		name: shape?.name ?? '',
		altText: shape?.altText ?? '',
		hidden: shape?.hidden === true,
		range: shape?.range ?? '',
		left: shown(shape?.left),
		top: shown(shape?.top),
		width: shown(shape?.width),
		height: shown(shape?.height),
		text: shape?.text ?? '',
		macro: shape?.macro ?? '',
		linkedCell: shape?.linkedCell ?? '',
		inputRange: shape?.inputRange ?? '',
		rotation: shown(shape?.rotation ?? 0),
		zOrder: '',
		fillType: typeOf(fill),
		fillColor: fill?.color ?? '#FFFFFF',
		fillTransparency: shown(fill?.transparency ?? 0),
		lineType: typeOf(line),
		lineColor: line?.color ?? '#000000',
		lineWeight: shown(line?.weight),
		lineDash: line?.dash ?? 'solid',
		fontName: font?.name ?? '',
		fontSize: shown(font?.size),
		fontBold: font?.bold === true,
		fontItalic: font?.italic === true,
		fontUnderline: font?.underline === true,
		fontColor: font?.color ?? '#000000',
		fontColorFromStyle: false,
	};
}

/** The shapes each host can add, as the add form offers them. */
function addableTypes(host: ShapeHost): Array<{ value: NewShapeType; label: string }> {
	const shapes = (Object.keys(PRESET_LABELS) as Array<keyof typeof PRESET_LABELS>)
		.map((value) => ({ value: value as NewShapeType, label: PRESET_LABELS[value] }));
	return [
		...shapes,
		{ value: 'textBox', label: 'Text Box' },
		...(host === 'excel' ? [{ value: 'button' as NewShapeType, label: 'Button (form control)' }] : []),
	];
}

function kindOfNew(type: string): ShapeKind {
	return type === 'textBox' ? 'textBox' : type === 'button' ? 'button' : 'shape';
}

/** The model the editor opens with. */
export function shapeEditorModel(target: ShapeEditorTarget, macros: readonly ShapeMacro[], kindLabel?: string): ShapeEditorModel {
	const inGroup = target.inGroup === true;
	const base = {
		host: target.host,
		fileName: target.fileName,
		surface: target.surface,
		inGroup,
		macros: [...macros],
		dashes: SHAPE_DASHES,
		...(target.stackCount !== undefined ? { stackCount: target.stackCount } : {}),
	};
	if (!target.shape) {
		const types = addableTypes(target.host);
		const fieldsByType: Record<string, ShapeFormFields> = {};
		for (const type of types) {
			fieldsByType[type.value] = shapeFormFields(target.host, kindOfNew(type.value), { inGroup: false, adding: true });
		}
		return {
			...base,
			mode: 'add',
			...(target.surfaces ? { surfaces: target.surfaces } : {}),
			types,
			fields: fieldsByType[types[0].value],
			fieldsByType,
			values: shapeFormValues(undefined, { type: types[0].value, surface: target.surface }),
		};
	}
	return {
		...base,
		mode: 'edit',
		shape: target.shape,
		...(kindLabel ? { kindLabel } : {}),
		fields: shapeFormFields(target.host, target.shape.kind, { inGroup, inline: target.shape.placement === 'inline' }),
		values: shapeFormValues(target.shape),
	};
}

/** What a Save gives: the edit to make, or what is wrong with the form. */
export interface ShapeFormResult {
	/** Undefined when nothing changed. */
	edit?: ShapeEdit;
	/** The surface the edit is made on. */
	surface: string;
	/** Problems by field, each said so the user can fix it. */
	errors: Record<string, string>;
}

const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;

/**
 * The edit a Save makes: every field that differs from what the form
 * started with, and nothing else. An add also takes what places the shape.
 */
export function shapeEditFromForm(model: ShapeEditorModel, values: ShapeFormValues): ShapeFormResult {
	const errors: Record<string, string> = {};
	const initial = model.values;
	const adding = model.mode === 'add';
	const fields = adding ? model.fieldsByType?.[values.type] ?? model.fields : model.fields;
	const edit: ShapeEdit = adding ? { action: 'add' } : { action: 'update', name: model.shape!.name };
	let changed = false;
	const set = <K extends keyof ShapeEdit>(key: K, value: ShapeEdit[K]): void => {
		edit[key] = value;
		changed = true;
	};

	/** A number field: undefined when empty, NaN (and an error) when it is not a number. */
	const number = (key: keyof ShapeFormValues, label: string, check?: (n: number) => string | undefined): number | undefined => {
		const text = String(values[key]).trim();
		if (!text) { return undefined; }
		const n = Number(text);
		if (!Number.isFinite(n)) {
			errors[key] = `${label} is a number; '${text}' is not one.`;
			return NaN;
		}
		const problem = check?.(n);
		if (problem) { errors[key] = problem; }
		return n;
	};
	const numberChanged = (key: keyof ShapeFormValues, label: string, check?: (n: number) => string | undefined): number | undefined => {
		const n = number(key, label, check);
		if (n === undefined || Number.isNaN(n)) { return undefined; }
		const before = String(initial[key]).trim();
		return before !== '' && Number(before) === n && !adding ? undefined : n;
	};
	const color = (key: keyof ShapeFormValues, label: string): string => {
		const value = String(values[key]).trim();
		if (!HEX_COLOR.test(value)) { errors[key] = `${label} is a color as #RRGGBB, such as #FF0000.`; }
		return value.toUpperCase();
	};
	const textChanged = (key: 'name' | 'altText' | 'range' | 'text' | 'macro' | 'linkedCell' | 'inputRange'): string | undefined => {
		const value = key === 'text' ? values.text.replace(/\r\n/g, '\n') : values[key].trim();
		return value !== (key === 'text' ? initial.text : initial[key].trim()) ? value : undefined;
	};

	// What is added, and where.
	if (adding) {
		if (!model.types?.some((type) => type.value === values.type)) {
			errors.type = 'Choose what to add.';
		}
		edit.type = values.type as NewShapeType;
		changed = true;
	}
	const name = values.name.trim();
	if (adding) {
		if (name) { edit.name = name; }
	} else if (!name) {
		errors.name = 'A shape needs a name.';
	} else if (name !== initial.name) {
		set('newName', name);
	}
	const altText = textChanged('altText');
	if (altText !== undefined) { set('altText', altText); }
	if (values.hidden !== initial.hidden) { set('hidden', values.hidden); }

	if (fields.position === 'range') {
		const range = values.range.trim();
		if (adding && !range) {
			errors.range = 'Give the cells the shape covers, such as B2:D4.';
		} else if (range && !/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(range)) {
			errors.range = `'${range}' is not a cell range; give one such as B2:D4.`;
		} else if (range && (adding || range.toUpperCase() !== initial.range.toUpperCase())) {
			set('range', range.toUpperCase());
		}
	} else if (fields.position === 'box' || fields.position === 'size') {
		const positive = (label: string) => (n: number): string | undefined => (n > 0 ? undefined : `${label} is above zero.`);
		const keys: Array<['left' | 'top' | 'width' | 'height', string, ((n: number) => string | undefined)?]> = [
			['width', 'The width', positive('The width')],
			['height', 'The height', positive('The height')],
		];
		if (fields.position === 'box') { keys.unshift(['left', 'Left'], ['top', 'Top']); }
		for (const [key, label, check] of keys) {
			const n = numberChanged(key, label, check);
			if (n !== undefined) { set(key, n); }
		}
		if (adding && model.host === 'powerpoint' && (edit.left === undefined || edit.top === undefined)) {
			errors.left = 'Give left and top, in points from the slide\'s top-left corner.';
		}
		if (adding && model.host === 'word' && (edit.left === undefined) !== (edit.top === undefined)) {
			errors[edit.left === undefined ? 'left' : 'top'] = 'Give both left and top to make the shape float, or neither.';
		}
	}

	if (fields.text) {
		const text = textChanged('text');
		if (text !== undefined && (!adding || text)) { set('text', text); }
	}
	if (fields.macro) {
		const macro = textChanged('macro');
		if (macro !== undefined && (!adding || macro)) { set('macro', macro); }
	}
	if (fields.linkedCell) {
		const linked = textChanged('linkedCell');
		if (linked !== undefined) { set('linkedCell', linked); }
	}
	if (fields.inputRange) {
		const listed = textChanged('inputRange');
		if (listed !== undefined) { set('inputRange', listed); }
	}
	if (fields.rotation) {
		const rotation = numberChanged('rotation', 'The rotation');
		if (rotation !== undefined && !(adding && rotation === 0)) { set('rotation', rotation); }
	}
	if (fields.zOrder && values.zOrder) {
		if (!Z_ORDER_COMMANDS.includes(values.zOrder)) {
			errors.zOrder = 'Choose front, back, forward or backward.';
		} else {
			set('zOrder', values.zOrder);
		}
	}

	if (fields.fill) {
		const fill = fillEdit(values, initial, color, number, errors);
		if (fill) { set('fill', fill); }
	}
	if (fields.line) {
		const line = lineEdit(values, initial, color, number, errors);
		if (line) { set('line', line); }
	}
	if (fields.font) {
		const font: FontEdit = {};
		const fontName = values.fontName.trim();
		if (fontName && fontName !== initial.fontName.trim()) { font.name = fontName; }
		const size = numberChanged('fontSize', 'The font size', (n) => (n >= 1 && n <= 409 ? undefined : 'A font size is from 1 to 409 points.'));
		if (size !== undefined) { font.size = size; }
		for (const [key, field] of [['bold', 'fontBold'], ['italic', 'fontItalic'], ['underline', 'fontUnderline']] as const) {
			if (values[field] !== initial[field]) { font[key] = values[field]; }
		}
		if (values.fontColorFromStyle) {
			font.color = '';
		} else if (values.fontColor.trim().toUpperCase() !== initial.fontColor.trim().toUpperCase()) {
			font.color = color('fontColor', 'The font color');
		}
		if (Object.keys(font).length > 0) { set('font', font); }
	}

	const surface = adding && model.surfaces ? values.surface || model.surface : model.surface;
	if (Object.keys(errors).length > 0) { return { surface, errors }; }
	return { ...(changed ? { edit } : {}), surface, errors };
}

type ColorOf = (key: keyof ShapeFormValues, label: string) => string;
type NumberOf = (key: keyof ShapeFormValues, label: string, check?: (n: number) => string | undefined) => number | undefined;

function fillEdit(values: ShapeFormValues, initial: ShapeFormValues, color: ColorOf, number: NumberOf, errors: Record<string, string>): FillEdit | undefined {
	if (values.fillType === 'none') {
		return initial.fillType === 'none' ? undefined : { type: 'none' };
	}
	if (values.fillType !== 'solid') { return undefined; }
	const changed = initial.fillType !== 'solid'
		|| values.fillColor.trim().toUpperCase() !== initial.fillColor.trim().toUpperCase()
		|| Number(values.fillTransparency || 0) !== Number(initial.fillTransparency || 0);
	if (!changed) { return undefined; }
	const fillColor = color('fillColor', 'The fill color');
	const transparency = number('fillTransparency', 'The transparency',
		(n) => (n >= 0 && n <= 100 ? undefined : 'Transparency is a percentage from 0 to 100.'));
	if (errors.fillColor || errors.fillTransparency) { return undefined; }
	return { type: 'solid', color: fillColor, ...(transparency ? { transparency } : {}) };
}

function lineEdit(values: ShapeFormValues, initial: ShapeFormValues, color: ColorOf, number: NumberOf, errors: Record<string, string>): LineEdit | undefined {
	if (values.lineType === 'none') {
		return initial.lineType === 'none' ? undefined : { type: 'none' };
	}
	if (values.lineType !== 'solid') { return undefined; }
	const turnedOn = initial.lineType !== 'solid';
	const out: LineEdit = { type: 'solid' };
	if (turnedOn || values.lineColor.trim().toUpperCase() !== initial.lineColor.trim().toUpperCase()) {
		out.color = color('lineColor', 'The outline color');
	}
	const weight = number('lineWeight', 'The outline weight',
		(n) => (n > 0 && n <= 1584 ? undefined : 'An outline weight is above 0 and up to 1584 points.'));
	if (weight !== undefined && !Number.isNaN(weight) && (turnedOn || weight !== Number(initial.lineWeight))) {
		out.weight = weight;
	}
	if (turnedOn || values.lineDash !== initial.lineDash) {
		if (!(SHAPE_DASHES as readonly string[]).includes(values.lineDash)) {
			errors.lineDash = 'Choose a dash style.';
		} else {
			out.dash = values.lineDash as ShapeDash;
		}
	}
	if (errors.lineColor || errors.lineWeight || errors.lineDash) { return undefined; }
	return turnedOn || Object.keys(out).length > 1 ? out : undefined;
}
