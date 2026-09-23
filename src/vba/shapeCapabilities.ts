// What each host lets a kind of shape be given.
//
// One table per host, read by the host's writer to refuse what it cannot
// write and by the shape editor to show only the fields a shape has, so the
// two cannot disagree. Each table was drawn from what the application
// itself offers for that kind of shape: Excel keeps a form control's look in
// VML and gives a font only to a button's caption, PowerPoint turns a
// picture and a line, Word neither groups nor rotates a drawing canvas.

import type { ShapeKind } from './shapes';

export type ShapeHost = 'excel' | 'word' | 'powerpoint';

export type FormatProperty = 'fill' | 'line' | 'font' | 'rotation';

/** A worksheet's DrawingML shapes. */
export const SHEET_FORMATTABLE: Record<FormatProperty, readonly ShapeKind[]> = {
	fill: ['shape', 'textBox'],
	line: ['shape', 'textBox', 'line', 'picture'],
	font: ['shape', 'textBox'],
	rotation: ['shape', 'textBox', 'line', 'picture', 'group'],
};

/** Excel's form controls, whose look is VML and which Excel does not rotate. */
export const SHEET_CONTROLS: readonly ShapeKind[] = [
	'button', 'checkBox', 'optionButton', 'dropDown', 'listBox', 'scrollBar', 'spinner',
	'label', 'groupBox', 'editBox', 'formControl',
];

/** Of the form controls: the one whose caption takes a font, and the two that take a fill and a line. */
export const SHEET_CONTROL_FORMAT: Record<'font' | 'fill', readonly ShapeKind[]> = {
	font: ['button'],
	fill: ['checkBox', 'optionButton'],
};

/** The form controls with a caption. */
export const SHEET_CAPTIONED: readonly ShapeKind[] = ['button', 'checkBox', 'optionButton', 'label', 'groupBox'];

/** The form controls with a cell link. */
export const SHEET_LINKED: readonly ShapeKind[] = ['checkBox', 'optionButton', 'dropDown', 'listBox', 'scrollBar', 'spinner'];

/** The form controls that list the cells of an input range. */
export const SHEET_LISTED: readonly ShapeKind[] = ['dropDown', 'listBox'];

/** A slide's shapes. */
export const SLIDE_FORMATTABLE: Record<FormatProperty, readonly ShapeKind[]> = {
	fill: ['shape', 'textBox', 'placeholder'],
	line: ['shape', 'textBox', 'placeholder', 'line', 'picture'],
	font: ['shape', 'textBox', 'placeholder'],
	rotation: ['shape', 'textBox', 'placeholder', 'line', 'picture', 'group'],
};

/** A Word document's shapes. */
export const STORY_FORMATTABLE: Record<FormatProperty, readonly ShapeKind[]> = {
	fill: ['shape', 'textBox'],
	line: ['shape', 'textBox', 'picture'],
	font: ['shape', 'textBox'],
	rotation: ['shape', 'textBox', 'picture', 'group'],
};

/** Whether a host sets this look property on this kind of shape. */
export function canFormat(host: ShapeHost, kind: ShapeKind, property: FormatProperty): boolean {
	switch (host) {
		case 'excel':
			if (SHEET_CONTROLS.includes(kind)) {
				if (property === 'rotation') { return false; }
				return property === 'font' ? SHEET_CONTROL_FORMAT.font.includes(kind) : SHEET_CONTROL_FORMAT.fill.includes(kind);
			}
			return SHEET_FORMATTABLE[property].includes(kind);
		case 'powerpoint':
			return SLIDE_FORMATTABLE[property].includes(kind);
		default:
			return STORY_FORMATTABLE[property].includes(kind);
	}
}

/** Whether a kind of shape holds text a host sets: a shape's words, or a form control's caption. */
export function holdsText(host: ShapeHost, kind: ShapeKind): boolean {
	if (kind === 'shape' || kind === 'textBox') { return true; }
	if (host === 'excel') { return SHEET_CAPTIONED.includes(kind); }
	return host === 'powerpoint' && kind === 'placeholder';
}

/**
 * Whether a click on the shape can run a macro in its host. Word runs none
 * from any shape; Excel and PowerPoint run none from a group (a shape in it
 * takes the macro); an ActiveX control runs event procedures instead.
 */
export function canRunMacro(host: ShapeHost, kind: ShapeKind): boolean {
	return host !== 'word' && kind !== 'group' && kind !== 'canvas' && kind !== 'activeX';
}
