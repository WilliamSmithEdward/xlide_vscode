// A VB6 form's designer header, read into the scene the form canvas draws.
//
// The header (frmHeader.ts) is the designer's own text: every control's
// class, name, and properties in twips, nested in its container. This turns
// that into the same FormScene an MSForms package becomes (oforms/preview.ts),
// so a VB6 form gets the canvas, the script, and the gestures the OFORMS
// designer already has. It draws what the header honestly says: bounds,
// captions, colors and fonts, a Line between its two points, a Shape in
// its shape; a picture only when the sidecar bytes are an image a browser
// paints; a control from an OCX as its true bounds and identity.
//
// Units: VB6 stores design-time geometry in twips (1440 per inch); the
// canvas speaks points (72 per inch), so a twip is a twentieth of a point.

import type { FormScene, SceneControl } from '../oforms/preview';
import { cssColor, esc, imageDataUri, sceneControl } from '../oforms/preview';
import { formatOleColor, parseOleColor } from '../oforms/markup';
import { getVb6ObjectModel } from '../../analyzer/host/vb6ObjectModel';
import type { HostConstant } from '../../analyzer/host/excelObjectModel';
import { frmProperty } from './frmHeader';
import type { FrmControl, FrmHeader, FrmProperty, FrmPropertyGroup, FrxRef } from './frmHeader';
import type { FrxValue } from './frx';

/** Sidecar lookup by reference, when the `.frx` was read. */
export type FrxLookup = (ref: FrxRef) => FrxValue | undefined;

export interface FrmSceneOptions {
	formName: string;
	frx?: FrxLookup;
}

/**
 * What the designer knows about one intrinsic kind. One table carries
 * everything the canvas, the toolbox, the pane and the gestures ask about a
 * kind, so a kind is described once and the answers cannot drift apart.
 */
export interface Vb6ControlSpec {
	/** The canvas kind: Label, TextBox, ScrollBar, ... or the designer's own class (Form, UserControl). */
	kind: string;
	/** A designer class (Form, MDIForm, UserControl, PropertyPage) rather than a control on it. */
	designer?: boolean;
	/** The designer's base name for a new one (Command1, Text1); absent for a kind the toolbox does not add. */
	base?: string;
	/** The size a new one takes, in twips: the toolbox defaults as `.frm` files commonly show them. A Timer and a Line have none. */
	size?: { width: number; height: number };
	/** The property that repeats a new control's name. */
	text?: 'Caption' | 'Text';
	/** Takes a TabIndex. */
	tab?: boolean;
	/** Writes its scale, the client area inside its border, when added. */
	scale?: boolean;
	/** Holds controls. */
	container?: boolean;
	/** Shows its Caption on the canvas. */
	captioned?: boolean;
	/**
	 * The event a double-click opens. Absent means Click, which the model
	 * must declare for the kind; `''` means the kind has no event a
	 * double-click should write.
	 */
	defaultEvent?: string;
	/**
	 * Where `designProperties` comes from. `fixtures`: the union of the
	 * fixture designers' own headers, which is what VB6 itself wrote.
	 * `model`: the `VB` model's property list for the kind, narrowed to the
	 * ones VB6's Properties window offers, because no fixture uses the kind.
	 * The second is inferred rather than measured, and a fixture that uses
	 * such a kind should replace it.
	 */
	vocabularyFrom: 'fixtures' | 'model';
	/**
	 * The design-time properties the designer writes for the kind. The pane
	 * lists these beside what a header states, blank until set; `Font` stands
	 * for the Font group's own rows. Geometry and Index have rows of their
	 * own, and a property kept in the `.frx` (a picture, a list) is listed
	 * only when the header states it, because the designer will not write one.
	 * Every name is checked against the model, bar the keys only the designer
	 * writes (a form's client position, a UserControl's toolbox bitmap).
	 */
	designProperties: readonly string[];
}

const FONT = 'Font';

/** The intrinsic kinds, keyed by prog id, in the toolbox's order. */
export const VB6_CONTROLS: Readonly<Record<string, Vb6ControlSpec>> = {
	'VB.Form': {
		kind: 'Form', designer: true, defaultEvent: 'Load', vocabularyFrom: 'fixtures',
		designProperties: ['AutoRedraw', 'BackColor', 'BorderStyle', 'Caption', 'ClientLeft', 'ClientTop', 'ControlBox', FONT, 'ForeColor',
			'Icon', 'KeyPreview', 'LinkTopic', 'MaxButton', 'MinButton', 'Picture', 'ScaleHeight', 'ScaleMode', 'ScaleWidth',
			'ShowInTaskbar', 'StartUpPosition', 'Tag'],
	},
	'VB.MDIForm': { kind: 'MDIForm', designer: true, defaultEvent: 'Load', vocabularyFrom: 'fixtures', designProperties: [] },
	'VB.UserControl': {
		kind: 'UserControl', designer: true, vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'BackStyle', 'ClientLeft', 'ClientTop', 'ScaleHeight', 'ScaleWidth', 'ToolboxBitmap'],
	},
	'VB.PropertyPage': { kind: 'PropertyPage', designer: true, vocabularyFrom: 'fixtures', designProperties: [] },
	'VB.Label': {
		kind: 'Label', base: 'Label', size: { width: 1215, height: 255 }, text: 'Caption', tab: true, captioned: true,
		vocabularyFrom: 'fixtures',
		designProperties: ['Alignment', 'AutoSize', 'BackColor', 'BackStyle', 'Caption', FONT, 'ForeColor', 'MouseIcon', 'MousePointer',
			'TabIndex', 'Tag', 'ToolTipText', 'UseMnemonic', 'Visible', 'WordWrap'],
	},
	'VB.TextBox': {
		kind: 'TextBox', base: 'Text', size: { width: 1215, height: 285 }, text: 'Text', tab: true, vocabularyFrom: 'fixtures',
		designProperties: ['Alignment', 'Appearance', 'BackColor', 'BorderStyle', 'Enabled', FONT, 'ForeColor', 'Locked', 'MultiLine',
			'ScrollBars', 'TabIndex', 'Tag', 'Text', 'ToolTipText', 'Visible'],
	},
	'VB.ComboBox': {
		kind: 'ComboBox', base: 'Combo', size: { width: 1215, height: 315 }, text: 'Text', tab: true, vocabularyFrom: 'fixtures',
		designProperties: ['Appearance', FONT, 'Style', 'TabIndex', 'TabStop', 'Tag'],
	},
	'VB.ListBox': {
		kind: 'ListBox', base: 'List', size: { width: 1215, height: 1035 }, tab: true, vocabularyFrom: 'model',
		designProperties: ['Appearance', 'BackColor', 'BorderStyle', 'Columns', 'Enabled', FONT, 'ForeColor', 'IntegralHeight',
			'MultiSelect', 'Sorted', 'Style', 'TabIndex', 'TabStop', 'Tag', 'ToolTipText', 'Visible'],
	},
	'VB.CheckBox': {
		kind: 'CheckBox', base: 'Check', size: { width: 1215, height: 255 }, text: 'Caption', tab: true, captioned: true,
		vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'Caption', FONT, 'TabIndex', 'Value'],
	},
	'VB.OptionButton': {
		kind: 'OptionButton', base: 'Option', size: { width: 1215, height: 255 }, text: 'Caption', tab: true, captioned: true,
		vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'Caption', 'TabIndex', 'Value'],
	},
	'VB.CommandButton': {
		kind: 'CommandButton', base: 'Command', size: { width: 1215, height: 495 }, text: 'Caption', tab: true, captioned: true,
		vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'Caption', 'Default', 'Enabled', FONT, 'MaskColor', 'Style', 'TabIndex', 'TabStop'],
	},
	'VB.Frame': {
		kind: 'Frame', base: 'Frame', size: { width: 1215, height: 1215 }, text: 'Caption', tab: true, container: true, captioned: true,
		vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'BorderStyle', 'Caption', FONT, 'TabIndex'],
	},
	'VB.PictureBox': {
		kind: 'PictureBox', base: 'Picture', size: { width: 1215, height: 1215 }, tab: true, scale: true, container: true,
		defaultEvent: 'Click', vocabularyFrom: 'fixtures',
		designProperties: ['Align', 'Appearance', 'AutoRedraw', 'BackColor', 'BorderStyle', FONT, 'ForeColor', 'ScaleHeight',
			'ScaleMode', 'ScaleWidth', 'TabIndex', 'TabStop', 'Tag', 'ToolTipText', 'Visible'],
	},
	'VB.Image': {
		kind: 'Image', base: 'Image', size: { width: 1215, height: 1215 }, vocabularyFrom: 'model',
		designProperties: ['Appearance', 'BorderStyle', 'Enabled', 'Stretch', 'Tag', 'ToolTipText', 'Visible'],
	},
	// A scroll bar raises Change, never Click: a double-click that wrote
	// HScroll1_Click would name an event VB6 does not have.
	'VB.HScrollBar': {
		kind: 'ScrollBar', base: 'HScroll', size: { width: 1215, height: 255 }, tab: true, defaultEvent: 'Change',
		vocabularyFrom: 'model',
		designProperties: ['Enabled', 'LargeChange', 'Max', 'Min', 'SmallChange', 'TabIndex', 'TabStop', 'Tag', 'Value', 'Visible'],
	},
	'VB.VScrollBar': {
		kind: 'ScrollBar', base: 'VScroll', size: { width: 255, height: 1215 }, tab: true, defaultEvent: 'Change',
		vocabularyFrom: 'model',
		designProperties: ['Enabled', 'LargeChange', 'Max', 'Min', 'SmallChange', 'TabIndex', 'TabStop', 'Tag', 'Value', 'Visible'],
	},
	'VB.Timer': {
		kind: 'Timer', base: 'Timer', defaultEvent: 'Timer', vocabularyFrom: 'model',
		designProperties: ['Enabled', 'Interval', 'Tag'],
	},
	// The model gives Line and Shape an Initialize event; VB6 gives them no
	// events at all, and the model's events are the one part the transcription
	// could not cross-read against Microsoft's reference
	// (docs/vb6_reference_data.md), so the designer writes neither.
	'VB.Line': { kind: 'Line', base: 'Line', defaultEvent: '', vocabularyFrom: 'fixtures', designProperties: ['BorderColor', 'BorderStyle'] },
	'VB.Shape': {
		kind: 'Shape', base: 'Shape', size: { width: 1215, height: 1215 }, defaultEvent: '', vocabularyFrom: 'fixtures',
		designProperties: ['BackColor', 'BackStyle', 'BorderColor', 'BorderWidth', 'FillColor', 'Shape', 'Visible'],
	},
	'VB.DriveListBox': {
		kind: 'DriveListBox', base: 'Drive', size: { width: 1215, height: 315 }, tab: true, defaultEvent: 'Change',
		vocabularyFrom: 'model',
		designProperties: ['Appearance', 'BackColor', 'Enabled', FONT, 'ForeColor', 'TabIndex', 'TabStop', 'Tag', 'ToolTipText', 'Visible'],
	},
	'VB.DirListBox': {
		kind: 'DirListBox', base: 'Dir', size: { width: 1215, height: 1035 }, tab: true, defaultEvent: 'Change',
		vocabularyFrom: 'model',
		designProperties: ['Appearance', 'BackColor', 'BorderStyle', 'Enabled', FONT, 'ForeColor', 'IntegralHeight', 'TabIndex',
			'TabStop', 'Tag', 'ToolTipText', 'Visible'],
	},
	'VB.FileListBox': {
		kind: 'FileListBox', base: 'File', size: { width: 1215, height: 1035 }, tab: true, vocabularyFrom: 'model',
		designProperties: ['Appearance', 'Archive', 'BackColor', 'BorderStyle', 'Enabled', FONT, 'ForeColor', 'Hidden',
			'IntegralHeight', 'MultiSelect', 'Normal', 'Pattern', 'ReadOnly', 'System', 'TabIndex', 'TabStop', 'Tag',
			'ToolTipText', 'Visible'],
	},
	// The Data control raises Reposition; it has no Click.
	'VB.Data': {
		kind: 'Data', base: 'Data', size: { width: 1755, height: 345 }, text: 'Caption', tab: true, captioned: true,
		defaultEvent: 'Reposition', vocabularyFrom: 'model',
		designProperties: ['Appearance', 'BackColor', 'BOFAction', 'Caption', 'Connect', 'DatabaseName', 'DefaultCursorType',
			'DefaultType', 'Enabled', 'EOFAction', 'Exclusive', FONT, 'ForeColor', 'Options', 'ReadOnly', 'RecordsetType',
			'RecordSource', 'TabIndex', 'TabStop', 'Tag', 'ToolTipText', 'Visible'],
	},
	'VB.OLE': {
		kind: 'OLE', base: 'OLE', size: { width: 1215, height: 1215 }, tab: true, vocabularyFrom: 'model',
		designProperties: ['Appearance', 'AutoActivate', 'AutoVerbMenu', 'BackColor', 'BackStyle', 'BorderStyle', 'Class',
			'DisplayType', 'Enabled', 'HostName', 'MiscFlags', 'OLEDropAllowed', 'OLETypeAllowed', 'SizeMode', 'SourceDoc',
			'SourceItem', 'TabIndex', 'TabStop', 'Tag', 'UpdateOptions', 'Visible'],
	},
};

/** The keys only the designer writes: the model has no property for them. */
export const VB6_DESIGNER_ONLY_KEYS: readonly string[] = ['ClientLeft', 'ClientTop', 'ToolboxBitmap'];

/** Properties whose value lives in the `.frx`: a picture the designer reads and never writes. */
export const VB6_SIDECAR_PICTURE_KEYS: ReadonlySet<string> = new Set([
	'picture', 'icon', 'mouseicon', 'dragicon', 'downpicture', 'disabledpicture', 'maskpicture', 'toolboxbitmap',
]);
/** Properties whose rows live in the `.frx`: a ListBox's or ComboBox's items. */
export const VB6_SIDECAR_LIST_KEYS: ReadonlySet<string> = new Set(['list', 'itemdata']);

/** The VB6 toolbox: the intrinsic controls a form can add, by name. */
export const VB6_TOOLBOX: readonly string[] = Object.entries(VB6_CONTROLS)
	.filter(([, spec]) => spec.base !== undefined)
	.map(([progId]) => progId.slice(3));

/** The design-time properties per prog id (see `Vb6ControlSpec.designProperties`). */
export const VB6_DESIGN_PROPERTIES: Readonly<Record<string, readonly string[]>> = Object.fromEntries(
	Object.entries(VB6_CONTROLS).map(([progId, spec]) => [progId, spec.designProperties]),
);

/** The Font group's members, in the order the designer writes them. */
export const VB6_FONT_FIELDS = ['Name', 'Size', 'Charset', 'Weight', 'Underline', 'Italic', 'Strikethrough'] as const;

/**
 * The gloss the designer writes after an enum value (`BorderStyle = 3
 * 'Fixed Dialog`), measured per kind, property and value on the fixture
 * headers. A write of a measured value carries its gloss; any other value
 * is written bare, which VB6 reads the same and re-glosses on its own save.
 */
export const VB6_ENUM_GLOSSES: Readonly<Record<string, Readonly<Record<string, string>>>> = {
	'VB.Form.BorderStyle': { '1': 'Fixed Single', '3': 'Fixed Dialog' },
	'VB.Line.BorderStyle': { '3': 'Dot' },
	'VB.Frame.BorderStyle': { '0': 'None' },
	'VB.TextBox.BorderStyle': { '0': 'None' },
	'VB.PictureBox.BorderStyle': { '0': 'None' },
	'VB.Label.Alignment': { '1': 'Right Justify', '2': 'Center' },
	'VB.TextBox.Alignment': { '1': 'Right Justify', '2': 'Center' },
	'VB.TextBox.Appearance': { '0': 'Flat' },
	'VB.PictureBox.Appearance': { '0': 'Flat' },
	'VB.ComboBox.Appearance': { '0': 'Flat' },
	'VB.Label.BackStyle': { '0': 'Transparent' },
	'VB.UserControl.BackStyle': { '0': 'Transparent' },
	'VB.Shape.BackStyle': { '1': 'Opaque' },
	'VB.PictureBox.Align': { '1': 'Align Top' },
	'VB.Label.MousePointer': { '99': 'Custom' },
	'VB.Form.ScaleMode': { '3': 'Pixel' },
	'VB.PictureBox.ScaleMode': { '3': 'Pixel' },
	'VB.TextBox.ScrollBars': { '2': 'Vertical' },
	'VB.Shape.Shape': { '3': 'Circle' },
	'VB.Form.StartUpPosition': { '1': 'CenterOwner', '2': 'CenterScreen', '3': 'Windows Default' },
	'VB.CommandButton.Style': { '1': 'Graphical' },
	'VB.ComboBox.Style': { '2': 'Dropdown List' },
	'VB.CheckBox.Value': { '1': 'Checked', '2': 'Grayed' },
};

export interface Vb6PaneVocabulary {
	/** Dropdown options per `Kind.Prop`: the constants of the enum the model declares for the property, by value. */
	enums: Record<string, [string, string][]>;
	/** The `Kind.Prop` names the model declares Boolean. */
	bools: string[];
}

let paneVocabulary: Vb6PaneVocabulary | undefined;

/**
 * What the pane offers for the design-time vocabulary, from the `VB` model:
 * a property declared as an enum whose constants the model holds gets a
 * dropdown of those constants (`Form.BorderStyle`: vbBSNone through
 * vbSizableToolWindow); a property declared Boolean gets True/False. An
 * enum the model declares but holds no constants for stays a text field
 * rather than a dropdown that could not offer every value.
 */
export function vb6PaneVocabulary(): Vb6PaneVocabulary {
	if (paneVocabulary) { return paneVocabulary; }
	const model = getVb6ObjectModel();
	const constantsByType = new Map<string, HostConstant[]>();
	for (const constant of Object.values(model.constants ?? {})) {
		if (!constant.type) { continue; }
		const list = constantsByType.get(constant.type) ?? [];
		list.push(constant);
		constantsByType.set(constant.type, list);
	}
	const enums: Record<string, [string, string][]> = {};
	const bools: string[] = [];
	for (const [progId, keys] of Object.entries(VB6_DESIGN_PROPERTIES)) {
		const type = model.types[progId];
		if (!type) { continue; }
		const kind = vb6CanvasKind(progId) ?? progId.slice(3);
		for (const key of keys) {
			if (key === 'Font') { continue; }
			const member = type.members.find((m) => m.kind === 'property' && m.name.toLowerCase() === key.toLowerCase());
			const declared = member?.declaredType;
			if (!declared) { continue; }
			if (declared === 'Boolean') {
				bools.push(`${kind}.${key}`);
				continue;
			}
			const constants = constantsByType.get(declared);
			if (constants && constants.length > 0) {
				enums[`${kind}.${key}`] = [...constants]
					.sort((a, b) => Number(a.value) - Number(b.value))
					.map((c) => [String(c.value), c.name]);
			}
		}
	}
	paneVocabulary = { enums, bools };
	return paneVocabulary;
}

/**
 * The event a double-click opens where VB6's differs from MSForms', by
 * canvas kind: a Form loads, a Timer ticks, a PictureBox clicks. A Line or
 * Shape has no events at all; the empty string is passed through to the
 * host, which refuses it by name.
 */
export const VB6_DEFAULT_EVENTS: Readonly<Record<string, string>> = Object.fromEntries(
	Object.values(VB6_CONTROLS)
		.filter((spec) => spec.defaultEvent !== undefined)
		.map((spec) => [spec.kind, spec.defaultEvent as string]),
);

/** The spec of a control's prog id, or undefined for a designer class, a menu, an OCX or a custom control. */
export function vb6ControlSpec(progId: string): Vb6ControlSpec | undefined {
	const spec = VB6_CONTROLS[progId];
	return spec && !spec.designer ? spec : undefined;
}

/** The canvas kind for a VB6 prog id, or undefined for an OCX or custom control. */
export function vb6CanvasKind(progId: string): string | undefined {
	return vb6ControlSpec(progId)?.kind;
}

/** A Line's two points in twips; a far point the header omits sits on the near one. */
export function frmLineEnds(control: FrmControl): { x1: number; y1: number; x2: number; y2: number } {
	const x1 = frmNumberOf(frmProperty(control, 'X1')?.value) ?? 0;
	const y1 = frmNumberOf(frmProperty(control, 'Y1')?.value) ?? 0;
	return {
		x1, y1,
		x2: frmNumberOf(frmProperty(control, 'X2')?.value) ?? x1,
		y2: frmNumberOf(frmProperty(control, 'Y2')?.value) ?? y1,
	};
}

/** Twips to points, printed shortest: 240 -> "12", 250 -> "12.5". */
export function twipsToPt(twips: number): string {
	const pt = twips / 20;
	return String(Math.round(pt * 100) / 100);
}

/** A VB6 property value spelled as the designer writes it: a quoted string, a number, a color, a boolean gloss. */
export function unquoteVb6(value: string): string {
	const trimmed = value.trim();
	if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
		return trimmed.slice(1, -1).replace(/""/g, '"');
	}
	return trimmed;
}

/**
 * The number a header value starts with. The designer's boolean spelling
 * `0   'False` pads the zero to the width of `-1`, which the parser cannot
 * split from its gloss, so the value text carries the gloss along; the
 * number is its leading token either way.
 */
export function frmNumberOf(value: string | undefined): number | undefined {
	if (value === undefined) { return undefined; }
	const m = /^\s*(-?\d+(?:\.\d+)?)(?:\s|$)/.exec(value);
	return m ? Number(m[1]) : undefined;
}

const declaredTypes = new Map<string, Map<string, string | undefined>>();

/** The type the `VB` model declares for a property of a kind (`String`, `Boolean`, `OLE_COLOR`, an enum), or undefined when it does not know the kind or the key. */
export function vb6DeclaredType(progId: string, key: string): string | undefined {
	let table = declaredTypes.get(progId);
	if (!table) {
		table = new Map();
		const type = getVb6ObjectModel().types[progId];
		for (const m of type?.members ?? []) {
			if (m.kind === 'property') { table.set(m.name.toLowerCase(), m.declaredType); }
		}
		declaredTypes.set(progId, table);
	}
	return table.get(key.toLowerCase());
}

function group(control: FrmControl, name: string): FrmPropertyGroup | undefined {
	const lower = name.toLowerCase();
	for (const m of control.members) {
		if (m.kind === 'group' && m.name.toLowerCase() === lower) { return m; }
	}
	return undefined;
}

function groupValue(g: FrmPropertyGroup | undefined, key: string): string | undefined {
	if (!g) { return undefined; }
	const lower = key.toLowerCase();
	for (const m of g.members) {
		if (m.kind === 'property' && m.key.toLowerCase() === lower) { return m.value; }
	}
	return undefined;
}

/** A string property: inline, or from the sidecar when the header points there. */
function textOf(control: FrmControl, key: string, frx: FrxLookup | undefined): string | undefined {
	const p = frmProperty(control, key);
	if (!p) { return undefined; }
	if (p.frx) {
		const value = frx?.(p.frx);
		if (value && (value.kind === 'longString' || value.kind === 'shortString')) { return value.text; }
		return undefined;
	}
	return unquoteVb6(p.value);
}

function colorCssOf(control: FrmControl, key: string): string | undefined {
	const p = frmProperty(control, key);
	if (!p || p.frx) { return undefined; }
	const value = parseOleColor(p.value);
	return value === undefined ? undefined : cssColor(value);
}

/** True/False as VB6 writes them: -1 and 0, with the `'True` gloss beside. */
function boolOf(control: FrmControl, key: string): boolean | undefined {
	const n = frmNumberOf(frmProperty(control, key)?.value);
	return n === undefined ? undefined : n !== 0;
}

/** The form's own default font, and every control's when it says nothing. */
const DEFAULT_FONT_CSS = "font-family:'MS Sans Serif',Tahoma,sans-serif;font-size:8.25pt;";

function fontCssOf(control: FrmControl): string {
	const g = group(control, 'Font');
	if (!g) { return ''; }
	const parts: string[] = [];
	const name = groupValue(g, 'Name');
	if (name !== undefined) { parts.push(`font-family:'${esc(unquoteVb6(name))}',Tahoma,sans-serif;`); }
	const size = frmNumberOf(groupValue(g, 'Size'));
	if (size !== undefined) { parts.push(`font-size:${size}pt;`); }
	const weight = frmNumberOf(groupValue(g, 'Weight'));
	if (weight !== undefined && weight >= 600) { parts.push('font-weight:bold;'); }
	if ((frmNumberOf(groupValue(g, 'Italic')) ?? 0) !== 0) { parts.push('font-style:italic;'); }
	const deco = [
		(frmNumberOf(groupValue(g, 'Underline')) ?? 0) !== 0 ? 'underline' : '',
		(frmNumberOf(groupValue(g, 'Strikethrough')) ?? 0) !== 0 ? 'line-through' : '',
	].filter(Boolean).join(' ');
	if (deco) { parts.push(`text-decoration:${deco};`); }
	return parts.join('');
}

/**
 * The image inside a sidecar picture record, as a data URI the browser can
 * paint, or undefined when the record holds no picture or one no browser
 * shows (a metafile). Measured on the fixtures' `.frx` files: the payload is
 * an optional 16-byte StdPicture class id, then `lt` and two zero bytes,
 * then the image's byte length and the image itself - a BMP, ICO, or JPEG
 * file as it stands. An empty picture is `lt` followed by zeros.
 */
export function pictureDataUriOf(bytes: Buffer): string | undefined {
	const marker = [0, 16].find((at) => bytes.length >= at + 8 && bytes[at] === 0x6c && bytes[at + 1] === 0x74);
	if (marker === undefined) { return undefined; }
	const length = bytes.readUInt32LE(marker + 4);
	const image = bytes.subarray(marker + 8, marker + 8 + length);
	return length === 0 ? undefined : imageDataUri(image);
}

function pictureCssOf(control: FrmControl, key: string, frx: FrxLookup | undefined): string {
	const p = frmProperty(control, key);
	if (!p?.frx) { return ''; }
	const value = frx?.(p.frx);
	const bytes = value && (value.kind === 'picture' || value.kind === 'opaque') ? value.bytes : undefined;
	const uri = bytes ? pictureDataUriOf(bytes) : undefined;
	return uri ? `background-image:url('${uri}');background-repeat:no-repeat;background-position:center center;` : '';
}

const ALIGNMENT_CSS: Readonly<Record<number, string>> = { 1: 'text-align:right;justify-content:flex-end;', 2: 'text-align:center;justify-content:center;' };

/** The display name of a control: its name, with its Index when it is an array element. */
export function vb6ControlName(control: FrmControl): string {
	const index = frmProperty(control, 'Index');
	return index && !index.frx ? `${control.name}(${index.value.trim()})` : control.name;
}

function boundsOf(control: FrmControl, kind: string | undefined): { left: number; top: number; width: number; height: number } {
	if (kind === 'Line') {
		const { x1, y1, x2, y2 } = frmLineEnds(control);
		return {
			left: Math.min(x1, x2), top: Math.min(y1, y2),
			width: Math.max(Math.abs(x2 - x1), 20), height: Math.max(Math.abs(y2 - y1), 20),
		};
	}
	const left = frmNumberOf(frmProperty(control, 'Left')?.value) ?? 0;
	const top = frmNumberOf(frmProperty(control, 'Top')?.value) ?? 0;
	if (kind === 'Timer') {
		// A Timer has no size; the designer shows it as an icon.
		return { left, top, width: 480, height: 480 };
	}
	return {
		left, top,
		width: frmNumberOf(frmProperty(control, 'Width')?.value) ?? 0,
		height: frmNumberOf(frmProperty(control, 'Height')?.value) ?? 0,
	};
}

function sceneControlsOfChildren(children: readonly FrmControl[], frx: FrxLookup | undefined): SceneControl[] {
	const out: SceneControl[] = [];
	children.forEach((control, index) => {
		if (control.progId === 'VB.Menu') { return; }
		const spec = vb6ControlSpec(control.progId);
		const kind = spec?.kind;
		const box = boundsOf(control, kind);
		const name = vb6ControlName(control);
		const font = fontCssOf(control);
		const parts = [
			`left:${twipsToPt(box.left)}pt;top:${twipsToPt(box.top)}pt;width:${twipsToPt(box.width)}pt;height:${twipsToPt(box.height)}pt;`,
			font,
		];
		const back = colorCssOf(control, 'BackColor');
		const fore = colorCssOf(control, 'ForeColor');
		if (kind === 'Label' && (frmNumberOf(frmProperty(control, 'BackStyle')?.value) ?? 1) === 0) {
			parts.push('background:transparent;');
		} else if (back) {
			parts.push(`background:${back};`);
		}
		if (fore) { parts.push(`color:${fore};`); }
		if (boolOf(control, 'Enabled') === false) { parts.push('color:#6d6d6d;'); }
		const alignment = frmNumberOf(frmProperty(control, 'Alignment')?.value);
		if ((kind === 'Label' || kind === 'TextBox') && alignment !== undefined && ALIGNMENT_CSS[alignment]) {
			parts.push(ALIGNMENT_CSS[alignment]);
		}
		if (kind === 'TextBox' && boolOf(control, 'MultiLine')) { parts.push('white-space:pre-wrap;'); }
		if (kind === 'Label' && (frmNumberOf(frmProperty(control, 'BorderStyle')?.value) ?? 0) === 1) {
			parts.push('border:1px solid #7a7a7a;');
		}
		if (kind === 'Line' || kind === 'Shape') {
			const border = colorCssOf(control, 'BorderColor') ?? '#000000';
			if (kind === 'Line') {
				parts.push(`color:${border};`);
			} else {
				const shape = frmNumberOf(frmProperty(control, 'Shape')?.value) ?? 0;
				const radius = shape === 2 || shape === 3 ? '50%' : shape === 4 || shape === 5 ? '12%' : '0';
				const fillStyle = frmNumberOf(frmProperty(control, 'FillStyle')?.value) ?? 1;
				const fill = fillStyle === 0 ? (colorCssOf(control, 'FillColor') ?? '#000000') : 'transparent';
				parts.push(`border:1px solid ${border};border-radius:${radius};background:${fill};`);
			}
		}
		const picture = kind === 'Image' || kind === 'PictureBox' ? pictureCssOf(control, 'Picture', frx) : '';
		parts.push(picture);
		const scene = sceneControl({ kind: kind ?? 'Foreign', name, index, style: parts.join('') });
		scene.caption = spec?.captioned ? (textOf(control, 'Caption', frx) ?? '') : '';
		if (kind === 'TextBox' || kind === 'ComboBox') { scene.value = textOf(control, 'Text', frx) ?? ''; }
		if (kind === 'CheckBox') { scene.on = (frmNumberOf(frmProperty(control, 'Value')?.value) ?? 0) === 1; }
		if (kind === 'OptionButton') { scene.on = boolOf(control, 'Value') === true; }
		scene.pictured = picture !== '';
		if (kind === 'Line') {
			const { x1, y1, x2, y2 } = frmLineEnds(control);
			scene.lineDown = (x1 <= x2) === (y1 <= y2);
		}
		// Frames and PictureBoxes hold controls; so does any OCX or UserControl
		// the header nests children under (an SSTab page, a container
		// UserControl): those children draw on the foreign control's surface.
		if (spec?.container || control.children.some((c) => c.progId !== 'VB.Menu')) {
			scene.containerFontCss = font || DEFAULT_FONT_CSS;
			scene.surfaceStyle = back ? `background:${back};` : '';
			scene.children = sceneControlsOfChildren(control.children, frx);
		}
		out.push(scene);
	});
	return out;
}

/** The form's menu bar: the captions of its top-level menus, ampersands dropped. */
export function vb6MenuCaptions(form: FrmControl): string[] {
	return form.children
		.filter((c) => c.progId === 'VB.Menu')
		.map((c) => (unquoteVb6(frmProperty(c, 'Caption')?.value ?? c.name)).replace(/&(?!&)/g, '').replace(/&&/g, '&'))
		.filter((caption) => caption !== '-');
}

/** Reads a VB6 form header into the scene the canvas draws. */
export function sceneOfFrmHeader(header: FrmHeader, options: FrmSceneOptions): FormScene {
	const form = header.form;
	const frx = options.frx;
	const caption = textOf(form, 'Caption', frx) ?? options.formName;
	const back = colorCssOf(form, 'BackColor') ?? cssColor(0x8000000f);
	const fore = colorCssOf(form, 'ForeColor') ?? '#000';
	const borderStyle = frmNumberOf(frmProperty(form, 'BorderStyle')?.value) ?? 2;
	const border = borderStyle === 0 ? 'border:none;' : borderStyle === 1 || borderStyle === 3 ? 'border:1px solid #7a7a7a;' : '';
	return {
		form: {
			name: options.formName,
			caption,
			widthPt: twipsToPt(frmNumberOf(frmProperty(form, 'ClientWidth')?.value) ?? 4800),
			heightPt: twipsToPt(frmNumberOf(frmProperty(form, 'ClientHeight')?.value) ?? 3600),
			backCss: back,
			foreCss: fore,
			borderCss: border,
			pictureCss: pictureCssOf(form, 'Picture', frx),
			fontCss: fontCssOf(form) || DEFAULT_FONT_CSS,
			scrollBars: 0,
			zoom: 100,
			menus: vb6MenuCaptions(form),
		},
		controls: sceneControlsOfChildren(form.children, frx),
		pictures: {},
		toolbox: [...VB6_TOOLBOX],
		defaultEvents: VB6_DEFAULT_EVENTS,
		enums: vb6PaneVocabulary().enums,
		bools: vb6PaneVocabulary().bools,
		// MSForms' own enum tables must not answer a VB6 row: the value sets differ.
		paneBareEnums: false,
	};
}

export interface FrmPropertyRow {
	prop: string;
	value: string;
}

/** The geometry keys the pane shows in points, in rows of their own. */
export const GEOMETRY_KEYS: ReadonlySet<string> = new Set(['left', 'top', 'width', 'height']);
/** The form's size lives in its client size; the pane shows it as Width and Height. */
const FORM_SIZE_KEYS = new Set(['clientwidth', 'clientheight']);

/**
 * A property value as the pane shows it: booleans by name, an enum by its
 * number, a string unquoted, a color in the pane's own spelling (`#rrggbb`
 * or a system color name) so its swatch and picker work as they do for an
 * MSForms form; a gesture writes it back as `&H00BBGGRR&`.
 */
export function frmDisplayValue(p: FrmProperty): string {
	if (p.comment === 'True' && p.value.trim() === '-1') { return 'True'; }
	if (/^0\s+'False$/.test(p.value.trim())) { return 'False'; }
	if (/color$/i.test(p.key) && /^&H/i.test(p.value.trim())) {
		const color = parseOleColor(p.value);
		if (color !== undefined) { return formatOleColor(color); }
	}
	return unquoteVb6(p.value);
}

/** The property rows the pane shows for one control: its header, as written, geometry in points. */
function controlRows(control: FrmControl, frx: FrxLookup | undefined, kind: string | undefined): FrmPropertyRow[] {
	const rows: FrmPropertyRow[] = [{ prop: 'Name', value: control.name }];
	const index = frmProperty(control, 'Index');
	if (index && !index.frx) { rows.push({ prop: 'Index', value: index.value.trim() }); }
	if (kind === 'Form') {
		rows.push(
			{ prop: 'Width', value: twipsToPt(frmNumberOf(frmProperty(control, 'ClientWidth')?.value) ?? 0) },
			{ prop: 'Height', value: twipsToPt(frmNumberOf(frmProperty(control, 'ClientHeight')?.value) ?? 0) },
		);
	} else if (kind !== 'Line' && kind !== 'Timer') {
		const box = boundsOf(control, kind);
		rows.push(
			{ prop: 'Left', value: twipsToPt(box.left) }, { prop: 'Top', value: twipsToPt(box.top) },
			{ prop: 'Width', value: twipsToPt(box.width) }, { prop: 'Height', value: twipsToPt(box.height) },
		);
	} else if (kind === 'Timer') {
		const box = boundsOf(control, kind);
		rows.push({ prop: 'Left', value: twipsToPt(box.left) }, { prop: 'Top', value: twipsToPt(box.top) });
	}
	for (const m of control.members) {
		if (m.kind === 'group') {
			if (m.name.toLowerCase() === 'font') {
				for (const f of m.members) {
					if (f.kind === 'property') { rows.push({ prop: `Font.${f.key}`, value: frmDisplayValue(f) }); }
				}
			}
			continue;
		}
		const lower = m.key.toLowerCase();
		if (lower === 'index') { continue; }
		if (kind === 'Form' ? FORM_SIZE_KEYS.has(lower) : (kind !== 'Line' && GEOMETRY_KEYS.has(lower))) { continue; }
		if (m.frx) {
			const value = frx?.(m.frx);
			rows.push({ prop: m.key, value: value && (value.kind === 'longString' || value.kind === 'shortString') ? value.text : `(${m.frx.file}:${m.frx.offset.toString(16).toUpperCase().padStart(4, '0')})` });
			continue;
		}
		rows.push({ prop: m.key, value: frmDisplayValue(m) });
	}
	// The kind's design-time vocabulary, blank where the header says nothing.
	// A value the designer keeps in the `.frx` is listed only when the header
	// states it: the gesture that would set a blank one is refused, and a row
	// that cannot be filled is a promise the pane does not keep.
	const present = new Set(rows.map((r) => r.prop.toLowerCase()));
	for (const key of VB6_DESIGN_PROPERTIES[control.progId] ?? []) {
		if (key === 'Font') {
			if (!present.has('font.name')) {
				for (const field of VB6_FONT_FIELDS) { rows.push({ prop: `Font.${field}`, value: '' }); }
			}
			continue;
		}
		const lower = key.toLowerCase();
		if (present.has(lower) || VB6_SIDECAR_PICTURE_KEYS.has(lower) || VB6_SIDECAR_LIST_KEYS.has(lower)) { continue; }
		rows.push({ prop: key, value: '' });
	}
	return rows;
}

/** Property rows per target ('' is the form), keyed the way the canvas names controls. */
export function listFrmProperties(header: FrmHeader, options: FrmSceneOptions): Record<string, { kind: string; rows: FrmPropertyRow[] }> {
	const out: Record<string, { kind: string; rows: FrmPropertyRow[] }> = {};
	const form = header.form;
	const formRows = controlRows(form, options.frx, 'Form');
	if (!formRows.some((r) => r.prop === 'Caption')) {
		// A header without a Caption line shows the name, as the canvas does.
		formRows.splice(1, 0, { prop: 'Caption', value: options.formName });
	}
	out[''] = { kind: form.progId.startsWith('VB.') ? form.progId.slice(3) : form.progId, rows: formRows };
	const walk = (children: readonly FrmControl[]): void => {
		for (const control of children) {
			if (control.progId === 'VB.Menu') { continue; }
			const kind = vb6CanvasKind(control.progId);
			out[vb6ControlName(control)] = { kind: kind ?? control.progId, rows: controlRows(control, options.frx, kind) };
			walk(control.children);
		}
	};
	walk(form.children);
	return out;
}
