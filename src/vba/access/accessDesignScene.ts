import {
	accessControlTypeName,
	accessDesignObjectName,
	isAccessDesignSection,
	type AccessDesign,
	type AccessDesignObject,
} from './accessDesign';
import { CONTROL_TYPES, PROPERTY_CODES, PROPERTY_SLOTS } from './accessDesignTable';
import {
	accessDesignEffectiveObjects,
	designObjectHolders,
	type AccessDesignKind,
} from './accessDesignEdit';
import { accessPaneVocabulary } from './accessPropertyValues';
import { cssColor, sceneControl, type FormScene, type SceneControl } from '../oforms/preview';

/**
 * An Access design as a designer scene, so the canvas that draws a UserForm
 * draws a form or report too.
 *
 * Access keeps geometry in twips (1/1440 inch) and the canvas in points
 * (1/72 inch), so every measurement is divided by twenty. A colour is an
 * OLE_COLOR, the same encoding MSForms uses, so the canvas converts it the
 * same way - system faces included.
 *
 * The scene carries only what can be drawn. A design's other properties reach
 * the property pane through the markup, which names every property its own
 * type's schema names.
 */

/** Twips to points: Access measures in 1/1440 inch, the canvas in 1/72. */
const TWIPS_PER_POINT = 20;

/** The kinds the toolbox offers; each is a type whose slots were measured. */
const TOOLBOX = [
	'Label', 'TextBox', 'CommandButton', 'CheckBox', 'OptionButton', 'ComboBox',
	'ListBox', 'Rectangle', 'Line', 'Image', 'ToggleButton', 'Subform',
];

/** How an Access control type draws on the canvas, where the names differ. */
const CANVAS_KINDS: ReadonlyMap<string, string> = new Map([
	// A Rectangle is a filled box with no caption, which is what a Label with
	// nothing to say draws; a Frame would put its own chrome around it and
	// take its fill from the scene's surface colour instead.
	['Rectangle', 'Label'], ['Subform', 'Foreign'], ['ObjectFrame', 'Foreign'],
	['BoundObjectFrame', 'Foreign'], ['CustomControl', 'Foreign'],
	['PageBreak', 'Foreign'], ['Attachment', 'Foreign'], ['WebBrowser', 'Foreign'],
	['NavigationControl', 'Foreign'], ['NavigationButton', 'CommandButton'],
	['Chart', 'Foreign'], ['EdgeBrowser', 'Foreign'], ['EmptyCell', 'Foreign'],
	['Tab', 'MultiPage'], ['Page', 'Frame'], ['OptionGroup', 'Frame'],
]);

function points(twips: number): number {
	return Math.round((twips / TWIPS_PER_POINT) * 100) / 100;
}

/** A named property of an object, read through its own type's schema. */
function propertyOf(object: AccessDesignObject, name: string): Buffer | undefined {
	const code = PROPERTY_CODES.get(name);
	return code === undefined
		? undefined
		: object.records.find((record) => record.code === code)?.value;
}

function numberOf(object: AccessDesignObject, name: string): number | undefined {
	const raw = propertyOf(object, name);
	if (!raw || raw.length === 0 || raw.length > 6) {
		return undefined;
	}
	return raw.readUIntLE(0, raw.length);
}

function signedOf(object: AccessDesignObject, name: string): number | undefined {
	const raw = propertyOf(object, name);
	if (!raw || raw.length === 0 || raw.length > 6) {
		return undefined;
	}
	return raw.readIntLE(0, raw.length);
}

function textOf(object: AccessDesignObject, name: string): string | undefined {
	const raw = propertyOf(object, name);
	return raw ? raw.toString('utf16le').replace(/\0+$/, '') : undefined;
}

/** The font a control draws its text in. */
function fontCss(object: AccessDesignObject): string {
	const face = textOf(object, 'FontName');
	const size = numberOf(object, 'FontSize');
	const weight = numberOf(object, 'FontWeight');
	const italic = numberOf(object, 'FontItalic');
	const underline = numberOf(object, 'FontUnderline');
	const parts: string[] = [];
	parts.push(`font-family:'${(face ?? 'Segoe UI').replace(/'/g, '')}',Tahoma,sans-serif;`);
	parts.push(`font-size:${size ?? 9}pt;`);
	if (weight !== undefined && weight >= 700) { parts.push('font-weight:700;'); }
	if (italic) { parts.push('font-style:italic;'); }
	if (underline) { parts.push('text-decoration:underline;'); }
	return parts.join('');
}

/** The CSS box a control occupies, with its colours and border. */
function styleOf(object: AccessDesignObject, kind: string, offsetPt = 0): string {
	const left = points(numberOf(object, 'Left') ?? 0);
	const top = points(numberOf(object, 'Top') ?? 0) + offsetPt;
	const width = points(numberOf(object, 'Width') ?? 0);
	const height = points(numberOf(object, 'Height') ?? 0);
	const parts = [
		`position:absolute;left:${left}pt;top:${top}pt;`,
		`width:${width}pt;height:${height}pt;`,
		fontCss(object),
	];
	const back = numberOf(object, 'BackColor');
	const backStyle = numberOf(object, 'BackStyle');
	// BackStyle 0 is transparent, which is what a label over a banner needs.
	if (back !== undefined && backStyle !== 0) {
		parts.push(`background:${cssColor(back)};`);
	} else if (kind === 'TextBox' || kind === 'ComboBox' || kind === 'ListBox') {
		parts.push('background:#fff;');
	}
	const fore = numberOf(object, 'ForeColor');
	if (fore !== undefined) {
		parts.push(`color:${cssColor(fore)};`);
	}
	const border = numberOf(object, 'BorderColor');
	const borderStyle = numberOf(object, 'BorderLineStyle') ?? numberOf(object, 'OldBorderStyle');
	if (border !== undefined && borderStyle !== 0) {
		parts.push(`border:1px solid ${cssColor(border)};`);
	}
	// Access's TextAlign: 1 left, 2 centre, 3 right. The canvas lays a label
	// out as a flex row, where text-align moves nothing, so the alignment is
	// given both ways and each kind takes the one that applies to it.
	const align = numberOf(object, 'TextAlign');
	if (align === 2) { parts.push('text-align:center;justify-content:center;'); }
	if (align === 3) { parts.push('text-align:right;justify-content:flex-end;'); }
	if (kind === 'Line') {
		// A line is drawn along the box it is given, so its border is the line
		// and the box itself has no fill.
		const rule = border !== undefined ? cssColor(border) : '#000';
		parts.push(`border:0;border-top:1px solid ${rule};background:none;`);
	}
	return parts.join('');
}

/** What a control shows on the canvas. */
function captionOf(object: AccessDesignObject, kind: string): string {
	if (kind === 'TextBox' || kind === 'ComboBox' || kind === 'ListBox') {
		return textOf(object, 'DefaultValue') ?? '';
	}
	return textOf(object, 'Caption') ?? textOf(object, 'ControlSource') ?? '';
}

function canvasKind(object: AccessDesignObject): string {
	const name = object.type === undefined ? undefined : CONTROL_TYPES.get(object.type);
	if (!name) {
		return 'Foreign';
	}
	return CANVAS_KINDS.get(name) ?? name;
}

/**
 * The design as a scene: every section stacked down the surface, with the
 * controls it holds placed inside it. Access draws a form as its sections in
 * order, so the canvas does the same and offsets each band by the height of
 * the ones above it.
 */
export function sceneOfAccessDesign(
	design: AccessDesign,
	name: string,
	kind: AccessDesignKind,
): FormScene {
	const vocabulary = accessPaneVocabulary();
	// A control that agrees with its type's defaults carries no record of its
	// own for them, so the canvas draws what Access draws only once they are
	// folded in.
	const objects = accessDesignEffectiveObjects(design);
	const holders = designObjectHolders(objects);
	const childrenOf = new Map<number, number[]>();
	holders.forEach((holder, index) => {
		const list = childrenOf.get(holder) ?? [];
		list.push(index);
		childrenOf.set(holder, list);
	});

	let index = 0;
	const build = (at: number, offsetPt: number): SceneControl => {
		const object = objects[at];
		const kindName = canvasKind(object);
		const kids = (childrenOf.get(at) ?? []).map((child) => build(child, 0));
		const control = sceneControl({
			kind: kindName,
			name: accessDesignObjectName(object) ?? `Object${at}`,
			index: index++,
			style: styleOf(object, kindName, offsetPt),
			caption: captionOf(object, kindName),
			children: kids,
		});
		if (kindName === 'CheckBox' || kindName === 'OptionButton' || kindName === 'ToggleButton') {
			control.on = (numberOf(object, 'DefaultValue') ?? 0) !== 0;
		}
		return control;
	};

	// Each section is a band down the surface, and its controls sit inside it:
	// Access stores a control's Top within its own section, so the canvas nests
	// them the same way and the two agree without an offset.
	const controls: SceneControl[] = [];
	// White is what the blank form and report Access writes are; a design that
	// names its own colour overrides it below.
	const formBack = numberOf(objects[0], 'BackColor') ?? 0xffffff;
	let bandTop = 0;
	let widest = 0;
	for (let at = 0; at < objects.length; at += 1) {
		const object = objects[at];
		if (!isAccessDesignSection(object)) {
			continue;
		}
		const height = points(numberOf(object, 'Height') ?? 0);
		const children = (childrenOf.get(at) ?? []).map((child) => {
			const span = points(
				(numberOf(objects[child], 'Left') ?? 0) + (numberOf(objects[child], 'Width') ?? 0),
			);
			widest = Math.max(widest, span);
			return build(child, 0);
		});
		// A section Access gave no colour of its own draws on the form's, which
		// is white on the blank template: a transparent label has to sit on
		// something, or the canvas shows through its text.
		const back = numberOf(object, 'BackColor') ?? formBack;
		controls.push(sceneControl({
			kind: 'Frame',
			name: accessDesignObjectName(object) ?? `Section${at}`,
			index: index++,
			style: `position:absolute;left:0pt;top:${bandTop}pt;right:0pt;height:${height}pt;`
				+ `${fontCss(object)}border:0;background:${cssColor(back)};`,
			surfaceStyle: `background:${cssColor(back)};`,
			children,
		}));
		bandTop += height;
	}

	const own = objects[0];
	const width = points(numberOf(own, 'Width') ?? 0) || widest || 400;
	const back = numberOf(own, 'BackColor');
	return {
		form: {
			name,
			caption: textOf(own, 'Caption') ?? name,
			widthPt: String(width),
			heightPt: String(bandTop || 200),
			backCss: back !== undefined ? cssColor(back) : '#ffffff',
			foreCss: signedOf(own, 'ForeColor') !== undefined
				? cssColor(numberOf(own, 'ForeColor')!)
				: '#000',
			borderCss: '',
			pictureCss: '',
			fontCss: fontCss(own),
			zoom: 100,
			scrollBars: 0,
			menus: [],
		},
		controls,
		pictures: {},
		toolbox: TOOLBOX,
		// A report has no events of its own on the canvas, and a form's default
		// is the click Access writes for a button.
		defaultEvents: kind === 'form' ? { CommandButton: 'Click' } : {},
		// The pane's rows come from the same schema the markup prints, so the
		// vocabulary is the schema's: a colour slot gets the picker, `FontName`
		// the font list, a Yes/No slot True/False, and a property whose
		// settings Access publishes gets those.
		enums: vocabulary.enums,
		bools: vocabulary.bools,
		paneEditors: vocabulary.editors,
		// Access spells the same property names as MSForms with different
		// values behind them, so the pane's built-in tables stay out of it.
		paneBareEnums: false,
	};
}

/** Whether the designer knows how to draw a control of this type. */
export function accessCanvasKind(typeName: string): string | undefined {
	return PROPERTY_SLOTS.has(typeName)
		? (CANVAS_KINDS.get(typeName) ?? typeName)
		: undefined;
}

export { accessControlTypeName };
