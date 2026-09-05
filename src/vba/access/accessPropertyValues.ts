import { PROPERTY_SLOTS } from './accessDesignTable';

/**
 * What the property pane offers for an Access design property: the value sets
 * of its enumerated properties, which of them are Yes/No, and which want a
 * colour swatch, a font list or a spinner.
 *
 * An Access property name means the same thing on every object that carries
 * it - the record code is global - so these are keyed by bare name rather
 * than by type.
 *
 * The value sets are the ones the Access VBA reference publishes for each
 * property; a property whose settings are not published stays a plain field
 * rather than a dropdown that could not offer every value. `OldBorderStyle`
 * carries the object model's `BorderStyle`, measured by setting
 * `TextBox.BorderStyle = 3` through Access and reading the design back: the
 * 3 landed on code 329, and code 11 (`BorderLineStyle`, which Access writes
 * on its own) is left unlabelled because its settings are not published.
 */
const ENUM_VALUES: Readonly<Record<string, ReadonlyArray<readonly [number, string]>>> = {
	BackStyle: [[0, 'Transparent'], [1, 'Normal']],
	OldBorderStyle: [
		[0, 'Transparent'], [1, 'Solid'], [2, 'Dashes'], [3, 'Short dashes'],
		[4, 'Dots'], [5, 'Sparse dots'], [6, 'Dash dot'], [7, 'Dash dot dot'],
	],
	BorderWidth: [
		[0, 'Hairline'], [1, '1 pt'], [2, '2 pt'], [3, '3 pt'],
		[4, '4 pt'], [5, '5 pt'], [6, '6 pt'],
	],
	SpecialEffect: [
		[0, 'Flat'], [1, 'Raised'], [2, 'Sunken'],
		[3, 'Etched'], [4, 'Shadowed'], [5, 'Chiseled'],
	],
	DisplayWhen: [[0, 'Always'], [1, 'Print Only'], [2, 'Screen Only']],
	TextAlign: [
		[0, 'General'], [1, 'Left'], [2, 'Center'], [3, 'Right'], [4, 'Distribute'],
	],
	TextFormat: [[0, 'Plain Text'], [1, 'Rich Text']],
	ScrollBars: [[0, 'None'], [2, 'Vertical']],
	ScrollBarAlign: [[0, 'System'], [1, 'Right'], [2, 'Left']],
	MultiSelect: [[0, 'None'], [1, 'Simple'], [2, 'Extended']],
	FilterLookup: [[0, 'Never'], [1, 'Database Default'], [2, 'Always']],
	PictureAlignment: [
		[0, 'Top Left'], [1, 'Top Right'], [2, 'Center'],
		[3, 'Bottom Left'], [4, 'Bottom Right'], [5, 'Form Center'],
	],
	PictureType: [[0, 'Embedded'], [1, 'Linked'], [2, 'Shared']],
	PictureCaptionArrangement: [
		[0, 'No Picture Caption'], [1, 'General'], [2, 'Top'],
		[3, 'Bottom'], [4, 'Left'], [5, 'Right'],
	],
	ReadingOrder: [[0, 'Context'], [1, 'Left-to-Right'], [2, 'Right-to-Left']],
	NumeralShapes: [[0, 'System'], [1, 'Arabic'], [2, 'National'], [3, 'Context']],
	IMEMode: [
		[0, 'No Control'], [1, 'On'], [2, 'Off'], [3, 'Disable'], [4, 'Hiragana'],
		[5, 'Full pitch Katakana'], [6, 'Half pitch Katakana'],
		[7, 'Full pitch Alpha/Num'], [8, 'Half pitch Alpha/Num'],
		[9, 'HangulFull'], [10, 'Hangul'],
	],
	HorizontalAnchor: [[0, 'Left'], [1, 'Right'], [2, 'Both']],
	VerticalAnchor: [[0, 'Top'], [1, 'Bottom'], [2, 'Both']],
	FontWeight: [
		[100, 'Thin'], [200, 'Extra Light'], [300, 'Light'], [400, 'Normal'],
		[500, 'Medium'], [600, 'Semi-bold'], [700, 'Bold'], [800, 'Extra Bold'],
		[900, 'Heavy'],
	],
	DecimalPlaces: [
		[255, 'Auto'], [0, '0'], [1, '1'], [2, '2'], [3, '3'], [4, '4'], [5, '5'],
		[6, '6'], [7, '7'], [8, '8'], [9, '9'], [10, '10'], [11, '11'], [12, '12'],
		[13, '13'], [14, '14'], [15, '15'],
	],
};

/** The value types a slot's third field carries, where the pane cares. */
const BOOLEAN_TYPE = 1;
const LONG_TYPE = 3;
const COLOR_TYPE = 4;
/** The one text slot Access sizes for a face name. */
const FONT_NAME = 'FontName';

/** Whether a slot holds an OLE colour rather than a long that reads like one. */
function isColorSlot(name: string, slot: readonly number[]): boolean {
	return slot[2] === COLOR_TYPE && slot[4] === 4 && name.endsWith('Color');
}

export type AccessPaneEditor = 'color' | 'font' | 'number';

export interface AccessPaneVocabulary {
	/** Dropdown options by property name, lowest value first. */
	enums: Record<string, [string, string][]>;
	/** The properties Access stores as Yes/No. */
	bools: string[];
	/** The properties that want an editor other than a text field. */
	editors: Record<string, AccessPaneEditor>;
}

let vocabulary: AccessPaneVocabulary | undefined;

/**
 * What the pane offers for every property any Access object type declares:
 * a Yes/No dropdown for a boolean slot, a swatch and picker for a colour
 * slot, the font list for `FontName`, a spinner for a plain long, and the
 * published value set where there is one.
 */
export function accessPaneVocabulary(): AccessPaneVocabulary {
	if (vocabulary) {
		return vocabulary;
	}
	const enums: Record<string, [string, string][]> = {};
	const bools = new Set<string>();
	const editors: Record<string, AccessPaneEditor> = {};
	for (const slots of PROPERTY_SLOTS.values()) {
		for (const [name, slot] of slots) {
			if (slot[2] === BOOLEAN_TYPE) {
				bools.add(name);
				continue;
			}
			if (isColorSlot(name, slot)) {
				editors[name] = 'color';
				continue;
			}
			if (name === FONT_NAME) {
				editors[name] = 'font';
				continue;
			}
			const values = ENUM_VALUES[name];
			if (values) {
				enums[name] = values.map(([value, word]) => [String(value), word]);
				continue;
			}
			if (slot[2] === LONG_TYPE) {
				editors[name] = 'number';
			}
		}
	}
	vocabulary = { enums, bools: [...bools].sort(), editors };
	return vocabulary;
}

/** Whether a property holds an OLE colour, which the markup spells rather than counts. */
export function isAccessColorProperty(name: string, slot: readonly number[]): boolean {
	return isColorSlot(name, slot);
}

/** Whether a property is Access's Yes/No, which the markup writes as True/False. */
export function isAccessBooleanProperty(slot: readonly number[]): boolean {
	return slot[2] === BOOLEAN_TYPE;
}
