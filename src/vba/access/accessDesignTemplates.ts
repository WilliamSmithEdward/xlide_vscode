import { readExtensionBinaryAsset } from '../../extensionAssets';
import { AccessFormatError } from './accessFormat';
import {
	buildAccessDesign,
	isAccessDesignSection,
	parseAccessDesign,
	type AccessDesignRecord,
} from './accessDesign';
import type { AccessDesignKind, AccessDesignPrototypes } from './accessDesignEdit';
import { CONTROL_TYPES, FONT_FAMILIES, PROPERTY_CODES, PROPERTY_SLOTS } from './accessDesignTable';
import { accessThemeColor, OFFICE_2023_THEME, sameAccessTheme, type AccessTheme } from './accessThemes';

/**
 * What a new Access form or report is built from.
 *
 * A design created from nothing has to be one Access would have created, so
 * the starting blobs are captured from Access itself: `assets/access/` holds a
 * blank form and a blank report exactly as Access 16.0 wrote them, with their
 * `TypeInfo`, `PropData` and catalog property blobs, and a design holding one
 * control of each type - which is where a control's defaults object comes
 * from. Access reads a control's themed properties against that object; a
 * control written without one renders as a default themed control and Access
 * drops its colours on the next save.
 *
 * The design's own GUID is record 208, and the catalog's property blob repeats
 * it, so a new design gets a fresh one written into both. Two designs sharing
 * a GUID is not something Access writes.
 */

/** The record whose value is the design's own GUID. */
const GUID_RECORD = 208;
const GUID_LENGTH = 16;

export interface AccessDesignTemplate {
	blob: Buffer;
	typeInfo: Buffer;
	propData: Buffer;
	/** The catalog row's `LvProp`, which repeats the design's GUID. */
	catalogProperties: Buffer;
	prototypes: AccessDesignPrototypes;
	/** The GUID the captured template carries, which a new design replaces. */
	capturedGuid: Buffer;
}

const cache = new Map<string, AccessDesignTemplate>();

/** The captured template for a kind of design, read once. */
export function accessDesignTemplate(kind: AccessDesignKind): AccessDesignTemplate {
	const found = cache.get(kind);
	if (found) {
		return found;
	}
	const read = (suffix: string): Buffer => {
		const file = `assets/access/${kind}.${suffix}`;
		try {
			return readExtensionBinaryAsset(file);
		} catch {
			throw new AccessFormatError(
				`The captured ${kind} template is missing (${file}); a design cannot be created `
				+ 'without one, because a design Access did not write is one Access repairs.',
			);
		}
	};
	const blob = read('blob');
	const design = parseAccessDesign(blob);
	const guid = design.objects
		.flatMap((object) => object.records)
		.find((record) => record.id === GUID_RECORD && record.value.length === GUID_LENGTH)?.value;
	if (!guid) {
		throw new AccessFormatError(`The captured ${kind} template carries no GUID record.`);
	}
	const template: AccessDesignTemplate = {
		blob,
		typeInfo: read('typeinfo'),
		propData: read('propdata'),
		catalogProperties: read('lvprop'),
		prototypes: prototypesOf(read('prototypes')),
		capturedGuid: Buffer.from(guid),
	};
	cache.set(kind, template);
	return template;
}

/** The control-defaults objects a captured design carries, by control type. */
function prototypesOf(blob: Buffer): AccessDesignPrototypes {
	const out = new Map<number, readonly AccessDesignRecord[]>();
	for (const object of parseAccessDesign(blob).objects.slice(1)) {
		if (isAccessDesignSection(object)) {
			break;
		}
		if (object.type !== undefined) {
			out.set(object.type, object.records);
		}
	}
	return out;
}

/**
 * The prototypes a design can draw on: the ones it already carries, and the
 * captured ones, drawn in the database's `theme`, for every type it does not.
 */
export function availablePrototypes(
	kind: AccessDesignKind,
	carried: AccessDesignPrototypes,
	theme: AccessTheme = OFFICE_2023_THEME,
): AccessDesignPrototypes {
	let out: Map<number, readonly AccessDesignRecord[]>;
	try {
		const captured = accessDesignTemplate(kind).prototypes;
		out = new Map(onCaptureTheme(theme) ? captured : rethemedAccessPrototypes(captured, theme));
	} catch {
		// Without the captured templates a design can still be edited; only a
		// control of a type it does not already hold is out of reach, and
		// addDesignControl refuses that with the reason.
		out = new Map();
	}
	for (const [type, records] of carried) {
		out.set(type, records);
	}
	return out;
}

/**
 * The template's blob and catalog properties with a GUID of their own, the
 * blob drawn in the database's `theme`.
 */
export function withDesignGuid(
	kind: AccessDesignKind,
	guid: Buffer,
	theme: AccessTheme = OFFICE_2023_THEME,
): { blob: Buffer; catalogProperties: Buffer } {
	if (guid.length !== GUID_LENGTH) {
		throw new AccessFormatError(`A design GUID is ${GUID_LENGTH} bytes, not ${guid.length}.`);
	}
	const template = accessDesignTemplate(kind);
	const design = parseAccessDesign(template.blob);
	const objects = design.objects.map((object) => ({
		...object,
		records: object.records.map((entry) => entry.id === GUID_RECORD
			&& entry.value.length === GUID_LENGTH
			? { ...entry, value: Buffer.from(guid) }
			: entry),
	}));
	// The catalog's property blob repeats the GUID; replacing the captured
	// bytes wherever they appear leaves the rest of the blob exactly as Access
	// wrote it, which no partial parse of it could promise.
	const catalogProperties = Buffer.from(template.catalogProperties);
	const at = catalogProperties.indexOf(template.capturedGuid);
	if (at < 0) {
		throw new AccessFormatError(
			`The captured ${kind} catalog properties do not repeat the design's GUID.`,
		);
	}
	guid.copy(catalogProperties, at);
	const blob = buildAccessDesign({ ...design, objects });
	return { blob: onCaptureTheme(theme) ? blob : rethemedAccessDesign(blob, theme), catalogProperties };
}

/**
 * Whether `theme` is the one the templates were captured in. Recomputing them
 * on it gives them back byte for byte (tests/accessThemes.test.ts), so the
 * work is skipped.
 */
function onCaptureTheme(theme: AccessTheme): boolean {
	return sameAccessTheme(theme, OFFICE_2023_THEME);
}

// --- what a design takes from its database's theme --------------------------------
// The templates were captured in a database on Office's 2023 theme. Access
// builds the same objects from whatever theme the database holds (issue #151,
// from pyOpenVBA 6.3.0, which had Access make a form and a report with a
// control of every type on each Office theme):
//
// - a font comes from the theme by its ThemeFontIndex, 0 the heading face and
//   1 the body face, with the face's family byte beside it (TextFontFamily)
//   unless it is the default a sans face takes;
// - a colour is worked out again from the slot, tint and shade beside it;
// - the design itself carries the body face, its family byte and the theme's
//   Background 2 (lt2) colour.
//
// On the 2023 theme this gives the templates back byte for byte. Access also
// sizes a tab control's strip and pages by the face's metrics, which is not
// reproduced.

const THEME_FONT_INDEX = PROPERTY_CODES.get('ThemeFontIndex')!;
const FONT_NAME = PROPERTY_CODES.get('FontName')!;
const TEXT_FONT_FAMILY = PROPERTY_CODES.get('TextFontFamily')!;
/** A family byte is one byte of value type 2. */
const FAMILY_VALUE_TYPE = 2;
const FAMILY_WIDTH = 1;
/** The design's own face, its family byte's slot, and its Background 2. */
const DESIGN_FONT = 160;
const DESIGN_FONT_FAMILY = { id: 56, code: 244 };
const DESIGN_BACKGROUND = 319;
const BACKGROUND_2 = 3;
/** A defaults object's family-byte slot where the measured slots have none. */
const FAMILY_IDS: Readonly<Record<string, number>> = { NavigationButton: 368 };
/** Each themed colour beside the slot, tint and shade it is worked out from. */
const THEMED_COLORS = ['Back', 'Border', 'Fore', 'Gridline', 'Hover', 'Pressed', 'HoverFore', 'PressedFore']
	.map((part) => ({
		color: PROPERTY_CODES.get(`${part}Color`),
		slot: PROPERTY_CODES.get(`${part}ThemeColorIndex`),
		tint: PROPERTY_CODES.get(`${part}Tint`),
		shade: PROPERTY_CODES.get(`${part}Shade`),
	}))
	.filter((entry): entry is { color: number; slot: number; tint: number | undefined; shade: number | undefined } =>
		entry.color !== undefined && entry.slot !== undefined);

function signed(record: AccessDesignRecord): number {
	const value = record.value;
	return value.length >= 4 ? value.readInt32LE(0) : value.length === 2 ? value.readInt16LE(0) : value[0] ?? 0;
}

/** Set the family byte `face` takes: none for the default, else at `slot`. */
function themedFace(
	out: Map<number, AccessDesignRecord>,
	face: string,
	family: AccessDesignRecord | undefined,
	slot: { id: number; code: number } | undefined,
): void {
	const byte = FONT_FAMILIES.get(face);
	if (byte === undefined) {
		if (family) {
			out.delete(family.id);
		}
	} else if (family) {
		out.set(family.id, { ...family, value: Buffer.from([byte]) });
	} else if (slot) {
		out.set(slot.id, { id: slot.id, code: slot.code, valueType: FAMILY_VALUE_TYPE, width: FAMILY_WIDTH, value: Buffer.from([byte]) });
	}
}

/**
 * Records captured on the 2023 theme, as `theme` gives them. `familyId` is
 * where the object's family byte goes when the theme's face needs one and
 * the object has none.
 */
export function rethemedAccessRecords(
	records: readonly AccessDesignRecord[],
	theme: AccessTheme,
	familyId: number | undefined,
): AccessDesignRecord[] {
	const byCode = new Map(records.map((record) => [record.code, record]));
	const out = new Map(records.map((record) => [record.id, record]));
	const index = byCode.get(THEME_FONT_INDEX);
	const font = byCode.get(FONT_NAME);
	if (index && font && (signed(index) === 0 || signed(index) === 1)) {
		const face = signed(index) === 0 ? theme.majorFont : theme.minorFont;
		out.set(font.id, { ...font, value: Buffer.from(face, 'utf16le') });
		themedFace(out, face, byCode.get(TEXT_FONT_FAMILY),
			familyId === undefined ? undefined : { id: familyId, code: TEXT_FONT_FAMILY });
	}
	for (const { color, slot, tint, shade } of THEMED_COLORS) {
		const found = byCode.get(color);
		const slotRecord = byCode.get(slot);
		if (!found || !slotRecord || signed(slotRecord) < 0) {
			continue;
		}
		const tintRecord = tint === undefined ? undefined : byCode.get(tint);
		const shadeRecord = shade === undefined ? undefined : byCode.get(shade);
		out.set(found.id, {
			...found,
			value: accessThemeColor(
				theme,
				signed(slotRecord),
				tintRecord ? tintRecord.value.readFloatLE(0) : 100,
				shadeRecord ? shadeRecord.value.readFloatLE(0) : 100,
			),
		});
	}
	return [...out.keys()].sort((a, b) => a - b).map((id) => out.get(id)!);
}

/** A control-defaults object's records for `type`, as `theme` gives them. */
function rethemedPrototype(type: number, records: readonly AccessDesignRecord[], theme: AccessTheme): AccessDesignRecord[] {
	const typeName = CONTROL_TYPES.get(type) ?? '';
	const measured = PROPERTY_SLOTS.get(typeName)?.get('TextFontFamily');
	return rethemedAccessRecords(records, theme, measured ? measured[0] : FAMILY_IDS[typeName]);
}

/** An empty design from the templates, as `theme` gives it. */
export function rethemedAccessDesign(blob: Buffer, theme: AccessTheme): Buffer {
	const design = parseAccessDesign(blob);
	const [own, ...rest] = design.objects;
	const records = new Map(own.records.map((record) => [record.id, record]));
	const font = own.records.find((record) => record.code === DESIGN_FONT);
	if (font) {
		records.set(font.id, { ...font, value: Buffer.from(theme.minorFont, 'utf16le') });
	}
	const background = own.records.find((record) => record.code === DESIGN_BACKGROUND);
	if (background) {
		records.set(background.id, { ...background, value: accessThemeColor(theme, BACKGROUND_2) });
	}
	themedFace(records, theme.minorFont,
		own.records.find((record) => record.code === DESIGN_FONT_FAMILY.code), DESIGN_FONT_FAMILY);
	const objects = [
		{ ...own, records: [...records.keys()].sort((a, b) => a - b).map((id) => records.get(id)!) },
		...rest.map((object) => ({ ...object, records: rethemedAccessRecords(object.records, theme, undefined) })),
	];
	return buildAccessDesign({ ...design, objects });
}

/** The captured control-defaults objects, as `theme` gives them. */
export function rethemedAccessPrototypes(prototypes: AccessDesignPrototypes, theme: AccessTheme): AccessDesignPrototypes {
	return new Map([...prototypes].map(([type, records]) => [type, rethemedPrototype(type, records, theme)]));
}
