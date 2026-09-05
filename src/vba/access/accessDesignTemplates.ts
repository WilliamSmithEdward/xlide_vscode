import { readExtensionBinaryAsset } from '../../extensionAssets';
import { AccessFormatError } from './accessFormat';
import {
	buildAccessDesign,
	isAccessDesignSection,
	parseAccessDesign,
	type AccessDesignRecord,
} from './accessDesign';
import type { AccessDesignKind, AccessDesignPrototypes } from './accessDesignEdit';

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
 * captured ones for every type it does not.
 */
export function availablePrototypes(
	kind: AccessDesignKind,
	carried: AccessDesignPrototypes,
): AccessDesignPrototypes {
	const out = new Map(accessDesignTemplate(kind).prototypes);
	for (const [type, records] of carried) {
		out.set(type, records);
	}
	return out;
}

/** The template's blob and catalog properties with a GUID of their own. */
export function withDesignGuid(
	kind: AccessDesignKind,
	guid: Buffer,
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
	return { blob: buildAccessDesign({ ...design, objects }), catalogProperties };
}
