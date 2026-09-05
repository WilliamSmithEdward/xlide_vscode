import { AccessFormatError } from './accessFormat';

/**
 * An Access form or report design.
 *
 * The design lives in `MSysAccessStorage` under `Forms` or `Reports`, laid out
 * the way a module's does under `Modules`: a numbered folder per object with
 * the design in a `Blob` beneath it, a `\x03DirData` beside the folders naming
 * them, and `TypeInfo`, `PropData` and an empty `BlobDelta` alongside.
 *
 * The blob is a header, a stream of property records, and a trailer. A record
 * is `<u32 id><u16 code><u32 value type><u32 width><u32 length>` and then that
 * many bytes, with the ids ascending inside one object. Three ids are not
 * properties but markers opening the next object, each followed by a `u16`:
 * `0xFE` opens a section, `0xFD` the next object at the same level, and `0xFF`
 * a control, carrying a second `u16` that names its type. Ids restart at a
 * marker, and that is what separates one object's records from the next.
 *
 * What a property MEANS is a separate question this does not answer: a
 * record's `code` is the property, a few are named below, and the rest are
 * handed back as they are. Ported from pyOpenVBA's `_designs.py`, whose own
 * measurements rebuild every design it read byte for byte
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/67).
 */

/** `<u32 id><u16 code><u32 type><u32 width><u32 length>`. */
const HEAD = 18;
/** The bytes before the first record, and the ones that close the stream. */
const DESIGN_HEADER = 10;
const DESIGN_TRAILER = 4;

/** Ids that open an object instead of carrying a property. */
const OPEN_SECTION = 0xfe;
const OPEN_SIBLING = 0xfd;
const OPEN_CONTROL = 0xff;
const MARKERS = new Set([OPEN_SECTION, OPEN_SIBLING, OPEN_CONTROL]);

/** `MSysObjects.Type` for each kind. */
export const ACCESS_FORM_TYPE = -32768;
export const ACCESS_REPORT_TYPE = -32764;

/** Control and section types, by the `u16` an `0xFF` marker carries. */
export const ACCESS_CONTROL_TYPES: Readonly<Record<number, string>> = {
	100: 'Label', 101: 'Rectangle', 102: 'Line', 103: 'Image',
	104: 'CommandButton', 105: 'OptionButton', 106: 'CheckBox', 107: 'OptionGroup',
	108: 'BoundObjectFrame', 109: 'TextBox', 110: 'ListBox', 111: 'ComboBox',
	112: 'Subform', 114: 'ObjectFrame', 118: 'PageBreak', 119: 'CustomControl',
	122: 'ToggleButton', 123: 'Tab', 124: 'Page', 126: 'Attachment',
	127: 'EmptyCell', 128: 'WebBrowser', 129: 'NavigationControl',
	130: 'NavigationButton', 133: 'Chart', 134: 'EdgeBrowser',
	152: 'Detail', 153: 'HeaderSection', 154: 'FooterSection',
	155: 'PageHeaderSection', 156: 'PageFooterSection',
	157: 'GroupHeaderSection', 158: 'GroupFooterSection',
};

/** 152 through 158 are the sections; everything else is a control. */
function isSectionType(type: number | undefined): boolean {
	return type !== undefined && type >= 152 && type <= 158;
}

/** The codes whose value is the object's own name. */
const NAME_CODES = new Set([20, 21]);

export interface AccessDesignRecord {
	id: number;
	/** The property. A few are named; most are the file's own numbers. */
	code: number;
	valueType: number;
	width: number;
	value: Buffer;
}

export interface AccessDesignObject {
	/** The marker that opened it, absent for the first object. */
	marker?: number;
	/** The control or section type the marker carried. */
	type?: number;
	/** The `u16` the marker carried before the type. */
	code?: number;
	records: AccessDesignRecord[];
}

export interface AccessDesign {
	/** The bytes before the first record, kept so a design can be rebuilt. */
	header: Buffer;
	objects: AccessDesignObject[];
	trailer: Buffer;
}

/** A design blob, as its objects. */
export function parseAccessDesign(blob: Buffer): AccessDesign {
	if (blob.length < DESIGN_HEADER + DESIGN_TRAILER) {
		throw new AccessFormatError('A design blob is too short to hold its header.');
	}
	const objects: AccessDesignObject[] = [];
	let records: AccessDesignRecord[] = [];
	let marker: number | undefined;
	let type: number | undefined;
	let code: number | undefined;
	const close = (): void => {
		objects.push({
			...(marker !== undefined ? { marker } : {}),
			...(type !== undefined ? { type } : {}),
			...(code !== undefined ? { code } : {}),
			records,
		});
	};

	let at = DESIGN_HEADER;
	let last = -1;
	while (at + 6 <= blob.length) {
		const id = blob.readUInt32LE(at);
		if (MARKERS.has(id)) {
			close();
			records = [];
			marker = id;
			code = blob.readUInt16LE(at + 4);
			at += 6;
			type = code;
			if (id === OPEN_CONTROL) {
				type = blob.readUInt16LE(at);
				at += 2;
			}
			last = -1;
			continue;
		}
		if (at + HEAD > blob.length) {
			break;
		}
		const recordCode = blob.readUInt16LE(at + 4);
		const valueType = blob.readUInt32LE(at + 6);
		const width = blob.readUInt32LE(at + 10);
		const length = blob.readUInt32LE(at + 14);
		// Ids ascend inside one object, so one that does not is the trailer
		// rather than a record - which is how the stream's end is found.
		if (id <= last || length > blob.length - at - HEAD || valueType > 0xffff) {
			break;
		}
		records.push({
			id,
			code: recordCode,
			valueType,
			width,
			value: blob.subarray(at + HEAD, at + HEAD + length),
		});
		last = id;
		at += HEAD + length;
	}
	close();
	return { header: blob.subarray(0, DESIGN_HEADER), objects, trailer: blob.subarray(at) };
}

/** The blob for a parsed design, byte for byte when nothing was changed. */
export function buildAccessDesign(design: AccessDesign): Buffer {
	const parts: Buffer[] = [Buffer.from(design.header)];
	for (const object of design.objects) {
		if (object.marker !== undefined) {
			const head = Buffer.alloc(object.marker === OPEN_CONTROL ? 8 : 6);
			head.writeUInt32LE(object.marker, 0);
			head.writeUInt16LE(object.code ?? 0, 4);
			if (object.marker === OPEN_CONTROL) {
				head.writeUInt16LE(object.type ?? 0, 6);
			}
			parts.push(head);
		}
		for (const record of object.records) {
			const head = Buffer.alloc(HEAD);
			head.writeUInt32LE(record.id, 0);
			head.writeUInt16LE(record.code, 4);
			head.writeUInt32LE(record.valueType, 6);
			head.writeUInt32LE(record.width, 10);
			head.writeUInt32LE(record.value.length, 14);
			parts.push(head, Buffer.from(record.value));
		}
	}
	parts.push(Buffer.from(design.trailer));
	return Buffer.concat(parts);
}

/** The object's own name, from whichever record carries it. */
export function accessDesignObjectName(object: AccessDesignObject): string | undefined {
	for (const record of object.records) {
		if (NAME_CODES.has(record.code)) {
			const text = decodeDesignText(record.value);
			if (text) {
				return text;
			}
		}
	}
	return undefined;
}

/** Whether the object is one of the design's sections rather than a control. */
export function isAccessDesignSection(object: AccessDesignObject): boolean {
	return isSectionType(object.type);
}

/** The type's name, when it is one this knows. */
export function accessControlTypeName(object: AccessDesignObject): string | undefined {
	return object.type === undefined ? undefined : ACCESS_CONTROL_TYPES[object.type];
}

/**
 * The design's sections and its NAMED controls. A design also carries unnamed
 * prototypes - the styles a new control is cut from - and those are not part
 * of what the designer shows.
 */
export function accessDesignControls(design: AccessDesign): AccessDesignObject[] {
	return design.objects.slice(1).filter(
		(object) => !isAccessDesignSection(object) && accessDesignObjectName(object) !== undefined,
	);
}

export function accessDesignSections(design: AccessDesign): AccessDesignObject[] {
	return design.objects.slice(1).filter(isAccessDesignSection);
}

/**
 * A design's text value. Jet 4 stores text as UTF-16LE, and reading it a byte
 * at a time gives a name that PRINTS correctly and compares equal to nothing:
 * `Banner` arrives as `B\0a\0n\0n\0e\0r`, whose NULs are invisible in a log and
 * fatal in a lookup. An odd length cannot be UTF-16, so that falls back.
 */
function decodeDesignText(value: Buffer): string {
	const text = value.length % 2 === 0
		? value.toString('utf16le')
		: value.toString('latin1');
	return text.replace(/\0+$/, '');
}
