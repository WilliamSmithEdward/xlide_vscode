import { AccessFormatError } from './accessFormat';
import {
	accessDesignObjectName,
	buildAccessDesign,
	isAccessDesignSection,
	parseAccessDesign,
	type AccessDesign,
	type AccessDesignObject,
	type AccessDesignRecord,
} from './accessDesign';
import {
	BUTTON_EXTRAS,
	COMPANIONS,
	CONTROL_SLOTS,
	CONTROL_TYPES,
	DEFAULT_IME_SENTENCE_MODE,
	DEFAULT_OVERLAP,
	DEFAULT_ROW_SOURCE_TYPE,
	DESIGN_OBJECT,
	FONT_FAMILIES,
	NO_PICTURE,
	OVERLAP_FLAGS,
	PROPERTY_CODES,
	PROPERTY_SLOTS,
	READ_ONLY_TYPES,
	TABBABLE,
	TEXT_VALUE_TYPES,
	TYPE_CODES,
	TYPE_EXTRAS,
} from './accessDesignTable';

/**
 * Editing an Access form or report design
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/67).
 *
 * A design is a stream of property records grouped into objects by three
 * marker ids, and what an object may carry is decided by its own type's
 * schema: a record's id is its slot in that schema, so a property written at
 * the wrong id is a property Access reads as something else. The slot tables
 * were read off designs Access itself wrote; a type whose slots were not
 * measured is refused rather than guessed.
 *
 * The markers depend on how many controls a section holds, so adding or
 * removing one rewrites the markers of the controls already there. Access does
 * not complain about a wrong count: it opens the form and shows only as many
 * controls as the number claims.
 *
 * Ported from pyOpenVBA's `_designs.py`.
 */

const OPEN_SECTION = 0xfe;
const OPEN_SIBLING = 0xfd;
const OPEN_CONTROL = 0xff;

const GUID_LENGTH = 16;
const PAGE = 'Page';
const PAGE_HOLDER = 'Tab';
/** A page break has no control-defaults object. */
const DEFAULTS_FREE = new Set([118]);
/**
 * Beside the `RowSourceType` text, a list or combo box carries the kind as a
 * byte at this slot when it is not the default table or query.
 */
const ROW_SOURCE_KIND_CODE = 92;
const ROW_SOURCE_KIND_SLOT = [48, ROW_SOURCE_KIND_CODE, 2, 1] as const;
const ROW_SOURCE_KINDS: ReadonlyMap<string, number> = new Map([
	['Value List', 1], ['Field List', 10],
]);
/** A command button's fill also sets a gradient. */
const BUTTON_FILL_COMPANIONS: ReadonlyArray<readonly [string, number]> = [['Gradient', 0]];

export type AccessDesignKind = 'form' | 'report';

/** A property value as a caller supplies it. */
export type AccessDesignValue = string | number | boolean | Buffer;

/** The object's name, or undefined for the design itself and the prototypes. */
function nameOf(object: AccessDesignObject): string | undefined {
	return accessDesignObjectName(object);
}

/** The schema an object's properties come from. */
function schemaOf(object: AccessDesignObject): string | undefined {
	return object.type === undefined ? DESIGN_OBJECT : CONTROL_TYPES.get(object.type);
}

function recordValue(object: AccessDesignObject, code: number): Buffer | undefined {
	return object.records.find((record) => record.code === code)?.value;
}

function record(slot: readonly number[], value: Buffer): AccessDesignRecord {
	return { id: slot[0], code: slot[1], valueType: slot[2], width: slot[3], value };
}

/** One property's value as its slot says to write it. */
export function encodeDesignProperty(
	slot: readonly number[],
	value: AccessDesignValue,
): Buffer {
	const valueType = slot[2];
	const size = slot[4] ?? 0;
	if (TEXT_VALUE_TYPES.has(valueType)) {
		if (typeof value !== 'string') {
			throw new AccessFormatError('This property takes text.');
		}
		return Buffer.from(value, 'utf16le');
	}
	if (valueType === 11) {
		if (!Buffer.isBuffer(value)) {
			throw new AccessFormatError('This property takes raw bytes.');
		}
		return Buffer.from(value);
	}
	if (valueType === 9) {
		if (!Buffer.isBuffer(value) || value.length !== GUID_LENGTH) {
			throw new AccessFormatError(`A GUID is ${GUID_LENGTH} bytes.`);
		}
		return Buffer.from(value);
	}
	if (valueType === 6) {
		const out = Buffer.alloc(4);
		out.writeFloatLE(Number(value), 0);
		return out;
	}
	if (valueType === 8) {
		const out = Buffer.alloc(8);
		out.writeDoubleLE(Number(value), 0);
		return out;
	}
	const number = typeof value === 'boolean' ? (value ? 1 : 0) : value;
	if (typeof number !== 'number' || !Number.isInteger(number)) {
		throw new AccessFormatError('This property takes a whole number.');
	}
	const width = size || 1;
	const out = Buffer.alloc(width);
	try {
		if (number < 0) {
			out.writeIntLE(number, 0, width);
		} else {
			out.writeUIntLE(number, 0, width);
		}
	} catch {
		throw new AccessFormatError(`${number} does not fit the property's ${width} bytes.`);
	}
	return out;
}

/** The records Access writes beside a colour or a font. */
function companionsOf(
	schema: string | undefined,
	name: string,
	value: AccessDesignValue,
): Array<readonly [string, number]> {
	const out = [...(COMPANIONS.get(name) ?? [])];
	if (name === 'BackColor' && schema === 'CommandButton') {
		out.push(...BUTTON_FILL_COMPANIONS);
	}
	if (name === 'FontName' && typeof value === 'string' && FONT_FAMILIES.has(value)) {
		out.push(['TextFontFamily', FONT_FAMILIES.get(value)!]);
	}
	return out;
}

/**
 * A design with one property of one of its objects changed. `target` names a
 * control or a section; undefined means the design itself. A property the
 * object already carries keeps its record's id, so the records stay in the
 * order Access reads them.
 */
export function setDesignProperty(
	design: AccessDesign,
	target: string | undefined,
	name: string,
	value: AccessDesignValue,
): AccessDesign {
	const at = design.objects.findIndex((object) => target === undefined
		? object.type === undefined
		: nameOf(object) === target);
	if (at < 0) {
		throw new AccessFormatError(target === undefined
			? 'This design has no object of its own.'
			: `This design has no object named ${target}.`);
	}
	const object = design.objects[at];
	const schema = schemaOf(object);
	const slots = schema === undefined ? undefined : PROPERTY_SLOTS.get(schema);
	if (!slots) {
		throw new AccessFormatError(`No property slots were measured for a ${schema ?? 'design'}.`);
	}
	const slot = slots.get(name);
	if (!slot) {
		throw new AccessFormatError(
			`A ${schema} has no ${name} to set; it has: ${[...slots.keys()].sort().join(', ')}.`,
		);
	}
	const raw = encodeDesignProperty(slot, value);
	const code = slot[1];
	let kept = object.records.filter((entry) => entry.code !== code);
	const existing = object.records.find((entry) => entry.code === code);
	// Keep the record where it is: its id is the schema's answer for this
	// object, whatever the table says.
	const replaced: AccessDesignRecord = existing
		? { ...existing, value: raw }
		: { id: slot[0], code, valueType: slot[2], width: slot[3], value: raw };

	if (name === 'RowSourceType') {
		// The text is what the property sheet shows; what Access acts on is a
		// one-byte companion record.
		kept = kept.filter((entry) => entry.code !== ROW_SOURCE_KIND_CODE);
		const companion = ROW_SOURCE_KINDS.get(String(value));
		if (companion !== undefined) {
			kept.push(record(ROW_SOURCE_KIND_SLOT, Buffer.from([companion])));
		}
	}
	for (const [extraName, setting] of companionsOf(schema, name, value)) {
		const extraSlot = slots.get(extraName);
		if (!extraSlot) {
			continue;
		}
		kept = kept.filter((entry) => entry.code !== extraSlot[1]);
		const found = object.records.find((entry) => entry.code === extraSlot[1]);
		const extraRaw = encodeDesignProperty(extraSlot, setting);
		kept.push(found
			? { ...found, value: extraRaw }
			: {
				id: extraSlot[0], code: extraSlot[1], valueType: extraSlot[2],
				width: extraSlot[3], value: extraRaw,
			});
	}
	const records = [...kept, replaced].sort((a, b) => a.id - b.id);
	const objects = [...design.objects];
	objects[at] = { ...object, records };
	return { ...design, objects };
}

/**
 * The markers a section's controls carry, which depend on how many there are.
 * One control is a single child, `0xFE <type>`; two or more open a group whose
 * first is `0xFF <count> <type>` and the rest `0xFD <type>`.
 */
function placed(controls: AccessDesignObject[]): AccessDesignObject[] {
	return controls.map((control, position) => {
		if (controls.length === 1) {
			return { ...control, marker: OPEN_SECTION, code: control.type };
		}
		if (position === 0) {
			return { ...control, marker: OPEN_CONTROL, code: controls.length };
		}
		return { ...control, marker: OPEN_SIBLING, code: control.type };
	});
}

/**
 * How many objects directly follow `objects[at]` as its own children, theirs
 * included. A marker says how its object was opened, so the run after a
 * control belongs to it when it opens with `0xFE` or `0xFF <count>`, and to
 * the level above when it opens with `0xFD`.
 */
function childrenSpan(objects: AccessDesignObject[], at: number, stop: number): number {
	const first = at + 1;
	if (first >= stop) {
		return 0;
	}
	const marker = objects[first].marker;
	let count: number;
	if (marker === OPEN_SECTION) {
		count = 1;
	} else if (marker === OPEN_CONTROL) {
		const code = objects[first].code;
		count = typeof code === 'number' && code > 0 ? code : 1;
	} else {
		return 0;
	}
	let walked = first;
	for (let i = 0; i < count && walked < stop; i += 1) {
		walked += 1 + childrenSpan(objects, walked, stop);
	}
	return walked - first;
}

/**
 * Each control directly inside the run, as its index and the children it owns.
 * The count a section's opening marker carries is of its own controls, not of
 * everything beneath them.
 */
function topLevel(
	objects: AccessDesignObject[], start: number, stop: number,
): Array<{ index: number; owned: number }> {
	if (start >= stop) {
		return [];
	}
	let count = 1;
	if (objects[start].marker === OPEN_CONTROL) {
		const code = objects[start].code;
		count = typeof code === 'number' && code > 0 ? code : 1;
	}
	const out: Array<{ index: number; owned: number }> = [];
	let at = start;
	for (let i = 0; i < count && at < stop; i += 1) {
		const owned = childrenSpan(objects, at, stop);
		out.push({ index: at, owned });
		at += 1 + owned;
	}
	return out;
}

/** The control-defaults objects a design carries, by control type. */
export type AccessDesignPrototypes = ReadonlyMap<number, readonly AccessDesignRecord[]>;

/**
 * The design with the control-defaults object for a type in place: inserted in
 * type order among the ones there, the run of top-level objects re-marked as
 * one group.
 *
 * Access reads a control's themed properties against these; without the object
 * it ignores a button's `UseTheme`, colour indexes and gradient, and drops them
 * on its next save.
 */
function withPrototype(
	objects: AccessDesignObject[],
	typeCode: number,
	prototypes: AccessDesignPrototypes,
): AccessDesignObject[] {
	const firstSection = objects.findIndex(isAccessDesignSection);
	const end = firstSection < 0 ? objects.length : firstSection;
	const ahead = objects.slice(1, end);
	if (DEFAULTS_FREE.has(typeCode) || ahead.some((object) => object.type === typeCode)) {
		return objects;
	}
	const records = prototypes.get(typeCode);
	if (!records) {
		const name = CONTROL_TYPES.get(typeCode) ?? typeCode;
		throw new AccessFormatError(`No control defaults were captured for a ${name}.`);
	}
	const fresh: AccessDesignObject = { type: typeCode, records: [...records] };
	const at = ahead.findIndex((object) => (object.type ?? 0) > typeCode);
	ahead.splice(at < 0 ? ahead.length : at, 0, fresh);
	const sections = objects
		.map((object, index) => ({ object, index }))
		.filter((entry) => isAccessDesignSection(entry.object));
	const top = placed([...ahead, ...sections.map((entry) => entry.object)]);
	const rest = objects.slice(end);
	sections.forEach((entry, position) => {
		rest[entry.index - end] = top[ahead.length + position];
	});
	return [objects[0], ...top.slice(0, ahead.length), ...rest];
}

/** One control, as the records Access writes for a new one. */
export function designControlObject(
	controlType: string,
	name: string,
	guid: Buffer,
	geometry: { left: number; top: number; width: number; height: number },
	caption?: string,
	tabIndex = 0,
): AccessDesignObject {
	const base = CONTROL_SLOTS.get(controlType);
	if (!base) {
		if (READ_ONLY_TYPES.has(controlType)) {
			throw new AccessFormatError(
				`A ${controlType} is read but not written: it is part of a navigation control, `
				+ 'which names a sibling subform and the buttons beside it, so one written alone '
				+ 'would not work.',
			);
		}
		throw new AccessFormatError(
			`A ${controlType} cannot be written yet; known: `
			+ `${[...CONTROL_SLOTS.keys()].sort().join(', ')}.`,
		);
	}
	const slots = new Map(base);
	// The control Access made for the measurement was the first that takes the
	// focus, so it carried no TabIndex; every later one does.
	const measured = PROPERTY_SLOTS.get(controlType);
	const tabSlot = measured?.get('TabIndex');
	if (tabSlot && !slots.has('TabIndex')) {
		slots.set('TabIndex', tabSlot.slice(0, 4));
	}
	const u16 = (value: number): Buffer => {
		const out = Buffer.alloc(2);
		out.writeUInt16LE(value, 0);
		return out;
	};
	const u32 = (value: number): Buffer => {
		const out = Buffer.alloc(4);
		out.writeUInt32LE(value, 0);
		return out;
	};
	const values = new Map<string, Buffer>([
		['OverlapFlags', Buffer.from([OVERLAP_FLAGS.get(controlType) ?? DEFAULT_OVERLAP])],
		['IMESentenceMode', Buffer.from([DEFAULT_IME_SENTENCE_MODE])],
		['Left', u16(geometry.left)],
		['Top', u16(geometry.top)],
		['Width', u16(geometry.width)],
		['Height', u16(geometry.height)],
		['Name', Buffer.from(name, 'utf16le')],
		['RowSourceType', Buffer.from(DEFAULT_ROW_SOURCE_TYPE, 'utf16le')],
		['GUID', Buffer.from(guid)],
		['Picture', Buffer.from(NO_PICTURE)],
		['LayoutCachedLeft', u32(geometry.left)],
		['LayoutCachedTop', u32(geometry.top)],
		['LayoutCachedWidth', u32(geometry.left + geometry.width)],
		['LayoutCachedHeight', u32(geometry.top + geometry.height)],
	]);
	if (caption !== undefined) {
		values.set(
			slots.has('Caption') ? 'Caption' : 'ControlSource', Buffer.from(caption, 'utf16le'),
		);
	}
	if (tabIndex && TABBABLE.has(controlType)) {
		values.set('TabIndex', u16(tabIndex));
	}
	for (const [code, number] of BUTTON_EXTRAS.get(controlType) ?? []) {
		values.set(`Unidentified${code}`, u32(number));
	}
	for (const [key, bytes] of TYPE_EXTRAS.get(controlType) ?? []) {
		values.set(key, Buffer.from(bytes));
	}
	const records = [...values]
		.filter(([key]) => slots.has(key))
		.sort((a, b) => slots.get(a[0])![0] - slots.get(b[0])![0])
		.map(([key, value]) => record(slots.get(key)!, value));
	return { type: TYPE_CODES.get(controlType)!, records };
}

/**
 * A design with one more control on it. A control belongs to a section and is
 * written immediately after it; `parent` names a tab control the new page
 * joins instead.
 */
export function addDesignControl(
	design: AccessDesign,
	controlType: string,
	name: string,
	guid: Buffer,
	prototypes: AccessDesignPrototypes,
	options: {
		section?: string;
		parent?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
		caption?: string;
	} = {},
): AccessDesign {
	const section = options.section ?? 'Detail';
	let objects = [...design.objects];
	if (objects.some((object) => nameOf(object) === name)) {
		throw new AccessFormatError(`This design already has an object named ${name}.`);
	}
	if ((controlType === PAGE) !== (options.parent !== undefined)) {
		throw new AccessFormatError(
			`A ${PAGE} needs a parent ${PAGE_HOLDER} and nothing else takes one.`,
		);
	}
	if (!objects.some((object) => isAccessDesignSection(object) && nameOf(object) === section)) {
		throw new AccessFormatError(`This design has no ${section} section.`);
	}
	const control = designControlObject(
		controlType,
		name,
		guid,
		{
			left: options.left ?? 0,
			top: options.top ?? 0,
			width: options.width ?? 1440,
			height: options.height ?? 240,
		},
		options.caption,
		objects.filter((object) => nameOf(object) !== undefined && object.type !== undefined
			&& TABBABLE.has(CONTROL_TYPES.get(object.type) ?? '')).length,
	);
	// The first control of a type brings the type's control-defaults object
	// ahead of the sections, which moves the sections along.
	objects = withPrototype(objects, TYPE_CODES.get(controlType)!, prototypes);
	const at = objects.findIndex(
		(object) => isAccessDesignSection(object) && nameOf(object) === section,
	);
	let end = at + 1;
	while (end < objects.length && !isAccessDesignSection(objects[end])) {
		end += 1;
	}
	if (options.parent !== undefined) {
		return { ...design, objects: nested(objects, at + 1, end, options.parent, control) };
	}
	const owners = topLevel(objects, at + 1, end);
	const marked = placed([...owners.map((owner) => objects[owner.index]), control]);
	const rebuilt: AccessDesignObject[] = [];
	owners.forEach((owner, position) => {
		rebuilt.push(marked[position]);
		rebuilt.push(...objects.slice(owner.index + 1, owner.index + 1 + owner.owned));
	});
	rebuilt.push(marked[marked.length - 1]);
	return {
		...design,
		objects: [...objects.slice(0, at + 1), ...rebuilt, ...objects.slice(end)],
	};
}

/** The design with a control added to what a tab control holds. */
function nested(
	objects: AccessDesignObject[],
	start: number,
	stop: number,
	parent: string,
	control: AccessDesignObject,
): AccessDesignObject[] {
	const owners = topLevel(objects, start, stop);
	const found = owners.find((owner) => nameOf(objects[owner.index]) === parent);
	if (!found) {
		throw new AccessFormatError(`This design has no control named ${parent} in that section.`);
	}
	const holder = objects[found.index];
	if (holder.type === undefined || CONTROL_TYPES.get(holder.type) !== PAGE_HOLDER) {
		throw new AccessFormatError(`${parent} is not a ${PAGE_HOLDER}, so it holds no ${PAGE}.`);
	}
	const children = placed([
		...objects.slice(found.index + 1, found.index + 1 + found.owned), control,
	]);
	return [
		...objects.slice(0, found.index + 1),
		...children,
		...objects.slice(found.index + 1 + found.owned),
	];
}

/** Where a control sits in the tab order, or undefined for one that never takes the focus. */
function tabIndexOf(object: AccessDesignObject, tabCode: number): number | undefined {
	if (object.type === undefined || !TABBABLE.has(CONTROL_TYPES.get(object.type) ?? '')) {
		return undefined;
	}
	const raw = recordValue(object, tabCode);
	return raw && raw.length > 0 ? raw.readUIntLE(0, raw.length) : 0;
}

/**
 * The object with its tab index one lower when it followed the removed
 * control; one that lands on 0 loses the record, as Access writes the first
 * control without one.
 */
function closedUp(
	object: AccessDesignObject, removed: number, tabCode: number,
): AccessDesignObject {
	const raw = recordValue(object, tabCode);
	if (object.type === undefined || !raw || raw.length === 0) {
		return object;
	}
	const index = raw.readUIntLE(0, raw.length);
	if (index <= removed) {
		return object;
	}
	const records = object.records
		.filter((entry) => !(entry.code === tabCode && index - 1 === 0))
		.map((entry) => {
			if (entry.code !== tabCode) {
				return entry;
			}
			const value = Buffer.alloc(entry.value.length);
			value.writeUIntLE(index - 1, 0, entry.value.length);
			return { ...entry, value };
		});
	return { ...object, records };
}

/**
 * A design with the named control taken off, and whatever it holds. The
 * controls left at that level are re-marked, and the tab order closes up
 * behind a control that took the focus.
 */
export function removeDesignControl(design: AccessDesign, name: string): AccessDesign {
	const objects = [...design.objects];
	const tabCode = PROPERTY_CODES.get('TabIndex');
	let result: AccessDesignObject[] | undefined;
	let removedTab: number | undefined;
	let at = 1;
	while (at < objects.length && result === undefined) {
		if (!isAccessDesignSection(objects[at])) {
			at += 1;
			continue;
		}
		let end = at + 1;
		while (end < objects.length && !isAccessDesignSection(objects[end])) {
			end += 1;
		}
		const owners = topLevel(objects, at + 1, end);
		const hit = owners.find((owner) => nameOf(objects[owner.index]) === name);
		if (hit) {
			removedTab = tabCode === undefined
				? undefined
				: tabIndexOf(objects[hit.index], tabCode);
			const rest = owners.filter((owner) => owner.index !== hit.index);
			const marked = placed(rest.map((owner) => objects[owner.index]));
			const rebuilt: AccessDesignObject[] = [];
			rest.forEach((owner, position) => {
				rebuilt.push(marked[position]);
				rebuilt.push(...objects.slice(owner.index + 1, owner.index + 1 + owner.owned));
			});
			result = [...objects.slice(0, at + 1), ...rebuilt, ...objects.slice(end)];
			break;
		}
		for (const owner of owners) {
			const children = objects.slice(owner.index + 1, owner.index + 1 + owner.owned);
			const child = children.findIndex((object) => nameOf(object) === name);
			if (child < 0) {
				continue;
			}
			if (childrenSpan(objects, owner.index + 1 + child, owner.index + 1 + owner.owned)) {
				throw new AccessFormatError(
					`${name} holds controls of its own; take those off first.`,
				);
			}
			removedTab = tabCode === undefined ? undefined : tabIndexOf(children[child], tabCode);
			const kept = placed(children.filter((_object, k) => k !== child));
			result = [
				...objects.slice(0, owner.index + 1), ...kept,
				...objects.slice(owner.index + 1 + owner.owned),
			];
			break;
		}
		at = end;
	}
	if (result === undefined) {
		throw new AccessFormatError(`This design has no control named ${name}.`);
	}
	if (removedTab !== undefined && tabCode !== undefined) {
		const removed = removedTab;
		result = result.map((object) => closedUp(object, removed, tabCode));
	}
	return { ...design, objects: result };
}

/**
 * What holds each control: its section, or the control it sits inside, as
 * indexes into the object list.
 */
export function designObjectHolders(objects: AccessDesignObject[]): Map<number, number> {
	const holders = new Map<number, number>();
	const walk = (start: number, stop: number, holder: number): void => {
		for (const owner of topLevel(objects, start, stop)) {
			holders.set(owner.index, holder);
			walk(owner.index + 1, owner.index + 1 + owner.owned, owner.index);
		}
	};
	let at = 1;
	while (at < objects.length) {
		if (!isAccessDesignSection(objects[at])) {
			at += 1;
			continue;
		}
		let end = at + 1;
		while (end < objects.length && !isAccessDesignSection(objects[end])) {
			end += 1;
		}
		walk(at + 1, end, at);
		at = end;
	}
	return holders;
}

/** The control-defaults objects a design already carries, by type. */
export function designPrototypes(design: AccessDesign): AccessDesignPrototypes {
	const out = new Map<number, readonly AccessDesignRecord[]>();
	for (const object of design.objects.slice(1)) {
		if (isAccessDesignSection(object)) {
			break;
		}
		if (object.type !== undefined) {
			out.set(object.type, object.records);
		}
	}
	return out;
}

/** Parse, edit and rebuild in one step, for a caller holding only bytes. */
export function editDesignBlob(
	blob: Buffer,
	edit: (design: AccessDesign) => AccessDesign,
): Buffer {
	return buildAccessDesign(edit(parseAccessDesign(blob)));
}
