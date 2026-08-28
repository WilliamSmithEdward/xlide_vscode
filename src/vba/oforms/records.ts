// [MS-OFORMS] embedded-control records, read and written losslessly.
//
// Every control record shares one skeleton: a four-byte header (minor 0x00,
// major 0x02, cbRecord), a property mask, a DataBlock of mask-selected fields
// at most 4 bytes each in a fixed order with self-size alignment, an
// ExtraDataBlock of the larger values (strings, sizes, positions), then
// outside cbRecord the StreamData (pictures), TextProps, and a per-type tail.
// So each control is a TABLE here, and one engine reads and writes them all.
//
// Lossless by construction: alignment padding is captured and replayed
// (its bytes are undefined, so recomputing them would break byte identity),
// string bytes are kept raw beside their decoded text, pictures are kept as
// opaque byte runs, and any tail this engine does not model is preserved
// verbatim. An unmodified parse therefore serializes to the identical bytes,
// which is what the fixture suite pins.

import { OformsReader, OformsWriter } from './bytes';

/** How a DataBlock field's bytes are interpreted. */
export type DataFieldKind =
	| 'u'         // unsigned integer of `size`
	| 'i'         // signed integer of `size`
	| 'lenBytes'  // CountOfBytesWithCompressionFlag for an ExtraDataBlock string
	| 'marker';   // 0xFFFF placeholder for StreamData content (pictures, fonts)

export interface DataFieldSpec {
	bit: number;
	name: string;
	size: 1 | 2 | 4;
	kind: DataFieldKind;
}

export interface ExtraFieldSpec {
	bit: number;
	name: string;
	kind:
		| 'size8'   // fmSize: width, height (HIMETRIC)
		| 'pos8'    // fmPosition: left, top (HIMETRIC)
		| 'str'     // fmString sized by the DataBlock field of the same name
		| 'arrRaw'; // TabStrip string array sized by `sizeFrom`, preserved raw
	sizeFrom?: string;
}

export interface RecordSpec {
	type: string;
	/** Header major version; control records are 0x02, the FormControl 0x04. */
	major?: number;
	/** Stop at the cb boundary and leave the rest to the caller (FormControl). */
	stopAfterExtra?: boolean;
	mask64?: boolean;
	data: readonly DataFieldSpec[];
	extra: readonly ExtraFieldSpec[];
	/** StreamData members in order; each present when its mask bit is set. */
	stream: ReadonlyArray<{ bit: number; name: string }>;
	textProps?: boolean;
	/** Bytes after TextProps (rgColumnInfo, TabStripTabFlags): preserved raw. */
	rawTail?: boolean;
}

export interface ParsedString {
	text: string;
	compressed: boolean;
	/** The exact stored bytes, replayed on write unless the text was edited. */
	raw: Buffer;
	edited?: boolean;
}

export interface ParsedRecord {
	spec: RecordSpec;
	maskLo: number;
	maskHi: number;
	values: Map<string, number>;
	strings: Map<string, ParsedString>;
	sizes: Map<string, { width: number; height: number }>;
	arrays: Map<string, Buffer>;
	pads: Map<string, Buffer>;
	streamData: Map<string, Buffer>;
	textProps?: ParsedRecord;
	tailRaw?: Buffer;
}

/** Decodes/encodes MBCS compressed strings for the project's code page. */
export interface OformsTextCodec {
	decode(bytes: Buffer, compressed: boolean): string;
	encode(text: string, compressed: boolean): Buffer;
}

/** A codec for the common case: compressed bytes are Latin-1/ASCII. */
export const LATIN1_CODEC: OformsTextCodec = {
	decode: (bytes, compressed) => bytes.toString(compressed ? 'latin1' : 'utf16le'),
	encode: (text, compressed) => Buffer.from(text, compressed ? 'latin1' : 'utf16le'),
};

// ---------------------------------------------------------------- tables
//
// Field order is the DataBlock order from each record's diagram, which is
// also mask-bit order; mask-only booleans carry no DataBlock field and are
// simply absent from these tables.

export const TEXT_PROPS_SPEC: RecordSpec = {
	type: 'TextProps',
	// TextProps ends at its cb boundary; whatever follows (a TabStrip's tab
	// flags, MorphData's column info) belongs to the record that embeds it.
	stopAfterExtra: true,
	data: [
		{ bit: 0, name: 'FontName', size: 4, kind: 'lenBytes' },
		{ bit: 1, name: 'FontEffects', size: 4, kind: 'u' },
		{ bit: 2, name: 'FontHeight', size: 4, kind: 'u' },
		{ bit: 4, name: 'FontCharSet', size: 1, kind: 'u' },
		{ bit: 5, name: 'FontPitchAndFamily', size: 1, kind: 'u' },
		{ bit: 6, name: 'ParagraphAlign', size: 1, kind: 'u' },
		{ bit: 7, name: 'FontWeight', size: 2, kind: 'u' },
	],
	extra: [{ bit: 0, name: 'FontName', kind: 'str' }],
	stream: [],
};

const MORPH_DATA_SPEC: RecordSpec = {
	type: 'MorphData',
	mask64: true,
	data: [
		{ bit: 0, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 3, name: 'MaxLength', size: 4, kind: 'u' },
		{ bit: 4, name: 'BorderStyle', size: 1, kind: 'u' },
		{ bit: 5, name: 'ScrollBars', size: 1, kind: 'u' },
		{ bit: 6, name: 'DisplayStyle', size: 1, kind: 'u' },
		{ bit: 7, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 9, name: 'PasswordChar', size: 2, kind: 'u' },
		{ bit: 10, name: 'ListWidth', size: 4, kind: 'u' },
		{ bit: 11, name: 'BoundColumn', size: 2, kind: 'u' },
		{ bit: 12, name: 'TextColumn', size: 2, kind: 'i' },
		{ bit: 13, name: 'ColumnCount', size: 2, kind: 'i' },
		{ bit: 14, name: 'ListRows', size: 2, kind: 'u' },
		{ bit: 15, name: 'cColumnInfo', size: 2, kind: 'u' },
		{ bit: 16, name: 'MatchEntry', size: 1, kind: 'u' },
		{ bit: 17, name: 'ListStyle', size: 1, kind: 'u' },
		{ bit: 18, name: 'ShowDropButtonWhen', size: 1, kind: 'u' },
		{ bit: 20, name: 'DropButtonStyle', size: 1, kind: 'u' },
		{ bit: 21, name: 'MultiSelect', size: 1, kind: 'u' },
		{ bit: 22, name: 'Value', size: 4, kind: 'lenBytes' },
		{ bit: 23, name: 'Caption', size: 4, kind: 'lenBytes' },
		{ bit: 24, name: 'PicturePosition', size: 4, kind: 'u' },
		{ bit: 25, name: 'BorderColor', size: 4, kind: 'u' },
		{ bit: 26, name: 'SpecialEffect', size: 4, kind: 'u' },
		{ bit: 27, name: 'MouseIcon', size: 2, kind: 'marker' },
		{ bit: 28, name: 'Picture', size: 2, kind: 'marker' },
		{ bit: 29, name: 'Accelerator', size: 2, kind: 'u' },
		{ bit: 32, name: 'GroupName', size: 4, kind: 'lenBytes' },
	],
	extra: [
		{ bit: 8, name: 'Size', kind: 'size8' },
		{ bit: 22, name: 'Value', kind: 'str' },
		{ bit: 23, name: 'Caption', kind: 'str' },
		{ bit: 32, name: 'GroupName', kind: 'str' },
	],
	stream: [
		{ bit: 27, name: 'MouseIcon' },
		{ bit: 28, name: 'Picture' },
	],
	textProps: true,
	rawTail: true, // rgColumnInfo for ComboBox/ListBox
};

const COMMAND_BUTTON_SPEC: RecordSpec = {
	type: 'CommandButton',
	data: [
		{ bit: 0, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 3, name: 'Caption', size: 4, kind: 'lenBytes' },
		{ bit: 4, name: 'PicturePosition', size: 4, kind: 'u' },
		{ bit: 6, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 7, name: 'Picture', size: 2, kind: 'marker' },
		{ bit: 8, name: 'Accelerator', size: 2, kind: 'u' },
		{ bit: 10, name: 'MouseIcon', size: 2, kind: 'marker' },
	],
	extra: [
		{ bit: 3, name: 'Caption', kind: 'str' },
		{ bit: 5, name: 'Size', kind: 'size8' },
	],
	stream: [
		{ bit: 7, name: 'Picture' },
		{ bit: 10, name: 'MouseIcon' },
	],
	textProps: true,
};

const LABEL_SPEC: RecordSpec = {
	type: 'Label',
	data: [
		{ bit: 0, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 3, name: 'Caption', size: 4, kind: 'lenBytes' },
		{ bit: 4, name: 'PicturePosition', size: 4, kind: 'u' },
		{ bit: 6, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 7, name: 'BorderColor', size: 4, kind: 'u' },
		{ bit: 8, name: 'BorderStyle', size: 2, kind: 'u' },
		{ bit: 9, name: 'SpecialEffect', size: 2, kind: 'u' },
		{ bit: 10, name: 'Picture', size: 2, kind: 'marker' },
		{ bit: 11, name: 'Accelerator', size: 2, kind: 'u' },
		{ bit: 12, name: 'MouseIcon', size: 2, kind: 'marker' },
	],
	extra: [
		{ bit: 3, name: 'Caption', kind: 'str' },
		{ bit: 5, name: 'Size', kind: 'size8' },
	],
	stream: [
		{ bit: 10, name: 'Picture' },
		{ bit: 12, name: 'MouseIcon' },
	],
	textProps: true,
};

const IMAGE_SPEC: RecordSpec = {
	type: 'Image',
	data: [
		{ bit: 3, name: 'BorderColor', size: 4, kind: 'u' },
		{ bit: 4, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 5, name: 'BorderStyle', size: 1, kind: 'u' },
		{ bit: 6, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 7, name: 'PictureSizeMode', size: 1, kind: 'u' },
		{ bit: 8, name: 'SpecialEffect', size: 1, kind: 'u' },
		{ bit: 10, name: 'Picture', size: 2, kind: 'marker' },
		{ bit: 11, name: 'PictureAlignment', size: 1, kind: 'u' },
		{ bit: 13, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 14, name: 'MouseIcon', size: 2, kind: 'marker' },
	],
	extra: [{ bit: 9, name: 'Size', kind: 'size8' }],
	stream: [
		{ bit: 10, name: 'Picture' },
		{ bit: 14, name: 'MouseIcon' },
	],
	// Image carries no TextProps ([MS-OFORMS] 2.3.1 applies-to list).
};

const SPIN_BUTTON_SPEC: RecordSpec = {
	type: 'SpinButton',
	data: [
		{ bit: 0, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 5, name: 'Min', size: 4, kind: 'i' },
		{ bit: 6, name: 'Max', size: 4, kind: 'i' },
		{ bit: 7, name: 'Position', size: 4, kind: 'i' },
		{ bit: 8, name: 'PrevEnabled', size: 4, kind: 'i' },
		{ bit: 9, name: 'NextEnabled', size: 4, kind: 'i' },
		{ bit: 10, name: 'SmallChange', size: 4, kind: 'i' },
		{ bit: 11, name: 'Orientation', size: 4, kind: 'i' },
		{ bit: 12, name: 'Delay', size: 4, kind: 'u' },
		{ bit: 13, name: 'MouseIcon', size: 2, kind: 'marker' },
		{ bit: 14, name: 'MousePointer', size: 1, kind: 'u' },
	],
	extra: [{ bit: 3, name: 'Size', kind: 'size8' }],
	stream: [{ bit: 13, name: 'MouseIcon' }],
	// SpinButton carries no TextProps.
};

const SCROLL_BAR_SPEC: RecordSpec = {
	type: 'ScrollBar',
	data: [
		{ bit: 0, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 4, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 5, name: 'Min', size: 4, kind: 'i' },
		{ bit: 6, name: 'Max', size: 4, kind: 'i' },
		{ bit: 7, name: 'Position', size: 4, kind: 'i' },
		{ bit: 9, name: 'PrevEnabled', size: 4, kind: 'i' },
		{ bit: 10, name: 'NextEnabled', size: 4, kind: 'i' },
		{ bit: 11, name: 'SmallChange', size: 4, kind: 'i' },
		{ bit: 12, name: 'LargeChange', size: 4, kind: 'i' },
		{ bit: 13, name: 'Orientation', size: 4, kind: 'i' },
		{ bit: 14, name: 'ProportionalThumb', size: 2, kind: 'i' },
		{ bit: 15, name: 'Delay', size: 4, kind: 'u' },
		{ bit: 16, name: 'MouseIcon', size: 2, kind: 'marker' },
	],
	extra: [{ bit: 3, name: 'Size', kind: 'size8' }],
	stream: [{ bit: 16, name: 'MouseIcon' }],
	// ScrollBar carries no TextProps.
};

const TAB_STRIP_SPEC: RecordSpec = {
	type: 'TabStrip',
	data: [
		{ bit: 0, name: 'ListIndex', size: 4, kind: 'i' },
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 5, name: 'ItemsSize', size: 4, kind: 'u' },
		{ bit: 6, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 8, name: 'TabOrientation', size: 4, kind: 'u' },
		{ bit: 9, name: 'TabStyle', size: 4, kind: 'u' },
		{ bit: 11, name: 'TabFixedWidth', size: 4, kind: 'u' },
		{ bit: 12, name: 'TabFixedHeight', size: 4, kind: 'u' },
		{ bit: 15, name: 'TipStringsSize', size: 4, kind: 'u' },
		{ bit: 17, name: 'NamesSize', size: 4, kind: 'u' },
		{ bit: 18, name: 'VariousPropertyBits', size: 4, kind: 'u' },
		{ bit: 20, name: 'TabsAllocated', size: 4, kind: 'u' },
		{ bit: 21, name: 'TagsSize', size: 4, kind: 'u' },
		{ bit: 22, name: 'TabData', size: 4, kind: 'u' },
		{ bit: 23, name: 'AcceleratorsSize', size: 4, kind: 'u' },
		{ bit: 24, name: 'MouseIcon', size: 2, kind: 'marker' },
	],
	extra: [
		{ bit: 4, name: 'Size', kind: 'size8' },
		{ bit: 5, name: 'Items', kind: 'arrRaw', sizeFrom: 'ItemsSize' },
		{ bit: 15, name: 'TipStrings', kind: 'arrRaw', sizeFrom: 'TipStringsSize' },
		{ bit: 17, name: 'TabNames', kind: 'arrRaw', sizeFrom: 'NamesSize' },
		{ bit: 21, name: 'Tags', kind: 'arrRaw', sizeFrom: 'TagsSize' },
		{ bit: 23, name: 'Accelerators', kind: 'arrRaw', sizeFrom: 'AcceleratorsSize' },
	],
	stream: [{ bit: 24, name: 'MouseIcon' }],
	textProps: true,
	rawTail: true, // TabStripTabFlags
};

/** ClsidCacheIndex -> the record persisted for it in an object stream. */
export const RECORD_SPECS_BY_CACHE_INDEX: ReadonlyMap<number, RecordSpec> = new Map([
	[12, IMAGE_SPEC],
	[15, MORPH_DATA_SPEC],
	[16, SPIN_BUTTON_SPEC],
	[17, COMMAND_BUTTON_SPEC],
	[18, TAB_STRIP_SPEC],
	[21, LABEL_SPEC],
	[23, MORPH_DATA_SPEC], // TextBox
	[24, MORPH_DATA_SPEC], // ListBox
	[25, MORPH_DATA_SPEC], // ComboBox
	[26, MORPH_DATA_SPEC], // CheckBox
	[27, MORPH_DATA_SPEC], // OptionButton
	[28, MORPH_DATA_SPEC], // ToggleButton
	[47, SCROLL_BAR_SPEC],
]);

/** Cached indices persisted as nested `iNN` storages rather than `o` records. */
export const CONTAINER_CACHE_INDICES: ReadonlySet<number> = new Set([7, 14, 57]);

const maskBit = (rec: Pick<ParsedRecord, 'maskLo' | 'maskHi'>, bit: number): boolean =>
	bit < 32 ? (rec.maskLo & (1 << bit)) !== 0 : (rec.maskHi & (1 << (bit - 32))) !== 0;

export function recordHas(rec: ParsedRecord, name: string): boolean {
	const field = rec.spec.data.find((f) => f.name === name)
		?? rec.spec.extra.find((f) => f.name === name);
	return field !== undefined && maskBit(rec, field.bit);
}

// ------------------------------------------------------------------ parse

/**
 * Parses one control record starting at `r.pos`, consuming the header, the
 * cb-bounded mask/data/extra, then StreamData, TextProps, and the raw tail up
 * to `end` (the caller's boundary: the site's ObjectStreamSize, or the
 * stream's length).
 */
export function parseRecord(
	r: OformsReader,
	spec: RecordSpec,
	codec: OformsTextCodec,
	end: number,
): ParsedRecord {
	const base = r.pos;
	const minor = r.u8();
	const major = r.u8();
	if (minor !== 0x00 || major !== (spec.major ?? 0x02)) {
		throw new RangeError(`${spec.type}: not a control record (version ${major}.${minor})`);
	}
	const cb = r.u16();
	const maskLo = r.u32();
	const maskHi = spec.mask64 ? r.u32() : 0;
	const rec: ParsedRecord = {
		spec, maskLo, maskHi,
		values: new Map(), strings: new Map(), sizes: new Map(), arrays: new Map(),
		pads: new Map(), streamData: new Map(),
	};

	// DataBlock, fields in table order, aligned to their own size.
	for (const f of spec.data) {
		if (!maskBit(rec, f.bit)) { continue; }
		const pad = r.pad(base, f.size);
		if (pad.length) { rec.pads.set(`before:${f.name}`, pad); }
		const v = f.size === 1 ? r.u8() : f.size === 2 ? (f.kind === 'i' ? r.i16() : r.u16()) : (f.kind === 'i' ? r.i32() : r.u32());
		rec.values.set(f.name, v);
	}
	const endPad = r.pad(base, 4);
	if (endPad.length) { rec.pads.set('data:end', endPad); }

	// ExtraDataBlock, in table order; strings pad to 4 after their bytes.
	for (const f of spec.extra) {
		if (!maskBit(rec, f.bit)) { continue; }
		if (f.kind === 'size8') {
			rec.sizes.set(f.name, { width: r.i32(), height: r.i32() });
		} else if (f.kind === 'pos8') {
			rec.sizes.set(f.name, { width: r.i32(), height: r.i32() });
		} else if (f.kind === 'str') {
			const lc = rec.values.get(f.name) ?? 0;
			const cbStr = lc & 0x7fffffff;
			const compressed = (lc & 0x80000000) !== 0;
			const raw = r.bytes(cbStr);
			rec.strings.set(f.name, { text: codec.decode(raw, compressed), compressed, raw });
			const pad = r.pad(base, 4);
			if (pad.length) { rec.pads.set(`str:${f.name}`, pad); }
		} else {
			const size = rec.values.get(f.sizeFrom ?? '') ?? 0;
			rec.arrays.set(f.name, r.bytes(size));
		}
	}

	const afterExtra = base + 4 + cb;
	if (spec.stopAfterExtra) {
		if (r.pos !== afterExtra) {
			throw new RangeError(`${spec.type}: cb mismatch (${r.pos - base - 4} != ${cb})`);
		}
		return rec;
	}
	if (r.pos !== afterExtra) {
		if (r.pos > afterExtra) {
			throw new RangeError(`${spec.type}: record overran cb by ${r.pos - afterExtra}`);
		}
		// Bytes inside cb this table did not account for would silently break
		// the write; refuse rather than guess.
		throw new RangeError(`${spec.type}: ${afterExtra - r.pos} unmodelled bytes inside cb`);
	}

	// StreamData: GuidAndPicture per masked member, preserved as bytes.
	for (const m of spec.stream) {
		if (!maskBit(rec, m.bit)) { continue; }
		rec.streamData.set(m.name, readGuidAndPicture(r));
	}

	if (spec.textProps) {
		rec.textProps = parseRecord(r, TEXT_PROPS_SPEC, codec, end);
	}

	if (r.pos < end) {
		if (!spec.rawTail) {
			throw new RangeError(`${spec.type}: ${end - r.pos} unexpected trailing bytes`);
		}
		rec.tailRaw = r.bytes(end - r.pos);
	}
	return rec;
}

/** A GuidAndPicture: 16-byte CLSID, 4-byte preamble, 4-byte size, data. */
function readGuidAndPicture(r: OformsReader): Buffer {
	const start = r.pos;
	r.bytes(16);
	r.u32(); // preamble
	const size = r.u32();
	r.bytes(size);
	const total = r.pos - start;
	r.pos = start;
	return r.bytes(total);
}

// -------------------------------------------------------------- serialize

export function serializeRecord(rec: ParsedRecord, codec: OformsTextCodec): Buffer {
	const spec = rec.spec;
	const w = new OformsWriter();
	const body = new OformsWriter();

	// Rebuild string length fields first: an edited string changes its
	// CountOfBytesWithCompressionFlag, and the DataBlock carries it.
	for (const f of spec.extra) {
		if (f.kind !== 'str' || !maskBit(rec, f.bit)) { continue; }
		const s = rec.strings.get(f.name);
		if (!s) { continue; }
		const raw = s.edited ? codec.encode(s.text, s.compressed) : s.raw;
		rec.values.set(f.name, (raw.length & 0x7fffffff) | (s.compressed ? 0x80000000 : 0));
	}

	// DataBlock relative to the record base: header(4) + mask(4 or 8).
	const headerSize = 4 + (spec.mask64 ? 8 : 4);
	const rel = (writer: OformsWriter): number => writer.position + headerSize;
	for (const f of spec.data) {
		if (!maskBit(rec, f.bit)) { continue; }
		const over = rel(body) % f.size;
		if (over !== 0) {
			const captured = rec.pads.get(`before:${f.name}`);
			const gap = f.size - over;
			body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
		}
		const v = rec.values.get(f.name) ?? 0;
		if (f.size === 1) { body.u8(v); }
		else if (f.size === 2) { f.kind === 'i' ? body.i16(v) : body.u16(v); }
		else { f.kind === 'i' ? body.i32(v) : body.u32(v); }
	}
	{
		const over = rel(body) % 4;
		if (over !== 0) {
			const captured = rec.pads.get('data:end');
			const gap = 4 - over;
			body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
		}
	}

	for (const f of spec.extra) {
		if (!maskBit(rec, f.bit)) { continue; }
		if (f.kind === 'size8' || f.kind === 'pos8') {
			const s = rec.sizes.get(f.name) ?? { width: 0, height: 0 };
			body.i32(s.width); body.i32(s.height);
		} else if (f.kind === 'str') {
			const s = rec.strings.get(f.name);
			const raw = s ? (s.edited ? codec.encode(s.text, s.compressed) : s.raw) : Buffer.alloc(0);
			body.bytes(raw);
			const over = rel(body) % 4;
			if (over !== 0) {
				const captured = rec.pads.get(`str:${f.name}`);
				const gap = 4 - over;
				body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
			}
		} else {
			body.bytes(rec.arrays.get(f.name) ?? Buffer.alloc(0));
		}
	}

	const bodyBytes = body.toBuffer();
	w.u8(0x00);
	w.u8(spec.major ?? 0x02);
	w.u16(bodyBytes.length + (spec.mask64 ? 8 : 4));
	w.u32(rec.maskLo >>> 0);
	if (spec.mask64) { w.u32(rec.maskHi >>> 0); }
	w.bytes(bodyBytes);

	for (const m of spec.stream) {
		if (!maskBit(rec, m.bit)) { continue; }
		const data = rec.streamData.get(m.name);
		if (!data) { throw new RangeError(`${spec.type}: masked StreamData ${m.name} has no bytes`); }
		w.bytes(data);
	}
	if (spec.textProps && rec.textProps) {
		w.bytes(serializeRecord(rec.textProps, codec));
	}
	if (rec.tailRaw) { w.bytes(rec.tailRaw); }
	return w.toBuffer();
}

// ------------------------------------------------------- record mutation

/** Sets or clears a numeric DataBlock field, adjusting the mask. */
export function setRecordValue(rec: ParsedRecord, name: string, value: number | undefined): void {
	const field = rec.spec.data.find((f) => f.name === name);
	if (!field) { throw new RangeError(`${rec.spec.type} has no field ${name}`); }
	applyMaskBit(rec, field.bit, value !== undefined);
	if (value === undefined) { rec.values.delete(name); }
	else { rec.values.set(name, value); }
}

/** Sets or clears a string property (Caption, Value, GroupName, FontName). */
export function setRecordString(rec: ParsedRecord, name: string, text: string | undefined): void {
	const extra = rec.spec.extra.find((f) => f.name === name && f.kind === 'str');
	if (!extra) { throw new RangeError(`${rec.spec.type} has no string ${name}`); }
	if (text === undefined) {
		applyMaskBit(rec, extra.bit, false);
		rec.strings.delete(name);
		rec.values.delete(name);
		return;
	}
	const existing = rec.strings.get(name);
	// New strings compress when every character fits a single byte, the same
	// choice the VBE makes; an existing string keeps its stored compression.
	const compressed = existing ? existing.compressed : [...text].every((c) => c.charCodeAt(0) <= 0xff);
	applyMaskBit(rec, extra.bit, true);
	rec.strings.set(name, { text, compressed, raw: Buffer.alloc(0), edited: true });
}

export function setRecordSize(rec: ParsedRecord, widthHimetric: number, heightHimetric: number): void {
	const extra = rec.spec.extra.find((f) => f.kind === 'size8');
	if (!extra) { throw new RangeError(`${rec.spec.type} has no Size`); }
	applyMaskBit(rec, extra.bit, true);
	rec.sizes.set(extra.name, { width: widthHimetric, height: heightHimetric });
}

function applyMaskBit(rec: ParsedRecord, bit: number, on: boolean): void {
	if (bit < 32) {
		rec.maskLo = on ? (rec.maskLo | (1 << bit)) >>> 0 : (rec.maskLo & ~(1 << bit)) >>> 0;
	} else {
		const b = bit - 32;
		rec.maskHi = on ? (rec.maskHi | (1 << b)) >>> 0 : (rec.maskHi & ~(1 << b)) >>> 0;
	}
}
