// The FormControl (`f` stream) and its object stream (`o`), parsed to a
// model that writes back byte-identically.
//
// A form's `f` is one record with a wider header (major 0x04) followed by
// three regions the generic engine does not cover: FormStreamData (mouse
// icon, font, picture blobs), FormSiteData (an optional class table, the
// site count, the run-length depth/type array, then one OleSiteConcreteControl
// per embedded control), and an optional DesignExtender tail. The `o` stream
// is the concatenation of each NON-container site's control record, in site
// order, each occupying exactly its site's ObjectStreamSize bytes.

import { OformsReader, OformsWriter } from './bytes';
import {
	LATIN1_CODEC,
	parseRecord,
	serializeRecord,
	CONTAINER_CACHE_INDICES,
	RECORD_SPECS_BY_CACHE_INDEX,
	type OformsTextCodec,
	type ParsedRecord,
	type ParsedString,
	type RecordSpec,
} from './records';

export const FORM_SPEC: RecordSpec = {
	type: 'Form',
	major: 0x04,
	stopAfterExtra: true,
	data: [
		{ bit: 1, name: 'BackColor', size: 4, kind: 'u' },
		{ bit: 2, name: 'ForeColor', size: 4, kind: 'u' },
		{ bit: 3, name: 'NextAvailableID', size: 4, kind: 'u' },
		{ bit: 6, name: 'BooleanProperties', size: 4, kind: 'u' },
		{ bit: 7, name: 'BorderStyle', size: 1, kind: 'u' },
		{ bit: 8, name: 'MousePointer', size: 1, kind: 'u' },
		{ bit: 9, name: 'ScrollBars', size: 1, kind: 'u' },
		{ bit: 13, name: 'GroupCnt', size: 4, kind: 'i' },
		{ bit: 15, name: 'MouseIcon', size: 2, kind: 'marker' },
		{ bit: 16, name: 'Cycle', size: 1, kind: 'u' },
		{ bit: 17, name: 'SpecialEffect', size: 1, kind: 'u' },
		{ bit: 18, name: 'BorderColor', size: 4, kind: 'u' },
		{ bit: 19, name: 'Caption', size: 4, kind: 'lenBytes' },
		{ bit: 20, name: 'Font', size: 2, kind: 'marker' },
		{ bit: 21, name: 'Picture', size: 2, kind: 'marker' },
		{ bit: 22, name: 'Zoom', size: 4, kind: 'u' },
		{ bit: 23, name: 'PictureAlignment', size: 1, kind: 'u' },
		{ bit: 25, name: 'PictureSizeMode', size: 1, kind: 'u' },
		{ bit: 26, name: 'ShapeCookie', size: 4, kind: 'u' },
		{ bit: 27, name: 'DrawBuffer', size: 4, kind: 'u' },
	],
	extra: [
		{ bit: 10, name: 'DisplayedSize', kind: 'size8' },
		{ bit: 11, name: 'LogicalSize', kind: 'size8' },
		{ bit: 12, name: 'ScrollPosition', kind: 'pos8' },
		{ bit: 19, name: 'Caption', kind: 'str' },
	],
	stream: [],
};

/** Site mask bits, DataBlock order ([MS-OFORMS] 2.2.10.12.3). */
const SITE_DATA_FIELDS: ReadonlyArray<{ bit: number; name: string; size: 2 | 4; signed?: boolean }> = [
	{ bit: 0, name: 'NameData', size: 4 },
	{ bit: 1, name: 'TagData', size: 4 },
	{ bit: 2, name: 'ID', size: 4, signed: true },
	{ bit: 3, name: 'HelpContextID', size: 4, signed: true },
	{ bit: 4, name: 'BitFlags', size: 4 },
	{ bit: 5, name: 'ObjectStreamSize', size: 4 },
	{ bit: 6, name: 'TabIndex', size: 2, signed: true },
	{ bit: 7, name: 'ClsidCacheIndex', size: 2 },
	{ bit: 9, name: 'GroupID', size: 2 },
	{ bit: 11, name: 'ControlTipTextData', size: 4 },
	{ bit: 12, name: 'RuntimeLicKeyData', size: 4 },
	{ bit: 13, name: 'ControlSourceData', size: 4 },
	{ bit: 14, name: 'RowSourceData', size: 4 },
];

/** Extra-block strings, in order, each sized by its DataBlock length field. */
const SITE_STRINGS: ReadonlyArray<{ lenField: string; name: string }> = [
	{ lenField: 'NameData', name: 'Name' },
	{ lenField: 'TagData', name: 'Tag' },
];
const SITE_STRINGS_AFTER_POSITION: ReadonlyArray<{ lenField: string; name: string }> = [
	{ lenField: 'ControlTipTextData', name: 'ControlTipText' },
	{ lenField: 'RuntimeLicKeyData', name: 'RuntimeLicKey' },
	{ lenField: 'ControlSourceData', name: 'ControlSource' },
	{ lenField: 'RowSourceData', name: 'RowSource' },
];

export interface SiteModel {
	mask: number;
	values: Map<string, number>;
	strings: Map<string, ParsedString>;
	/** fmPosition, HIMETRIC, present when mask bit 8 is set. */
	position?: { left: number; top: number };
	pads: Map<string, Buffer>;
}

export function siteName(site: SiteModel): string {
	return site.strings.get('Name')?.text ?? '';
}

export function siteId(site: SiteModel): number {
	return site.values.get('ID') ?? 0;
}

export function siteCacheIndex(site: SiteModel): number {
	return site.values.get('ClsidCacheIndex') ?? 0;
}

export function siteIsContainer(site: SiteModel): boolean {
	return CONTAINER_CACHE_INDICES.has(siteCacheIndex(site));
}

export interface FormStreamModel {
	record: ParsedRecord;
	/** FormStreamData blobs, raw, in stream order. */
	mouseIcon?: Buffer;
	fontRaw?: Buffer;
	pictureRaw?: Buffer;
	/** The class-table region including its count, preserved verbatim. */
	classTableRaw: Buffer;
	classTablePresent: boolean;
	sites: SiteModel[];
	/**
	 * The SiteDepthsAndTypes region plus its trailing pad, replayed when the
	 * site LIST is unchanged and recomputed (one entry per site, depth 0,
	 * type ST_Ole) when sites were added or removed.
	 */
	depthsRaw: Buffer;
	sitesStructurallyChanged?: boolean;
	/** DesignExtender and anything else after the sites, preserved verbatim. */
	trailingRaw: Buffer;
}

// ------------------------------------------------------------------ parse

export function parseFormStream(
	data: Buffer,
	codec: OformsTextCodec = LATIN1_CODEC,
): FormStreamModel {
	const r = new OformsReader(data);
	const record = parseRecord(r, FORM_SPEC, codec, data.length);

	const model: FormStreamModel = {
		record,
		classTableRaw: Buffer.alloc(0),
		classTablePresent: false,
		sites: [],
		depthsRaw: Buffer.alloc(0),
		trailingRaw: Buffer.alloc(0),
	};

	// FormStreamData, in mask order: MouseIcon (15), Font (20), Picture (21).
	if (record.maskLo & (1 << 15)) { model.mouseIcon = readGuidAndBlob(r); }
	if (record.maskLo & (1 << 20)) { model.fontRaw = readGuidAndFontRaw(r); }
	if (record.maskLo & (1 << 21)) { model.pictureRaw = readGuidAndBlob(r); }

	// FormSiteData. Whether the class-table count is stored depends on a flag
	// buried in BooleanProperties, which itself may be defaulted - so, as the
	// proven reader before this one did, both layouts are tried and the one
	// whose counts reconcile wins.
	const siteStart = r.pos;
	const withTable = trySiteData(data, siteStart, true, codec)
		?? trySiteData(data, siteStart, false, codec);
	if (!withTable) {
		throw new RangeError('FormSiteData did not reconcile under either class-table layout');
	}
	model.classTablePresent = withTable.classTablePresent;
	model.classTableRaw = withTable.classTableRaw;
	model.sites = withTable.sites;
	model.depthsRaw = withTable.depthsRaw;
	model.trailingRaw = withTable.trailingRaw;
	return model;
}

interface SiteDataParse {
	classTablePresent: boolean;
	classTableRaw: Buffer;
	sites: SiteModel[];
	depthsRaw: Buffer;
	trailingRaw: Buffer;
}

function trySiteData(
	data: Buffer,
	start: number,
	classTableStored: boolean,
	codec: OformsTextCodec,
): SiteDataParse | undefined {
	try {
		const r = new OformsReader(data);
		r.pos = start;
		let classTableRaw: Buffer = Buffer.alloc(0);
		if (classTableStored) {
			const tableStart = r.pos;
			const count = r.u16();
			for (let i = 0; i < count; i++) {
				const version = r.u16();
				if (version !== 0x0000) { return undefined; }
				const cbClassTable = r.u16();
				r.bytes(cbClassTable);
			}
			const end = r.pos;
			r.pos = tableStart;
			classTableRaw = r.bytes(end - tableStart);
		}
		const countOfSites = r.u32();
		const countOfBytes = r.u32();
		if (countOfSites > 10_000) { return undefined; }
		if (r.pos + countOfBytes > data.length) { return undefined; }
		const sitesEnd = r.pos + countOfBytes;

		// SiteDepthsAndTypes: per-site entries or run-length counted ones,
		// preserved raw with the pad that follows.
		const depthsStart = r.pos;
		let accounted = 0;
		while (accounted < countOfSites) {
			r.u8(); // depth
			const typeOrCount = r.u8();
			if (typeOrCount & 0x80) {
				accounted += typeOrCount & 0x7f;
				r.u8(); // OptionalType
			} else {
				accounted += 1;
			}
		}
		if (accounted !== countOfSites) { return undefined; }
		{
			const over = (r.pos - depthsStart) % 4;
			if (over !== 0) { r.bytes(4 - over); }
		}
		const depthsEnd = r.pos;
		r.pos = depthsStart;
		const depthsRaw = r.bytes(depthsEnd - depthsStart);

		const sites: SiteModel[] = [];
		for (let i = 0; i < countOfSites; i++) {
			sites.push(readSite(r, codec));
		}
		if (r.pos !== sitesEnd) { return undefined; }
		const trailingRaw: Buffer = r.bytes(data.length - r.pos);
		return { classTablePresent: classTableStored, classTableRaw, sites, depthsRaw, trailingRaw };
	} catch {
		return undefined;
	}
}

function readSite(r: OformsReader, codec: OformsTextCodec): SiteModel {
	const siteStart = r.pos;
	const version = r.u16();
	if (version !== 0x0000) {
		throw new RangeError('unsupported OleSiteConcreteControl version');
	}
	const cbSite = r.u16();
	const mask = r.u32();
	const site: SiteModel = { mask, values: new Map(), strings: new Map(), pads: new Map() };

	for (const f of SITE_DATA_FIELDS) {
		if (!(mask & (1 << f.bit))) { continue; }
		const pad = r.pad(siteStart, f.size);
		if (pad.length) { site.pads.set(`before:${f.name}`, pad); }
		const v = f.size === 2
			? (f.signed ? r.i16() : r.u16())
			: (f.signed ? r.i32() : r.u32());
		site.values.set(f.name, v);
	}
	{
		const pad = r.pad(siteStart, 4);
		if (pad.length) { site.pads.set('data:end', pad); }
	}

	const readString = (lenField: string, name: string): void => {
		const lc = site.values.get(lenField) ?? 0;
		const cb = lc & 0x7fffffff;
		const compressed = (lc & 0x80000000) !== 0;
		const raw = r.bytes(cb);
		site.strings.set(name, { text: codec.decode(raw, compressed), compressed, raw });
		const pad = r.pad(siteStart, 4);
		if (pad.length) { site.pads.set(`str:${name}`, pad); }
	};
	for (const s of SITE_STRINGS) {
		if (site.values.has(s.lenField)) { readString(s.lenField, s.name); }
	}
	if (mask & (1 << 8)) {
		site.position = { left: r.i32(), top: r.i32() };
	}
	for (const s of SITE_STRINGS_AFTER_POSITION) {
		if (site.values.has(s.lenField)) { readString(s.lenField, s.name); }
	}

	const end = siteStart + 4 + cbSite;
	if (r.pos !== end) {
		throw new RangeError(`site record has ${end - r.pos} unmodelled bytes`);
	}
	return site;
}

function readGuidAndBlob(r: OformsReader): Buffer {
	const start = r.pos;
	r.bytes(16);
	r.u32(); // preamble
	const size = r.u32();
	r.bytes(size);
	const total = r.pos - start;
	r.pos = start;
	return r.bytes(total);
}

const STDFONT_GUID_HEAD = 0x0be35203;
const TEXTPROPS_GUID_HEAD = 0xafc20920;

function readGuidAndFontRaw(r: OformsReader): Buffer {
	const start = r.pos;
	const guidHead = r.data.readUInt32LE(r.pos);
	r.bytes(16);
	if (guidHead === STDFONT_GUID_HEAD) {
		const version = r.u8();
		if (version !== 0x01) { throw new RangeError('unknown StdFont version'); }
		r.bytes(2 + 1 + 2 + 4); // charset, flags, weight, height
		const faceLen = r.u8();
		r.bytes(faceLen);
	} else if (guidHead === TEXTPROPS_GUID_HEAD) {
		r.bytes(2); // minor, major
		const cb = r.u16();
		r.bytes(cb);
	} else {
		throw new RangeError('unknown font GUID in FormStreamData');
	}
	const total = r.pos - start;
	r.pos = start;
	return r.bytes(total);
}

/** The StdFont fields of a form's font blob, when it is one. */
export function parseStdFont(fontRaw: Buffer): {
	charset: number; flags: number; weight: number; heightTenThousandthsPt: number; face: string;
} | undefined {
	if (fontRaw.length < 16 || fontRaw.readUInt32LE(0) !== STDFONT_GUID_HEAD) {
		return undefined;
	}
	const r = new OformsReader(fontRaw);
	r.bytes(16);
	r.u8(); // version
	const charset = r.i16();
	const flags = r.u8();
	const weight = r.i16();
	const height = r.u32();
	const faceLen = r.u8();
	const face = r.bytes(faceLen).toString('latin1');
	return { charset, flags, weight, heightTenThousandthsPt: height, face };
}

/** Composes a StdFont blob for a new form's Font StreamData. */
export function composeStdFont(
	face: string,
	heightTenThousandthsPt: number,
	options: { bold?: boolean; italic?: boolean; underline?: boolean; strikeout?: boolean; charset?: number } = {},
): Buffer {
	const w = new OformsWriter();
	w.u32(STDFONT_GUID_HEAD);
	// MS-DTYP 2.3.4.2: Data2 and Data3 are stored little-endian too - the
	// GUID {0BE35203-8F91-11CE-...} persists as 03 52 E3 0B 91 8F CE 11.
	// Big-endian here made Excel refuse the font as an unregistered
	// component and killed the form load (measured).
	w.bytes(Buffer.from('918fce119de300aa004bb851', 'hex'));
	w.u8(0x01);
	w.i16(options.charset ?? 0);
	// FONTFLAGS: the fBold bit "MUST be set to zero" - bold rides sWeight.
	w.u8((options.italic ? 0x02 : 0) | (options.underline ? 0x04 : 0) | (options.strikeout ? 0x08 : 0));
	w.i16(options.bold ? 700 : 400);
	w.u32(heightTenThousandthsPt >>> 0);
	const faceBytes = Buffer.from(face, 'latin1');
	w.u8(faceBytes.length);
	w.bytes(faceBytes);
	return w.toBuffer();
}

// -------------------------------------------------------------- serialize

export function serializeFormStream(model: FormStreamModel, codec: OformsTextCodec = LATIN1_CODEC): Buffer {
	const w = new OformsWriter();
	w.bytes(serializeRecord(model.record, codec));
	if (model.record.maskLo & (1 << 15)) { w.bytes(required(model.mouseIcon, 'MouseIcon')); }
	if (model.record.maskLo & (1 << 20)) { w.bytes(required(model.fontRaw, 'Font')); }
	if (model.record.maskLo & (1 << 21)) { w.bytes(required(model.pictureRaw, 'Picture')); }

	if (model.classTablePresent) { w.bytes(model.classTableRaw); }

	const sitesBytes = model.sites.map((s) => serializeSite(s, codec));
	const depths = model.sitesStructurallyChanged
		? composeDepths(model.sites.length)
		: model.depthsRaw;
	w.u32(model.sites.length);
	w.u32(depths.length + sitesBytes.reduce((n, b) => n + b.length, 0));
	w.bytes(depths);
	for (const b of sitesBytes) { w.bytes(b); }
	w.bytes(model.trailingRaw);
	return w.toBuffer();
}

/**
 * The SiteDepthsAndTypes for `count` uniform sites (depth 0, ST_Ole), in the
 * RUN-LENGTH form Excel itself always writes: one counted entry per run of
 * up to 127. The per-entry form is spec-legal too, but matching the only
 * producer fm20 is tested against costs nothing.
 */
function composeDepths(count: number): Buffer {
	const entries: number[] = [];
	let remaining = count;
	while (remaining > 0) {
		const run = Math.min(remaining, 0x7f);
		entries.push(0x00, 0x80 | run, 0x01);
		remaining -= run;
	}
	const raw = Buffer.from(entries);
	const over = raw.length % 4;
	return over === 0 ? raw : Buffer.concat([raw, Buffer.alloc(4 - over)]);
}

function serializeSite(site: SiteModel, codec: OformsTextCodec): Buffer {
	// Refresh the length fields of edited strings before the data block.
	const refresh = (lenField: string, name: string): Buffer | undefined => {
		const s = site.strings.get(name);
		if (!s) { return undefined; }
		const raw = s.edited ? codec.encode(s.text, s.compressed) : s.raw;
		site.values.set(lenField, (raw.length & 0x7fffffff) | (s.compressed ? 0x80000000 : 0));
		return raw;
	};
	const stringBytes = new Map<string, Buffer>();
	for (const s of [...SITE_STRINGS, ...SITE_STRINGS_AFTER_POSITION]) {
		const raw = refresh(s.lenField, s.name);
		if (raw !== undefined) { stringBytes.set(s.name, raw); }
	}

	const body = new OformsWriter();
	const rel = (): number => body.position + 8; // version(2) + cb(2) + mask(4)
	for (const f of SITE_DATA_FIELDS) {
		if (!(site.mask & (1 << f.bit))) { continue; }
		const over = rel() % f.size;
		if (over !== 0) {
			const captured = site.pads.get(`before:${f.name}`);
			const gap = f.size - over;
			body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
		}
		const v = site.values.get(f.name) ?? 0;
		if (f.size === 2) { f.signed ? body.i16(v) : body.u16(v); }
		else { f.signed ? body.i32(v) : body.u32(v); }
	}
	{
		const over = rel() % 4;
		if (over !== 0) {
			const captured = site.pads.get('data:end');
			const gap = 4 - over;
			body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
		}
	}
	const writeString = (name: string): void => {
		const raw = stringBytes.get(name);
		if (raw === undefined) { return; }
		body.bytes(raw);
		const over = rel() % 4;
		if (over !== 0) {
			const captured = site.pads.get(`str:${name}`);
			const gap = 4 - over;
			body.bytes(captured && captured.length === gap ? captured : Buffer.alloc(gap));
		}
	};
	for (const s of SITE_STRINGS) { writeString(s.name); }
	if (site.mask & (1 << 8)) {
		const p = site.position ?? { left: 0, top: 0 };
		body.i32(p.left); body.i32(p.top);
	}
	for (const s of SITE_STRINGS_AFTER_POSITION) { writeString(s.name); }

	const bodyBytes = body.toBuffer();
	const w = new OformsWriter();
	w.u16(0x0000);
	w.u16(bodyBytes.length + 4); // cbSite counts the mask
	w.u32(site.mask >>> 0);
	w.bytes(bodyBytes);
	return w.toBuffer();
}

function required(b: Buffer | undefined, what: string): Buffer {
	if (!b) { throw new RangeError(`masked ${what} StreamData has no bytes`); }
	return b;
}

// -------------------------------------------------------------- o stream

export type ObjectStreamEntry =
	| { kind: 'record'; site: SiteModel; record: ParsedRecord }
	| { kind: 'raw'; site: SiteModel; bytes: Buffer }
	| { kind: 'container'; site: SiteModel };

/**
 * Parses the object stream against the parsed form's site list: each
 * non-container site owns the next ObjectStreamSize bytes, holding a typed
 * record when the cache index names one this engine models and a preserved
 * raw run otherwise (third-party ActiveX).
 */
export function parseObjectStream(
	o: Buffer,
	sites: readonly SiteModel[],
	codec: OformsTextCodec = LATIN1_CODEC,
): ObjectStreamEntry[] {
	const out: ObjectStreamEntry[] = [];
	let offset = 0;
	for (const site of sites) {
		if (siteIsContainer(site)) {
			out.push({ kind: 'container', site });
			continue;
		}
		const size = site.values.get('ObjectStreamSize') ?? 0;
		const spec = RECORD_SPECS_BY_CACHE_INDEX.get(siteCacheIndex(site));
		const slice = o.subarray(offset, offset + size);
		if (!spec) {
			out.push({ kind: 'raw', site, bytes: Buffer.from(slice) });
		} else {
			const r = new OformsReader(Buffer.from(slice));
			out.push({ kind: 'record', site, record: parseRecord(r, spec, codec, slice.length) });
		}
		offset += size;
	}
	if (offset !== o.length) {
		throw new RangeError(`object stream is ${o.length} bytes but sites claim ${offset}`);
	}
	return out;
}

/**
 * Serializes the object stream and refreshes each site's ObjectStreamSize to
 * the bytes its record now occupies.
 */
export function serializeObjectStream(
	entries: readonly ObjectStreamEntry[],
	codec: OformsTextCodec = LATIN1_CODEC,
): Buffer {
	const parts: Buffer[] = [];
	for (const entry of entries) {
		if (entry.kind === 'container') { continue; }
		const bytes = entry.kind === 'raw' ? entry.bytes : serializeRecord(entry.record, codec);
		entry.site.values.set('ObjectStreamSize', bytes.length);
		parts.push(bytes);
	}
	return Buffer.concat(parts);
}
