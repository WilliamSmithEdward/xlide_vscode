// A form's designer storage, read natively.
//
// Excel stores a UserForm's control tree in a binary designer storage inside
// vbaProject.bin - a storage named after the module, whose `f` stream is the
// MS-OFORMS FormControl. The module's text carries none of it, which is why a
// workbook form's controls were unknowable without a host. This reads them out
// of the workbook itself: names, and the type each control answers to.
//
// Decoded against [MS-OFORMS] v20250819 (docs/[MS-OFORMS].pdf) and verified
// byte-for-byte against a real Excel-authored form before being written: every
// offset below was first walked by hand in a hexdump of that form's `f` stream.
//
// Conservative by construction: any structural inconsistency - a version this
// was not written for, a size that walks out of bounds - returns undefined
// rather than a guessed control list. Absence of knowledge is recoverable;
// wrong knowledge paints wrong diagnostics.

/** A control the designer declares, as the analyzer wants it. */
export interface DesignerControl {
	name: string;
	type: string;
}

/** Decodes string bytes: MBCS project-code-page when compressed, else UTF-16LE. */
export type DesignerTextDecoder = (bytes: Buffer, compressed: boolean) => string;

// FormEmbeddedActiveXControlCached ([MS-OFORMS] 2.4.5): ClsidCacheIndex values
// below 0x7FFF name a control class directly. The VBE persists the specific
// morph indices (23/25/26/...) rather than generic MorphData - observed on the
// verification form - but 15 remains legal from other producers and resolves
// through DisplayStyle in the `o` stream.
const CACHED_CONTROL_CLASSES: Readonly<Record<number, string>> = {
	7: 'MSForms.Form',
	12: 'MSForms.Image',
	14: 'MSForms.Frame',
	15: 'MorphData',
	16: 'MSForms.SpinButton',
	17: 'MSForms.CommandButton',
	18: 'MSForms.TabStrip',
	21: 'MSForms.Label',
	23: 'MSForms.TextBox',
	24: 'MSForms.ListBox',
	25: 'MSForms.ComboBox',
	26: 'MSForms.CheckBox',
	27: 'MSForms.OptionButton',
	28: 'MSForms.ToggleButton',
	47: 'MSForms.ScrollBar',
	57: 'MSForms.MultiPage',
};

// fmDisplayStyle ([MS-OFORMS] 2.5.20.1): the type of a MorphDataControl.
const MORPH_DISPLAY_STYLES: Readonly<Record<number, string>> = {
	1: 'MSForms.TextBox',
	2: 'MSForms.ListBox',
	3: 'MSForms.ComboBox',
	4: 'MSForms.CheckBox',
	5: 'MSForms.OptionButton',
	6: 'MSForms.ToggleButton',
	7: 'MSForms.ComboBox',
};

/** Bounds-checked little-endian reader over one stream. */
class StreamReader {
	pos = 0;
	constructor(private readonly data: Buffer) {}
	get length(): number { return this.data.length; }
	u8(): number { this.need(1); return this.data.readUInt8(this.pos++); }
	u16(): number { this.need(2); const v = this.data.readUInt16LE(this.pos); this.pos += 2; return v; }
	u32(): number { this.need(4); const v = this.data.readUInt32LE(this.pos); this.pos += 4; return v; }
	bytes(n: number): Buffer { this.need(n); const v = this.data.subarray(this.pos, this.pos + n); this.pos += n; return v; }
	skip(n: number): void { this.need(n); this.pos += n; }
	/** Aligns to `size` relative to `base` ([MS-OFORMS] 2.1.1.2.4). */
	align(base: number, size: number): void {
		const rel = this.pos - base;
		const over = rel % size;
		if (over !== 0) { this.skip(size - over); }
	}
	private need(n: number): void {
		if (n < 0 || this.pos + n > this.data.length) {
			throw new RangeError('designer stream walked out of bounds');
		}
	}
}

interface SiteRecord {
	name: string;
	clsidCacheIndex: number;
	objectStreamSize: number;
}

/**
 * Reads the controls out of a form's designer streams: `f` (the FormControl
 * with its sites) and, only when a generic MorphData site must be typed, `o`
 * (each site's own control record). Returns undefined when the streams are not
 * what this was written for - the caller treats that as "no knowledge", never
 * as "no controls".
 */
export function parseFormDesignerStreams(
	fStream: Buffer,
	oStream: Buffer | undefined,
	decode: DesignerTextDecoder,
): DesignerControl[] | undefined {
	try {
		const sites = readFormSites(fStream);
		const out: DesignerControl[] = [];
		let oOffset = 0;
		for (const site of sites) {
			let type: string;
			if (site.clsidCacheIndex >= 0x8000) {
				// A non-cached ActiveX control (index into the ClassTable). Its
				// members are not enumerable here; the NAME is what suppresses a
				// false undeclared finding, and an unresolvable type keeps every
				// member lookup honestly empty.
				type = 'ActiveX.Control';
			} else {
				const cached = CACHED_CONTROL_CLASSES[site.clsidCacheIndex];
				if (cached === 'MorphData') {
					type = morphDisplayType(oStream, oOffset) ?? 'MSForms.Control';
				} else {
					// An index this table does not know is a structure this
					// reader does not fully understand; the safe claim is the
					// base surface every control carries.
					type = cached ?? 'MSForms.Control';
				}
			}
			out.push({ name: site.name, type });
			oOffset += site.objectStreamSize;
		}
		return out;
	} catch {
		return undefined;
	}

	function readFormSites(data: Buffer): SiteRecord[] {
		const r = new StreamReader(data);
		const minor = r.u8();
		const major = r.u8();
		if (minor !== 0x00 || major !== 0x04) {
			throw new RangeError('not a FormControl stream');
		}
		const cbForm = r.u16();
		const formPropMask = r.u32();
		// PropMask is included in cbForm; skip the rest of DataBlock+Extra.
		r.pos = 4 + cbForm;

		// FormStreamData: font and picture blobs, present only when their bits
		// are set. Skipped structurally so SiteData is found wherever it is.
		if (formPropMask & (1 << 15)) { skipGuidAndPicture(r); } // MouseIcon
		if (formPropMask & (1 << 20)) { skipGuidAndFont(r); }    // Font
		if (formPropMask & (1 << 21)) { skipGuidAndPicture(r); } // Picture

		// FormSiteData. The class-table count is stored unless the form's
		// BooleanProperties set FORM_FLAG_DONTSAVECLASSTABLE (bit 15); the flags
		// live in the DataBlock this reader deliberately does not walk, so both
		// layouts are tried and the one whose sizes prove out wins. The counts
		// make the wrong reading collapse immediately: CountOfBytes must equal
		// exactly what remains after the arrays it describes.
		const withTable = trySiteData(data, r.pos, true) ?? trySiteData(data, r.pos, false);
		if (!withTable) {
			throw new RangeError('FormSiteData did not reconcile');
		}
		return withTable;
	}

	function trySiteData(data: Buffer, start: number, classTableStored: boolean): SiteRecord[] | undefined {
		try {
			const r = new StreamReader(data);
			r.pos = start;
			if (classTableStored) {
				const count = r.u16();
				for (let i = 0; i < count; i++) {
					const version = r.u16();
					if (version !== 0x0000) { return undefined; }
					const cbClassTable = r.u16();
					r.skip(cbClassTable);
				}
			}
			const countOfSites = r.u32();
			const countOfBytes = r.u32();
			if (countOfSites > 10_000) { return undefined; }
			if (r.pos + countOfBytes !== data.length) {
				// SiteData is the last thing in the stream, so an exact match is
				// the cheap proof this layout is the real one.
				return undefined;
			}
			// SiteDepthsAndTypes: one entry per site, or one counted entry for a
			// run of consecutive sites with the same depth and type.
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
			r.align(depthsStart, 4);

			const sites: SiteRecord[] = [];
			for (let i = 0; i < countOfSites; i++) {
				sites.push(readSite(r));
			}
			return sites;
		} catch {
			return undefined;
		}
	}

	function readSite(r: StreamReader): SiteRecord {
		const siteStart = r.pos;
		const version = r.u16();
		if (version !== 0x0000) {
			throw new RangeError('unsupported OleSiteConcreteControl version');
		}
		const cbSite = r.u16();
		const mask = r.u32();
		// SiteDataBlock, fields in mask-bit order, aligned to their own size
		// relative to the site's start ([MS-OFORMS] 2.2.10.12.3).
		let nameCb = 0;
		let nameCompressed = false;
		let clsidCacheIndex = 0;
		let objectStreamSize = 0;
		const read4 = (): number => { r.align(siteStart, 4); return r.u32(); };
		const read2 = (): number => { r.align(siteStart, 2); return r.u16(); };
		if (mask & (1 << 0)) { const v = read4(); nameCb = v & 0x7fffffff; nameCompressed = (v & 0x80000000) !== 0; }
		if (mask & (1 << 1)) { read4(); } // TagData
		if (mask & (1 << 2)) { read4(); } // ID
		if (mask & (1 << 3)) { read4(); } // HelpContextID
		if (mask & (1 << 4)) { read4(); } // BitFlags
		if (mask & (1 << 5)) { objectStreamSize = read4(); }
		if (mask & (1 << 6)) { read2(); } // TabIndex
		if (mask & (1 << 7)) { clsidCacheIndex = read2(); }
		if (mask & (1 << 9)) { read2(); } // GroupID
		if (mask & (1 << 11)) { read4(); } // ControlTipTextData
		if (mask & (1 << 12)) { read4(); } // RuntimeLicKeyData
		if (mask & (1 << 13)) { read4(); } // ControlSourceData
		if (mask & (1 << 14)) { read4(); } // RowSourceData
		r.align(siteStart, 4); // DataBlock ends on a 4-byte boundary.

		// SiteExtraDataBlock opens with the Name string, padded to 4 bytes.
		let name = '';
		if (nameCb > 0) {
			name = decode(r.bytes(nameCb), nameCompressed);
			r.align(siteStart, 4);
		}
		// The rest of the extra block (Tag, Position, tooltips) is jumped, not
		// walked: cbSite states where the next site begins.
		r.pos = siteStart + 4 + cbSite;
		return { name, clsidCacheIndex, objectStreamSize };
	}

	function morphDisplayType(o: Buffer | undefined, offset: number): string | undefined {
		if (!o) {
			return undefined;
		}
		try {
			const r = new StreamReader(o);
			r.pos = offset;
			const minor = r.u8();
			const major = r.u8();
			if (minor !== 0x00 || major !== 0x02) {
				return undefined;
			}
			r.u16(); // cbMorphData
			const maskLo = r.u32();
			r.u32(); // MorphDataPropMask is 8 bytes; DisplayStyle sits in the low word.
			// DataBlock in bit order: A VariousPropertyBits(4), B BackColor(4),
			// C ForeColor(4), D MaxLength(4) - all 4-byte and 4-aligned here -
			// then E BorderStyle(1), F ScrollBars(1), G DisplayStyle(1).
			for (const bit of [0, 1, 2, 3]) {
				if (maskLo & (1 << bit)) { r.u32(); }
			}
			if (maskLo & (1 << 4)) { r.u8(); }
			if (maskLo & (1 << 5)) { r.u8(); }
			if (!(maskLo & (1 << 6))) {
				// DisplayStyle not stored: the file format default is Text.
				return MORPH_DISPLAY_STYLES[1];
			}
			return MORPH_DISPLAY_STYLES[r.u8()];
		} catch {
			return undefined;
		}
	}

	function skipGuidAndFont(r: StreamReader): void {
		const guid = r.bytes(16);
		// StdFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} vs TextProps
		// {AFC20920-DA4E-11CE-B943-00AA006887B4} ([MS-OFORMS] 2.4.6).
		if (guid.readUInt32LE(0) === 0x0be35203) {
			const version = r.u8();
			if (version !== 0x01) { throw new RangeError('unknown StdFont version'); }
			r.skip(2 + 1 + 2 + 4); // charset, flags, weight, height
			const faceLen = r.u8();
			r.skip(faceLen);
			return;
		}
		if (guid.readUInt32LE(0) === 0xafc20920) {
			r.skip(2); // minor, major
			const cbTextProps = r.u16();
			r.skip(cbTextProps);
			return;
		}
		throw new RangeError('unknown font GUID in FormStreamData');
	}

	function skipGuidAndPicture(r: StreamReader): void {
		r.bytes(16); // GUID
		r.u32(); // Preamble
		const size = r.u32();
		r.skip(size);
	}
}
