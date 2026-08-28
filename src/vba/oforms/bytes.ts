// The byte substrate for [MS-OFORMS] records.
//
// Two facts of the format shape everything here. Fields inside a DataBlock
// are aligned to their own size relative to the record start, and the bytes
// a reader skips to reach that alignment are UNDEFINED - Excel writes
// whatever was in its buffer ([MS-OFORMS] 2.1.1.2.4). So a reader that wants
// to write the same bytes back must CAPTURE padding, not merely skip it, and
// a writer replays captured padding verbatim while emitting zeros where no
// capture exists (new content, or content whose shape an edit changed).

/** Little-endian reader that captures alignment padding instead of skipping it. */
export class OformsReader {
	pos = 0;
	constructor(readonly data: Buffer) {}
	get remaining(): number { return this.data.length - this.pos; }
	u8(): number { this.need(1); return this.data.readUInt8(this.pos++); }
	u16(): number { this.need(2); const v = this.data.readUInt16LE(this.pos); this.pos += 2; return v; }
	u32(): number { this.need(4); const v = this.data.readUInt32LE(this.pos); this.pos += 4; return v; }
	i16(): number { this.need(2); const v = this.data.readInt16LE(this.pos); this.pos += 2; return v; }
	i32(): number { this.need(4); const v = this.data.readInt32LE(this.pos); this.pos += 4; return v; }
	bytes(n: number): Buffer { this.need(n); const v = Buffer.from(this.data.subarray(this.pos, this.pos + n)); this.pos += n; return v; }
	/** Padding to align `size` relative to `base`; the captured bytes come back. */
	pad(base: number, size: number): Buffer {
		const over = (this.pos - base) % size;
		return over === 0 ? EMPTY : this.bytes(size - over);
	}
	private need(n: number): void {
		if (n < 0 || this.pos + n > this.data.length) {
			throw new RangeError(`oforms read out of bounds at ${this.pos}+${n}/${this.data.length}`);
		}
	}
}

const EMPTY = Buffer.alloc(0);

/** Little-endian writer whose alignment replays captured padding when given. */
export class OformsWriter {
	private chunks: Buffer[] = [];
	private length = 0;
	u8(v: number): void { const b = Buffer.alloc(1); b.writeUInt8(v & 0xff); this.push(b); }
	u16(v: number): void { const b = Buffer.alloc(2); b.writeUInt16LE(v & 0xffff); this.push(b); }
	u32(v: number): void { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); this.push(b); }
	i16(v: number): void { const b = Buffer.alloc(2); b.writeInt16LE(v | 0); this.push(b); }
	i32(v: number): void { const b = Buffer.alloc(4); b.writeInt32LE(v | 0); this.push(b); }
	bytes(v: Buffer): void { this.push(v); }
	/**
	 * Aligns `size` relative to `base`, replaying `captured` when its length
	 * matches the gap (the unmodified round-trip) and zeros otherwise.
	 */
	pad(base: number, size: number, captured?: Buffer): void {
		const over = (this.length - base) % size;
		if (over === 0) { return; }
		const gap = size - over;
		this.push(captured && captured.length === gap ? captured : Buffer.alloc(gap));
	}
	get position(): number { return this.length; }
	toBuffer(): Buffer { return Buffer.concat(this.chunks, this.length); }
	private push(b: Buffer): void { this.chunks.push(b); this.length += b.length; }
}

/** HIMETRIC (0.01 mm) per point: 2540 per inch over 72 points per inch. */
export const HIMETRIC_PER_POINT = 2540 / 72;

export function himetricToPoints(h: number): number {
	return h / HIMETRIC_PER_POINT;
}

export function pointsToHimetric(pt: number): number {
	return Math.round(pt * HIMETRIC_PER_POINT);
}
