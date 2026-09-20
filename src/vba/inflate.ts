// DEFLATE decompression (RFC 1951) in pure TypeScript, for builds that have
// no zlib: the web extension host runs in a browser, where `node:zlib` does
// not exist and the only built-in inflate is DecompressionStream, which is
// asynchronous. The whole container engine is synchronous and reads a part at
// a time out of an already-loaded archive, so an async inflate would mean
// threading promises through every reader in src/vba/. This is ~200 lines
// instead, and zlib itself is the oracle: tests/inflate.test.ts asserts byte
// identity against inflateRawSync over every fixture container in the repo.
//
// Only decompression lives here. Nothing needs a pure-TS deflate, because
// ZipArchive carries untouched entries over with their original compressed
// bytes and can store (method 0) the few entries an edit rewrites.

/** Length codes 257..285: base length and extra bits. */
const LENGTH_BASE = [
	3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
	35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258,
];
const LENGTH_EXTRA = [
	0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
	3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];

/** Distance codes 0..29: base distance and extra bits. */
const DIST_BASE = [
	1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
	257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA = [
	0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
	7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];

/** The order code lengths for the code-length alphabet arrive in. */
const CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

export class InflateError extends Error {
	/**
	 * Set when the stream ended mid-symbol rather than being malformed.
	 * PowerPoint writes compressed storages with no final block, so that
	 * caller treats a truncated end as success.
	 */
	readonly truncated: boolean;

	constructor(message: string, truncated = false) {
		super(message);
		this.truncated = truncated;
	}
}

export interface InflateOptions {
	/** Presizes the output. The ZIP central directory always knows it. */
	expectedSize?: number;
	/**
	 * Return the bytes decoded so far instead of throwing when the input runs
	 * out before a final block. This is what zlib's Z_SYNC_FLUSH finish does,
	 * and PowerPoint's compressed storages need it.
	 */
	allowTruncated?: boolean;
}

/**
 * A canonical Huffman decoding table.
 *
 * `table` is indexed by the next `maxBits` bits of input (LSB first, which is
 * how DEFLATE stores codes) and holds `symbol << 4 | codeLength`. Every index
 * whose low `codeLength` bits match a code resolves to that code, so one
 * array read decodes a symbol. A 15-bit table is 32768 entries, and filling
 * it costs one pass over 2^maxBits regardless of the alphabet size.
 */
interface Huffman {
	table: Uint32Array;
	maxBits: number;
}

function buildHuffman(lengths: ArrayLike<number>, count: number): Huffman {
	let maxBits = 0;
	for (let i = 0; i < count; i++) {
		if (lengths[i] > maxBits) {
			maxBits = lengths[i];
		}
	}
	if (maxBits === 0) {
		// An empty alphabet is legal (a block with no distance codes). Any
		// lookup against it is a malformed stream, and table[x] === 0 gives a
		// zero code length, which decode() rejects.
		return { table: new Uint32Array(1), maxBits: 1 };
	}
	if (maxBits > 15) {
		throw new InflateError(`Huffman code length ${maxBits} exceeds 15.`);
	}

	// Canonical code assignment: codes of each length are consecutive, and
	// each length starts where the previous length left off, shifted up.
	const blCount = new Uint16Array(maxBits + 1);
	for (let i = 0; i < count; i++) {
		blCount[lengths[i]]++;
	}
	blCount[0] = 0;
	const nextCode = new Uint32Array(maxBits + 2);
	let code = 0;
	for (let bits = 1; bits <= maxBits; bits++) {
		code = (code + blCount[bits - 1]) << 1;
		nextCode[bits] = code;
	}

	const size = 1 << maxBits;
	const table = new Uint32Array(size);
	for (let symbol = 0; symbol < count; symbol++) {
		const len = lengths[symbol];
		if (len === 0) {
			continue;
		}
		// DEFLATE packs Huffman codes most-significant bit first, while the
		// bit reader hands back least-significant bits first, so the code is
		// reversed before it indexes the table.
		let canonical = nextCode[len]++;
		let reversed = 0;
		for (let b = 0; b < len; b++) {
			reversed = (reversed << 1) | (canonical & 1);
			canonical >>>= 1;
		}
		const entry = (symbol << 4) | len;
		for (let i = reversed; i < size; i += 1 << len) {
			table[i] = entry;
		}
	}
	return { table, maxBits };
}

class BitReader {
	private buf = 0;
	private cnt = 0;
	pos = 0;

	constructor(private readonly src: Uint8Array) {}

	/**
	 * Loads at least `need` bits, or as many as remain. Running out is not an
	 * error here: a Huffman code shorter than maxBits can still decode from a
	 * short buffer, and the callers that truly need the bits check `cnt`.
	 */
	private fill(need: number): void {
		while (this.cnt < need && this.pos < this.src.length) {
			this.buf |= this.src[this.pos++] << this.cnt;
			this.cnt += 8;
		}
	}

	bits(need: number): number {
		if (need === 0) {
			return 0;
		}
		this.fill(need);
		if (this.cnt < need) {
			throw new InflateError('Truncated deflate stream.', true);
		}
		const value = this.buf & ((1 << need) - 1);
		this.buf >>>= need;
		this.cnt -= need;
		return value;
	}

	decode(huffman: Huffman): number {
		this.fill(huffman.maxBits);
		const entry = huffman.table[this.buf & ((1 << huffman.maxBits) - 1)];
		const len = entry & 15;
		if (len === 0) {
			throw new InflateError('Invalid Huffman code in deflate stream.');
		}
		if (len > this.cnt) {
			throw new InflateError('Truncated deflate stream.', true);
		}
		this.buf >>>= len;
		this.cnt -= len;
		return entry >>> 4;
	}

	/**
	 * The source offset of the next unconsumed bit, rounded up to a byte
	 * boundary. Stored blocks are byte-aligned and are read straight from the
	 * source, so the reader has to say where it actually got to: `pos` runs
	 * ahead by whatever is still sitting in the bit buffer.
	 */
	byteAlignedPosition(): number {
		return this.pos - (this.cnt >>> 3);
	}

	/** Restarts reading at a byte offset, discarding the bit buffer. */
	seekToByte(at: number): void {
		this.pos = at;
		this.buf = 0;
		this.cnt = 0;
	}
}

class Output {
	private buf: Uint8Array;
	length = 0;

	constructor(initial: number) {
		this.buf = new Uint8Array(Math.max(initial, 64));
	}

	private grow(extra: number): void {
		if (this.length + extra <= this.buf.length) {
			return;
		}
		let size = this.buf.length * 2;
		while (size < this.length + extra) {
			size *= 2;
		}
		const next = new Uint8Array(size);
		next.set(this.buf.subarray(0, this.length));
		this.buf = next;
	}

	byte(value: number): void {
		this.grow(1);
		this.buf[this.length++] = value;
	}

	/** Copies `len` bytes from `dist` back, byte at a time: the ranges overlap
	 * by design (that is how DEFLATE encodes runs), so set() cannot be used. */
	copy(dist: number, len: number): void {
		if (dist > this.length) {
			throw new InflateError('Deflate back-reference precedes the output.');
		}
		this.grow(len);
		let from = this.length - dist;
		for (let i = 0; i < len; i++) {
			this.buf[this.length++] = this.buf[from++];
		}
	}

	append(src: Uint8Array, start: number, len: number): void {
		this.grow(len);
		this.buf.set(src.subarray(start, start + len), this.length);
		this.length += len;
	}

	finish(): Buffer {
		return Buffer.from(this.buf.buffer, this.buf.byteOffset, this.length);
	}
}

let fixedLiteral: Huffman | undefined;
let fixedDistance: Huffman | undefined;

function fixedTables(): { literal: Huffman; distance: Huffman } {
	if (!fixedLiteral || !fixedDistance) {
		const litLengths = new Uint8Array(288);
		litLengths.fill(8, 0, 144);
		litLengths.fill(9, 144, 256);
		litLengths.fill(7, 256, 280);
		litLengths.fill(8, 280, 288);
		fixedLiteral = buildHuffman(litLengths, 288);
		const distLengths = new Uint8Array(30);
		distLengths.fill(5);
		fixedDistance = buildHuffman(distLengths, 30);
	}
	return { literal: fixedLiteral, distance: fixedDistance };
}

function readDynamicTables(reader: BitReader): { literal: Huffman; distance: Huffman } {
	const hlit = reader.bits(5) + 257;
	const hdist = reader.bits(5) + 1;
	const hclen = reader.bits(4) + 4;

	const clenLengths = new Uint8Array(19);
	for (let i = 0; i < hclen; i++) {
		clenLengths[CLEN_ORDER[i]] = reader.bits(3);
	}
	const clenTable = buildHuffman(clenLengths, 19);

	// The literal/length and distance code lengths share one run-length
	// encoded sequence, so they are read together and split afterwards.
	const lengths = new Uint8Array(hlit + hdist);
	let i = 0;
	while (i < lengths.length) {
		const symbol = reader.decode(clenTable);
		if (symbol < 16) {
			lengths[i++] = symbol;
		} else if (symbol === 16) {
			if (i === 0) {
				throw new InflateError('Deflate code-length repeat with nothing to repeat.');
			}
			const previous = lengths[i - 1];
			let repeat = 3 + reader.bits(2);
			while (repeat-- > 0 && i < lengths.length) {
				lengths[i++] = previous;
			}
		} else if (symbol === 17) {
			let repeat = 3 + reader.bits(3);
			while (repeat-- > 0 && i < lengths.length) {
				lengths[i++] = 0;
			}
		} else {
			let repeat = 11 + reader.bits(7);
			while (repeat-- > 0 && i < lengths.length) {
				lengths[i++] = 0;
			}
		}
	}

	return {
		literal: buildHuffman(lengths.subarray(0, hlit), hlit),
		distance: buildHuffman(lengths.subarray(hlit), hdist),
	};
}

/**
 * Inflates a raw DEFLATE stream (no zlib or gzip wrapper), the form ZIP
 * entries with compression method 8 are stored in.
 *
 * `expectedSize` presizes the output when the caller knows it, which the ZIP
 * reader always does from the central directory.
 */
export function inflateRaw(src: Uint8Array, options: InflateOptions = {}): Buffer {
	const out = new Output(options.expectedSize ?? src.length * 4);
	try {
		inflateInto(src, out);
	} catch (err) {
		// A stream that ends mid-symbol is only an error when the caller
		// expected a final block; a malformed one always is.
		if (options.allowTruncated && err instanceof InflateError && err.truncated) {
			return out.finish();
		}
		throw err;
	}
	return out.finish();
}

function inflateInto(src: Uint8Array, out: Output): void {
	const reader = new BitReader(src);

	for (;;) {
		const final = reader.bits(1);
		const type = reader.bits(2);

		if (type === 0) {
			const start = reader.byteAlignedPosition();
			if (start + 4 > src.length) {
				throw new InflateError('Truncated stored block header.', true);
			}
			const len = src[start] | (src[start + 1] << 8);
			const nlen = src[start + 2] | (src[start + 3] << 8);
			if ((len ^ 0xffff) !== nlen) {
				throw new InflateError('Stored block length check failed.');
			}
			const from = start + 4;
			if (from + len > src.length) {
				throw new InflateError('Truncated stored block.', true);
			}
			out.append(src, from, len);
			reader.seekToByte(from + len);
		} else if (type === 3) {
			throw new InflateError('Reserved deflate block type.');
		} else {
			const { literal, distance } = type === 1 ? fixedTables() : readDynamicTables(reader);

			for (;;) {
				const symbol = reader.decode(literal);
				if (symbol < 256) {
					out.byte(symbol);
					continue;
				}
				if (symbol === 256) {
					break;
				}
				const lengthIndex = symbol - 257;
				if (lengthIndex >= LENGTH_BASE.length) {
					throw new InflateError(`Invalid length symbol ${symbol}.`);
				}
				const len = LENGTH_BASE[lengthIndex] + reader.bits(LENGTH_EXTRA[lengthIndex]);
				const distSymbol = reader.decode(distance);
				if (distSymbol >= DIST_BASE.length) {
					throw new InflateError(`Invalid distance symbol ${distSymbol}.`);
				}
				const dist = DIST_BASE[distSymbol] + reader.bits(DIST_EXTRA[distSymbol]);
				out.copy(dist, len);
			}
		}

		if (final === 1) {
			return;
		}
	}
}

/**
 * Inflates a zlib-wrapped stream (RFC 1950): a two-byte header, the raw
 * DEFLATE data, and an Adler-32 trailer. PowerPoint's compressed CFB storages
 * use this form rather than the raw one ZIP entries use.
 */
export function inflate(src: Uint8Array, options: InflateOptions = {}): Buffer {
	if (src.length < 2) {
		throw new InflateError('Truncated zlib stream.', true);
	}
	const cmf = src[0];
	const flg = src[1];
	if ((cmf & 0x0f) !== 8) {
		throw new InflateError(`Unsupported zlib compression method ${cmf & 0x0f}.`);
	}
	if (((cmf << 8) | flg) % 31 !== 0) {
		throw new InflateError('Invalid zlib header check.');
	}
	if (flg & 0x20) {
		throw new InflateError('Preset dictionaries are not supported.');
	}
	return inflateRaw(src.subarray(2), options);
}
