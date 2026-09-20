// A Buffer for the web extension host, which has no Node globals.
//
// The container engine is built on Buffer throughout - 418 static calls and
// twenty numeric accessors across 170 files - and rewriting all of it to
// Uint8Array would touch every byte path in a codebase whose oracle is byte
// identity. The browser build injects this class as `Buffer` instead (see
// webBuild.js), so the engine runs unmodified.
//
// Node's own Buffer is the oracle: tests/webBuffer.test.ts asserts identical
// results for every method and encoding here, over generated inputs. Two
// details are easy to get wrong and are deliberate below:
//
//   - slice() shares memory. Node's Buffer.prototype.slice aliases subarray,
//     unlike Uint8Array.prototype.slice, which copies. Inheriting the copying
//     one would silently break every in-place edit the engine makes.
//   - subarray() must return a Buffer, which it does for free: TypedArray's
//     species construction calls this class, so no override is needed - and
//     no constructor may be declared that changes the (buffer, offset,
//     length) signature species relies on.

type Encoding = 'utf8' | 'latin1' | 'ascii' | 'hex' | 'base64' | 'base64url' | 'utf16le';

function normalizeEncoding(encoding?: string): Encoding {
	switch ((encoding ?? 'utf8').toLowerCase()) {
		case 'utf8':
		case 'utf-8':
			return 'utf8';
		case 'latin1':
		case 'binary':
			return 'latin1';
		case 'ascii':
			return 'ascii';
		case 'hex':
			return 'hex';
		case 'base64':
			return 'base64';
		case 'base64url':
			return 'base64url';
		case 'utf16le':
		case 'utf-16le':
		case 'ucs2':
		case 'ucs-2':
			return 'utf16le';
		default:
			throw new TypeError(`Unknown encoding: ${encoding}`);
	}
}

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const BASE64_LOOKUP = (() => {
	const table = new Int16Array(256).fill(-1);
	for (let i = 0; i < BASE64_ALPHABET.length; i++) {
		table[BASE64_ALPHABET.charCodeAt(i)] = i;
	}
	// Node decodes the URL-safe alphabet through the same call, so '-' and
	// '_' are digits rather than characters to skip. A VBA name like
	// `VB_Name` read as base64 turns on this.
	table['-'.charCodeAt(0)] = 62;
	table['_'.charCodeAt(0)] = 63;
	return table;
})();

// TextEncoder/TextDecoder are globals in Node and in the browser, and they
// carry UTF-8's hard parts: multi-byte sequences, surrogate pairs, and the
// U+FFFD substitution Node's own toString('utf8') performs.
const utf8Encoder = new TextEncoder();
const utf8Decoder = new TextDecoder('utf-8');

function encodeString(text: string, encoding: Encoding): Uint8Array {
	switch (encoding) {
		case 'utf8':
			return utf8Encoder.encode(text);
		case 'latin1':
		case 'ascii': {
			// Encoding is the same for both: Node truncates each code unit to
			// its low byte and does not mask the high bit here. Only
			// toString('ascii') masks to 0x7f.
			const out = new Uint8Array(text.length);
			for (let i = 0; i < text.length; i++) {
				out[i] = text.charCodeAt(i) & 0xff;
			}
			return out;
		}
		case 'hex': {
			// Node stops at the first character pair that is not hex, and
			// reads each character as its low byte (see decodeBase64).
			const usable = text.length - (text.length % 2);
			const out = new Uint8Array(usable / 2);
			let written = 0;
			for (let i = 0; i < usable; i += 2) {
				const high = hexDigit(text.charCodeAt(i) & 0xff);
				const low = hexDigit(text.charCodeAt(i + 1) & 0xff);
				if (high < 0 || low < 0) {
					break;
				}
				out[written++] = (high << 4) | low;
			}
			return out.subarray(0, written);
		}
		case 'base64':
		case 'base64url':
			// One decoder for both: the lookup already maps '-' and '_'
			// alongside '+' and '/', exactly as Node's does.
			return decodeBase64(text);
		case 'utf16le': {
			// charCodeAt yields UTF-16 code units, which is exactly what this
			// encoding stores; a lone surrogate round-trips unchanged.
			const out = new Uint8Array(text.length * 2);
			for (let i = 0; i < text.length; i++) {
				const unit = text.charCodeAt(i);
				out[i * 2] = unit & 0xff;
				out[i * 2 + 1] = unit >>> 8;
			}
			return out;
		}
	}
}

/** -1 when the byte is not a hex digit. */
function hexDigit(code: number): number {
	if (code >= 0x30 && code <= 0x39) {
		return code - 0x30;
	}
	if (code >= 0x61 && code <= 0x66) {
		return code - 0x61 + 10;
	}
	if (code >= 0x41 && code <= 0x46) {
		return code - 0x41 + 10;
	}
	return -1;
}

function decodeBase64(text: string): Uint8Array {
	// Node reads the string as Latin-1 first, so each UTF-16 code unit is
	// truncated to its low byte before the alphabet is consulted. That is why
	// a Cyrillic string decodes to real bytes rather than to nothing: 'д'
	// (U+0434) becomes 0x34, the character '4'. Anything outside the alphabet
	// is skipped, and '=' ends the data.
	const out = new Uint8Array((text.length * 3) >>> 2);
	let written = 0;
	let acc = 0;
	let bits = 0;
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i) & 0xff;
		if (code === 0x3d) {
			break;
		}
		const digit = BASE64_LOOKUP[code];
		if (digit < 0) {
			continue;
		}
		acc = (acc << 6) | digit;
		bits += 6;
		if (bits >= 8) {
			bits -= 8;
			out[written++] = (acc >>> bits) & 0xff;
		}
	}
	return out.subarray(0, written);
}

function encodeBase64(bytes: Uint8Array): string {
	let out = '';
	let i = 0;
	for (; i + 2 < bytes.length; i += 3) {
		const triple = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
		out +=
			BASE64_ALPHABET[(triple >>> 18) & 63] +
			BASE64_ALPHABET[(triple >>> 12) & 63] +
			BASE64_ALPHABET[(triple >>> 6) & 63] +
			BASE64_ALPHABET[triple & 63];
	}
	const left = bytes.length - i;
	if (left === 1) {
		const value = bytes[i] << 16;
		out += BASE64_ALPHABET[(value >>> 18) & 63] + BASE64_ALPHABET[(value >>> 12) & 63] + '==';
	} else if (left === 2) {
		const value = (bytes[i] << 16) | (bytes[i + 1] << 8);
		out +=
			BASE64_ALPHABET[(value >>> 18) & 63] +
			BASE64_ALPHABET[(value >>> 12) & 63] +
			BASE64_ALPHABET[(value >>> 6) & 63] +
			'=';
	}
	return out;
}

function decodeBytes(bytes: Uint8Array, encoding: Encoding): string {
	switch (encoding) {
		case 'utf8':
			return utf8Decoder.decode(bytes);
		case 'latin1': {
			let out = '';
			for (let i = 0; i < bytes.length; i++) {
				out += String.fromCharCode(bytes[i]);
			}
			return out;
		}
		case 'ascii': {
			let out = '';
			for (let i = 0; i < bytes.length; i++) {
				out += String.fromCharCode(bytes[i] & 0x7f);
			}
			return out;
		}
		case 'hex': {
			let out = '';
			for (let i = 0; i < bytes.length; i++) {
				out += bytes[i].toString(16).padStart(2, '0');
			}
			return out;
		}
		case 'base64':
			return encodeBase64(bytes);
		case 'base64url':
			// The URL-safe alphabet, and no padding: '=' is what makes a
			// base64 string unusable inside a URI path, which is the whole
			// reason this encoding exists.
			return encodeBase64(bytes).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
		case 'utf16le': {
			// Deliberately not TextDecoder: Node maps code units straight
			// through, so a lone surrogate survives a round trip, while
			// TextDecoder would replace it with U+FFFD. VBA project streams
			// are full of UTF-16 that must come back byte-identical. An odd
			// trailing byte is dropped, as Node does.
			let out = '';
			for (let i = 0; i + 1 < bytes.length; i += 2) {
				out += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
			}
			return out;
		}
	}
}

/**
 * Pulls a byte count back to a character boundary. Node's write() drops a
 * character that will not fit rather than storing half of one, so writing
 * three-byte characters into ten bytes writes nine.
 */
function wholeCharacters(bytes: Uint8Array, count: number, encoding: Encoding): number {
	if (count >= bytes.length) {
		return count;
	}
	if (encoding === 'utf16le') {
		return count - (count % 2);
	}
	if (encoding !== 'utf8') {
		return count;
	}
	// A continuation byte at the cut means the sequence that owns it starts
	// earlier; retreat to that lead byte.
	let end = count;
	while (end > 0 && (bytes[end] & 0xc0) === 0x80) {
		end--;
	}
	return end;
}

/** Clamps a start/end pair to the buffer the way Node's readers do. */
function clamp(value: number | undefined, fallback: number, limit: number): number {
	let n = value === undefined ? fallback : Math.trunc(value);
	if (Number.isNaN(n)) {
		n = fallback;
	}
	if (n < 0) {
		n += limit;
		if (n < 0) {
			n = 0;
		}
	}
	return Math.min(Math.max(n, 0), limit);
}

class WebBufferClass extends Uint8Array {
	// No constructor is declared on purpose: TypedArray species construction
	// calls `new WebBufferClass(buffer, byteOffset, length)` for subarray(),
	// and a custom signature would break it.

	// ---- reading numbers ----

	readUInt8(offset = 0): number {
		return this[offset];
	}

	readInt8(offset = 0): number {
		const value = this[offset];
		return value & 0x80 ? value - 0x100 : value;
	}

	readUInt16LE(offset = 0): number {
		return this[offset] | (this[offset + 1] << 8);
	}

	readUInt16BE(offset = 0): number {
		return (this[offset] << 8) | this[offset + 1];
	}

	readInt16LE(offset = 0): number {
		const value = this.readUInt16LE(offset);
		return value & 0x8000 ? value - 0x10000 : value;
	}

	readInt16BE(offset = 0): number {
		const value = this.readUInt16BE(offset);
		return value & 0x8000 ? value - 0x10000 : value;
	}

	readUInt32LE(offset = 0): number {
		return (
			(this[offset] |
				(this[offset + 1] << 8) |
				(this[offset + 2] << 16)) +
			this[offset + 3] * 0x1000000
		);
	}

	readUInt32BE(offset = 0): number {
		return (
			this[offset] * 0x1000000 +
			((this[offset + 1] << 16) | (this[offset + 2] << 8) | this[offset + 3])
		);
	}

	readInt32LE(offset = 0): number {
		return (
			this[offset] |
			(this[offset + 1] << 8) |
			(this[offset + 2] << 16) |
			(this[offset + 3] << 24)
		);
	}

	readInt32BE(offset = 0): number {
		return (
			(this[offset] << 24) |
			(this[offset + 1] << 16) |
			(this[offset + 2] << 8) |
			this[offset + 3]
		);
	}

	readUIntLE(offset: number, byteLength: number): number {
		let value = 0;
		let scale = 1;
		for (let i = 0; i < byteLength; i++) {
			value += this[offset + i] * scale;
			scale *= 0x100;
		}
		return value;
	}

	readUIntBE(offset: number, byteLength: number): number {
		let value = 0;
		for (let i = 0; i < byteLength; i++) {
			value = value * 0x100 + this[offset + i];
		}
		return value;
	}

	/**
	 * A DataView over this buffer's own window, for the IEEE-754 and 64-bit
	 * accessors. Built per call rather than cached: subarray() hands out
	 * views sharing one ArrayBuffer, so a cached view would outlive its
	 * offsets.
	 */
	private view(): DataView {
		return new DataView(this.buffer, this.byteOffset, this.byteLength);
	}

	readFloatLE(offset = 0): number {
		return this.view().getFloat32(offset, true);
	}

	readFloatBE(offset = 0): number {
		return this.view().getFloat32(offset, false);
	}

	readDoubleLE(offset = 0): number {
		return this.view().getFloat64(offset, true);
	}

	readDoubleBE(offset = 0): number {
		return this.view().getFloat64(offset, false);
	}

	/** Signed, variable width: the unsigned value sign-extended from its top bit. */
	readIntLE(offset: number, byteLength: number): number {
		const value = this.readUIntLE(offset, byteLength);
		const half = Math.pow(2, 8 * byteLength - 1);
		return value >= half ? value - half * 2 : value;
	}

	readIntBE(offset: number, byteLength: number): number {
		const value = this.readUIntBE(offset, byteLength);
		const half = Math.pow(2, 8 * byteLength - 1);
		return value >= half ? value - half * 2 : value;
	}

	readBigUInt64BE(offset = 0): bigint {
		return this.view().getBigUint64(offset, false);
	}

	readBigInt64BE(offset = 0): bigint {
		return this.view().getBigInt64(offset, false);
	}

	readBigUInt64LE(offset = 0): bigint {
		return (
			BigInt(this.readUInt32LE(offset)) + (BigInt(this.readUInt32LE(offset + 4)) << 32n)
		);
	}

	readBigInt64LE(offset = 0): bigint {
		return (
			BigInt(this.readUInt32LE(offset)) + (BigInt(this.readInt32LE(offset + 4)) << 32n)
		);
	}

	// ---- writing numbers ----

	writeUInt8(value: number, offset = 0): number {
		this[offset] = value & 0xff;
		return offset + 1;
	}

	writeInt8(value: number, offset = 0): number {
		this[offset] = value & 0xff;
		return offset + 1;
	}

	writeUInt16LE(value: number, offset = 0): number {
		this[offset] = value & 0xff;
		this[offset + 1] = (value >>> 8) & 0xff;
		return offset + 2;
	}

	writeUInt16BE(value: number, offset = 0): number {
		this[offset] = (value >>> 8) & 0xff;
		this[offset + 1] = value & 0xff;
		return offset + 2;
	}

	writeInt16LE(value: number, offset = 0): number {
		return this.writeUInt16LE(value & 0xffff, offset);
	}

	writeUInt32LE(value: number, offset = 0): number {
		this[offset] = value & 0xff;
		this[offset + 1] = (value >>> 8) & 0xff;
		this[offset + 2] = (value >>> 16) & 0xff;
		this[offset + 3] = (value >>> 24) & 0xff;
		return offset + 4;
	}

	writeUInt32BE(value: number, offset = 0): number {
		this[offset] = (value >>> 24) & 0xff;
		this[offset + 1] = (value >>> 16) & 0xff;
		this[offset + 2] = (value >>> 8) & 0xff;
		this[offset + 3] = value & 0xff;
		return offset + 4;
	}

	writeInt32LE(value: number, offset = 0): number {
		return this.writeUInt32LE(value >>> 0, offset);
	}

	writeInt32BE(value: number, offset = 0): number {
		return this.writeUInt32BE(value >>> 0, offset);
	}

	writeUIntLE(value: number, offset: number, byteLength: number): number {
		let rest = value;
		for (let i = 0; i < byteLength; i++) {
			this[offset + i] = rest & 0xff;
			rest = Math.floor(rest / 0x100);
		}
		return offset + byteLength;
	}

	writeUIntBE(value: number, offset: number, byteLength: number): number {
		let rest = value;
		for (let i = byteLength - 1; i >= 0; i--) {
			this[offset + i] = rest & 0xff;
			rest = Math.floor(rest / 0x100);
		}
		return offset + byteLength;
	}

	writeIntLE(value: number, offset: number, byteLength: number): number {
		return this.writeUIntLE(value < 0 ? value + Math.pow(2, 8 * byteLength) : value, offset, byteLength);
	}

	writeFloatLE(value: number, offset = 0): number {
		this.view().setFloat32(offset, value, true);
		return offset + 4;
	}

	writeFloatBE(value: number, offset = 0): number {
		this.view().setFloat32(offset, value, false);
		return offset + 4;
	}

	writeDoubleLE(value: number, offset = 0): number {
		this.view().setFloat64(offset, value, true);
		return offset + 8;
	}

	writeDoubleBE(value: number, offset = 0): number {
		this.view().setFloat64(offset, value, false);
		return offset + 8;
	}

	writeBigUInt64LE(value: bigint, offset = 0): number {
		this.view().setBigUint64(offset, value, true);
		return offset + 8;
	}

	writeBigUInt64BE(value: bigint, offset = 0): number {
		this.view().setBigUint64(offset, value, false);
		return offset + 8;
	}

	writeBigInt64LE(value: bigint, offset = 0): number {
		this.view().setBigInt64(offset, value, true);
		return offset + 8;
	}

	writeBigInt64BE(value: bigint, offset = 0): number {
		this.view().setBigInt64(offset, value, false);
		return offset + 8;
	}

	// ---- bytes and text ----

	toString(encoding?: string, start?: number, end?: number): string {
		const from = clamp(start, 0, this.length);
		const to = clamp(end, this.length, this.length);
		if (to <= from) {
			return '';
		}
		return decodeBytes(this.subarray(from, to), normalizeEncoding(encoding));
	}

	write(text: string, offset?: number, length?: number, encoding?: string): number {
		// Node's overloads: (text), (text, encoding), (text, offset, encoding),
		// (text, offset, length, encoding).
		let at = 0;
		let max = this.length;
		let enc = encoding;
		if (typeof offset === 'string') {
			enc = offset;
		} else if (offset !== undefined) {
			at = offset;
			max = this.length - at;
			if (typeof length === 'string') {
				enc = length;
			} else if (length !== undefined) {
				max = Math.min(length, max);
			}
		}
		const resolved = normalizeEncoding(enc);
		const bytes = encodeString(text, resolved);
		const count = wholeCharacters(bytes, Math.min(bytes.length, max), resolved);
		this.set(bytes.subarray(0, count), at);
		return count;
	}

	/**
	 * Shares memory, like Node's. Uint8Array's slice copies, and inheriting
	 * that would break every in-place edit the engine makes through a slice.
	 */
	slice(start?: number, end?: number): WebBufferClass {
		return this.subarray(start, end) as WebBufferClass;
	}

	equals(other: Uint8Array): boolean {
		if (other.length !== this.length) {
			return false;
		}
		for (let i = 0; i < this.length; i++) {
			if (this[i] !== other[i]) {
				return false;
			}
		}
		return true;
	}

	compare(other: Uint8Array): number {
		const shared = Math.min(this.length, other.length);
		for (let i = 0; i < shared; i++) {
			if (this[i] !== other[i]) {
				return this[i] < other[i] ? -1 : 1;
			}
		}
		if (this.length === other.length) {
			return 0;
		}
		return this.length < other.length ? -1 : 1;
	}

	copy(target: Uint8Array, targetStart = 0, sourceStart = 0, sourceEnd = this.length): number {
		const from = clamp(sourceStart, 0, this.length);
		const to = clamp(sourceEnd, this.length, this.length);
		const count = Math.min(to - from, target.length - targetStart);
		if (count <= 0) {
			return 0;
		}
		target.set(this.subarray(from, from + count), targetStart);
		return count;
	}

	fill(value: number | string, start?: number, end?: number, encoding?: string): this {
		const from = clamp(start, 0, this.length);
		const to = clamp(end, this.length, this.length);
		if (typeof value === 'number') {
			super.fill(value & 0xff, from, to);
			return this;
		}
		const bytes = encodeString(value, normalizeEncoding(encoding));
		if (bytes.length === 0) {
			super.fill(0, from, to);
			return this;
		}
		for (let i = from; i < to; i++) {
			this[i] = bytes[(i - from) % bytes.length];
		}
		return this;
	}

	indexOf(value: number | string | Uint8Array, byteOffset?: number, encoding?: string): number {
		return this.search(value, byteOffset, encoding, false);
	}

	lastIndexOf(value: number | string | Uint8Array, byteOffset?: number, encoding?: string): number {
		return this.search(value, byteOffset, encoding, true);
	}

	includes(value: number | string | Uint8Array, byteOffset?: number, encoding?: string): boolean {
		return this.indexOf(value, byteOffset, encoding) !== -1;
	}

	private search(
		value: number | string | Uint8Array,
		byteOffset: number | undefined,
		encoding: string | undefined,
		last: boolean,
	): number {
		let needle: Uint8Array;
		if (typeof value === 'number') {
			needle = Uint8Array.of(value & 0xff);
		} else if (typeof value === 'string') {
			needle = encodeString(value, normalizeEncoding(encoding));
		} else {
			needle = value;
		}
		if (needle.length === 0) {
			return last ? this.length : 0;
		}
		if (needle.length > this.length) {
			return -1;
		}

		const lastStart = this.length - needle.length;
		let start: number;
		if (byteOffset === undefined) {
			start = last ? lastStart : 0;
		} else {
			start = Math.trunc(byteOffset);
			if (start < 0) {
				start += this.length;
			}
			start = last ? Math.min(start, lastStart) : Math.max(start, 0);
		}

		const step = last ? -1 : 1;
		for (let at = start; at >= 0 && at <= lastStart; at += step) {
			let matched = true;
			for (let i = 0; i < needle.length; i++) {
				if (this[at + i] !== needle[i]) {
					matched = false;
					break;
				}
			}
			if (matched) {
				return at;
			}
		}
		return -1;
	}
}

/**
 * Buffer's statics are attached rather than declared in the class body.
 * TypeScript checks a subclass's static side against Uint8Array's, and
 * Buffer.from(value, encoding) is deliberately incompatible with
 * Uint8Array.from(arrayLike, mapfn) - which is exactly why Node declares
 * Buffer as an interface rather than as a subclass.
 */
const statics = {
	alloc(size: number, fill?: number | string, encoding?: string): WebBufferClass {
		const buffer = new WebBufferClass(size);
		if (fill !== undefined && fill !== 0) {
			buffer.fill(fill as never, 0, size, encoding);
		}
		return buffer;
	},

	allocUnsafe(size: number): WebBufferClass {
		return new WebBufferClass(size);
	},

	from(
		value: string | ArrayLike<number> | ArrayBufferLike | Uint8Array,
		encodingOrOffset?: string | number,
		length?: number,
	): WebBufferClass {
		if (typeof value === 'string') {
			const bytes = encodeString(value, normalizeEncoding(encodingOrOffset as string));
			const out = new WebBufferClass(bytes.length);
			out.set(bytes);
			return out;
		}
		if (value instanceof Uint8Array) {
			// Copies, as Node does for this overload.
			const out = new WebBufferClass(value.length);
			out.set(value);
			return out;
		}
		if (value instanceof ArrayBuffer) {
			// Shares memory, as Node does for this overload.
			const offset = (encodingOrOffset as number) ?? 0;
			const count = length ?? value.byteLength - offset;
			return new WebBufferClass(value, offset, count);
		}
		const array = value as ArrayLike<number>;
		const out = new WebBufferClass(array.length);
		for (let i = 0; i < array.length; i++) {
			out[i] = array[i] & 0xff;
		}
		return out;
	},

	concat(list: readonly Uint8Array[], totalLength?: number): WebBufferClass {
		let total = totalLength;
		if (total === undefined) {
			total = 0;
			for (const part of list) {
				total += part.length;
			}
		}
		const out = new WebBufferClass(total);
		let at = 0;
		for (const part of list) {
			if (at >= total) {
				break;
			}
			const room = total - at;
			out.set(part.length > room ? part.subarray(0, room) : part, at);
			at += part.length;
		}
		return out;
	},

	isBuffer(value: unknown): boolean {
		return value instanceof WebBufferClass;
	},

	/**
	 * The number of bytes the string WOULD take, which for hex and base64 is
	 * a formula over the length rather than a decode: Node reports
	 * byteLength('Module1', 'hex') as 3 even though no pair of those
	 * characters is valid hex and from() returns nothing.
	 */
	byteLength(value: string | Uint8Array, encoding?: string): number {
		if (typeof value !== 'string') {
			return value.length;
		}
		switch (normalizeEncoding(encoding)) {
			case 'latin1':
			case 'ascii':
				return value.length;
			case 'utf16le':
				return value.length * 2;
			case 'hex':
				return value.length >>> 1;
			case 'base64':
			case 'base64url': {
				let length = value.length;
				if (value.charCodeAt(length - 1) === 0x3d) {
					length--;
				}
				if (length > 1 && value.charCodeAt(length - 1) === 0x3d) {
					length--;
				}
				return (length * 3) >>> 2;
			}
			case 'utf8':
				return utf8Encoder.encode(value).length;
		}
	},
};

/**
 * Declared explicitly because an intersection of the class with `statics`
 * leaves `from` ambiguous: TypeScript resolves it to Uint8Array's inherited
 * static rather than to Buffer's (value, encoding) form.
 */
interface WebBufferConstructor {
	new (length: number): WebBufferClass;
	new (buffer: ArrayBufferLike, byteOffset?: number, length?: number): WebBufferClass;
	readonly prototype: WebBufferClass;
	alloc(size: number, fill?: number | string, encoding?: string): WebBufferClass;
	allocUnsafe(size: number): WebBufferClass;
	from(
		value: string | ArrayLike<number> | ArrayBufferLike | Uint8Array,
		encodingOrOffset?: string | number,
		length?: number,
	): WebBufferClass;
	concat(list: readonly Uint8Array[], totalLength?: number): WebBufferClass;
	isBuffer(value: unknown): boolean;
	byteLength(value: string | Uint8Array, encoding?: string): number;
}

export type WebBuffer = WebBufferClass;
export const WebBuffer = Object.assign(
	WebBufferClass,
	statics,
) as unknown as WebBufferConstructor;
