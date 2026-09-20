// A synchronous `crypto` for the web extension host.
//
// The browser's own SubtleCrypto.digest is asynchronous, and the four callers
// here - module content tokens, the project-operations marker, the form
// preview's temp name, the dirty-module backup key - all compute a hash
// inline while building a string. Making them async would ripple through
// their callers for no gain, so SHA-256 and SHA-1 are implemented directly.
// They are ~150 lines of fully specified arithmetic, and node:crypto is the
// oracle: tests/webCrypto.test.ts asserts identical digests.
//
// These are identity hashes, not security boundaries: they name a cached
// token or a backup file. They still have to agree with the desktop build's
// digests exactly, or the two would disagree about whether a module changed.
//
// The browser build aliases `crypto` to this module (see webBuild.js), so the
// five files that import it are unchanged.

import { WebBuffer } from './webBuffer';

const SHA256_K = new Uint32Array([
	0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
	0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
	0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
	0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
	0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
	0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
	0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
	0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

/** Length-prefixes and pads a message to a multiple of 64 bytes (both algorithms). */
function padMessage(message: Uint8Array): Uint8Array {
	const bitLength = message.length * 8;
	const padded = new Uint8Array(((message.length + 9 + 63) >> 6) << 6);
	padded.set(message);
	padded[message.length] = 0x80;
	// The length is a 64-bit big-endian count of bits. Messages here are far
	// below 2^32 bits, so the high word stays zero.
	const view = new DataView(padded.buffer);
	view.setUint32(padded.length - 4, bitLength >>> 0, false);
	view.setUint32(padded.length - 8, Math.floor(bitLength / 0x100000000), false);
	return padded;
}

function sha256(message: Uint8Array): Uint8Array {
	const h = new Uint32Array([
		0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
	]);
	const padded = padMessage(message);
	const view = new DataView(padded.buffer);
	const w = new Uint32Array(64);

	for (let block = 0; block < padded.length; block += 64) {
		for (let i = 0; i < 16; i++) {
			w[i] = view.getUint32(block + i * 4, false);
		}
		for (let i = 16; i < 64; i++) {
			const a = w[i - 15];
			const b = w[i - 2];
			const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
			const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
			w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
		}

		let [a, b, c, d, e, f, g, hh] = h;
		for (let i = 0; i < 64; i++) {
			const s1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
			const ch = (e & f) ^ (~e & g);
			const temp1 = (hh + s1 + ch + SHA256_K[i] + w[i]) >>> 0;
			const s0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
			const maj = (a & b) ^ (a & c) ^ (b & c);
			const temp2 = (s0 + maj) >>> 0;

			hh = g;
			g = f;
			f = e;
			e = (d + temp1) >>> 0;
			d = c;
			c = b;
			b = a;
			a = (temp1 + temp2) >>> 0;
		}

		h[0] = (h[0] + a) >>> 0;
		h[1] = (h[1] + b) >>> 0;
		h[2] = (h[2] + c) >>> 0;
		h[3] = (h[3] + d) >>> 0;
		h[4] = (h[4] + e) >>> 0;
		h[5] = (h[5] + f) >>> 0;
		h[6] = (h[6] + g) >>> 0;
		h[7] = (h[7] + hh) >>> 0;
	}

	const out = new Uint8Array(32);
	const outView = new DataView(out.buffer);
	for (let i = 0; i < 8; i++) {
		outView.setUint32(i * 4, h[i], false);
	}
	return out;
}

function sha1(message: Uint8Array): Uint8Array {
	const h = new Uint32Array([0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0]);
	const padded = padMessage(message);
	const view = new DataView(padded.buffer);
	const w = new Uint32Array(80);

	for (let block = 0; block < padded.length; block += 64) {
		for (let i = 0; i < 16; i++) {
			w[i] = view.getUint32(block + i * 4, false);
		}
		for (let i = 16; i < 80; i++) {
			const x = w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16];
			w[i] = (x << 1) | (x >>> 31);
		}

		let [a, b, c, d, e] = h;
		for (let i = 0; i < 80; i++) {
			let f: number;
			let k: number;
			if (i < 20) {
				f = (b & c) | (~b & d);
				k = 0x5a827999;
			} else if (i < 40) {
				f = b ^ c ^ d;
				k = 0x6ed9eba1;
			} else if (i < 60) {
				f = (b & c) | (b & d) | (c & d);
				k = 0x8f1bbcdc;
			} else {
				f = b ^ c ^ d;
				k = 0xca62c1d6;
			}
			const temp = (((a << 5) | (a >>> 27)) + f + e + k + w[i]) >>> 0;
			e = d;
			d = c;
			c = (b << 30) | (b >>> 2);
			b = a;
			a = temp;
		}

		h[0] = (h[0] + a) >>> 0;
		h[1] = (h[1] + b) >>> 0;
		h[2] = (h[2] + c) >>> 0;
		h[3] = (h[3] + d) >>> 0;
		h[4] = (h[4] + e) >>> 0;
	}

	const out = new Uint8Array(20);
	const outView = new DataView(out.buffer);
	for (let i = 0; i < 5; i++) {
		outView.setUint32(i * 4, h[i], false);
	}
	return out;
}

type Algorithm = 'sha256' | 'sha1';

/** The subset of node:crypto's Hash that XLIDE uses. */
export class WebHash {
	private readonly chunks: Uint8Array[] = [];

	constructor(private readonly algorithm: Algorithm) {}

	update(data: string | Uint8Array, encoding?: string): this {
		this.chunks.push(
			typeof data === 'string' ? WebBuffer.from(data, encoding ?? 'utf8') : data,
		);
		return this;
	}

	digest(): WebBuffer;
	digest(encoding: 'hex' | 'base64'): string;
	digest(encoding?: 'hex' | 'base64'): WebBuffer | string {
		const message = WebBuffer.concat(this.chunks);
		const bytes = this.algorithm === 'sha256' ? sha256(message) : sha1(message);
		const buffer = WebBuffer.from(bytes);
		return encoding === undefined ? buffer : buffer.toString(encoding);
	}
}

export function createHash(algorithm: string): WebHash {
	const normalized = algorithm.toLowerCase().replace('-', '');
	if (normalized !== 'sha256' && normalized !== 'sha1') {
		throw new Error(
			`XLIDE in the browser implements sha256 and sha1, not ${algorithm}.`,
		);
	}
	return new WebHash(normalized);
}

export function randomBytes(size: number): WebBuffer {
	const out = new Uint8Array(size);
	crypto.getRandomValues(out);
	return WebBuffer.from(out);
}

export function randomUUID(): string {
	return crypto.randomUUID();
}

export default { createHash, randomBytes, randomUUID };
