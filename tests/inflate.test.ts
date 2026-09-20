import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { InflateError, inflate, inflateRaw } from '../src/vba/inflate';

// zlib is the oracle: every case below deflates with Node and asserts that
// the pure-TS inflate returns bytes identical to inflateRawSync's. The web
// extension host has no zlib, so this is what stands between a browser build
// and silently corrupting someone's workbook.

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'binaries');

/** Every deflate level, including 0 (stored blocks) and 9 (deepest search). */
const LEVELS = [0, 1, 2, 4, 6, 9];

function roundTrip(data: Buffer, level: number): void {
	const deflated = zlib.deflateRawSync(data, { level });
	const ours = inflateRaw(deflated, { expectedSize: data.length });
	expect(ours.length).toBe(data.length);
	expect(ours.equals(data)).toBe(true);
}

describe('inflateRaw', () => {
	it('matches zlib on an empty payload', () => {
		for (const level of LEVELS) {
			roundTrip(Buffer.alloc(0), level);
		}
	});

	it('matches zlib on a single byte', () => {
		for (const level of LEVELS) {
			roundTrip(Buffer.from([0x41]), level);
		}
	});

	it('matches zlib on highly repetitive data (long back-references)', () => {
		const data = Buffer.from('Sub Foo()\r\n    Debug.Print 1\r\nEnd Sub\r\n'.repeat(5000));
		for (const level of LEVELS) {
			roundTrip(data, level);
		}
	});

	it('matches zlib on a run of one byte (maximum-length matches)', () => {
		const data = Buffer.alloc(300000, 0x20);
		for (const level of LEVELS) {
			roundTrip(data, level);
		}
	});

	it('matches zlib on incompressible data (multiple stored blocks)', () => {
		// Deterministic pseudo-random bytes: incompressible, so level 0 emits
		// a chain of stored blocks (65535 bytes each) and the higher levels
		// fall back to stored blocks too.
		const data = Buffer.alloc(200000);
		let state = 0x12345678;
		for (let i = 0; i < data.length; i++) {
			state = (Math.imul(state, 1103515245) + 12345) >>> 0;
			data[i] = (state >>> 16) & 0xff;
		}
		for (const level of LEVELS) {
			roundTrip(data, level);
		}
	});

	it('matches zlib on data that mixes compressible and random regions', () => {
		const parts: Buffer[] = [];
		let state = 0xc0ffee;
		for (let block = 0; block < 40; block++) {
			parts.push(Buffer.from('<row r="1" spans="1:4">'.repeat(200)));
			const noise = Buffer.alloc(3000);
			for (let i = 0; i < noise.length; i++) {
				state = (Math.imul(state, 1103515245) + 12345) >>> 0;
				noise[i] = (state >>> 16) & 0xff;
			}
			parts.push(noise);
		}
		const data = Buffer.concat(parts);
		for (const level of LEVELS) {
			roundTrip(data, level);
		}
	});

	it('matches zlib on every byte value', () => {
		const data = Buffer.alloc(256 * 64);
		for (let i = 0; i < data.length; i++) {
			data[i] = i & 0xff;
		}
		for (const level of LEVELS) {
			roundTrip(data, level);
		}
	});

	it('matches zlib on fixed-Huffman blocks', () => {
		// Small inputs take the fixed-Huffman path, where no dynamic table is
		// worth its header.
		for (let size = 1; size <= 64; size++) {
			const data = Buffer.from('ABCABCABC'.repeat(size).slice(0, size));
			roundTrip(data, 9);
		}
	});

	it('does not need the expected size to be supplied', () => {
		const data = Buffer.from('Attribute VB_Name = "Module1"\r\n'.repeat(400));
		const deflated = zlib.deflateRawSync(data, { level: 6 });
		expect(inflateRaw(deflated).equals(data)).toBe(true);
	});

	it('matches zlib on the bytes of every binary fixture in the repo', () => {
		const files = fs.readdirSync(FIXTURE_DIR).filter((f) => !f.startsWith('.'));
		expect(files.length).toBeGreaterThan(20);
		for (const file of files) {
			const data = fs.readFileSync(path.join(FIXTURE_DIR, file));
			for (const level of [0, 6, 9]) {
				const deflated = zlib.deflateRawSync(data, { level });
				const ours = inflateRaw(deflated, { expectedSize: data.length });
				expect(ours.equals(data), `${file} at level ${level}`).toBe(true);
			}
		}
	});
});

describe('inflate (zlib-wrapped)', () => {
	it('matches zlib through the RFC 1950 wrapper', () => {
		const data = Buffer.from('PowerPoint compressed storage payload.'.repeat(500));
		for (const level of LEVELS) {
			const deflated = zlib.deflateSync(data, { level });
			expect(inflate(deflated, { expectedSize: data.length }).equals(data)).toBe(true);
		}
	});

	it('rejects a stream that is not zlib-wrapped', () => {
		const raw = zlib.deflateRawSync(Buffer.from('not wrapped'));
		expect(() => inflate(raw)).toThrow(InflateError);
	});
});

describe('allowTruncated', () => {
	// PowerPoint stores VBA inside a compressed CFB storage whose deflate
	// stream stops without a final block, which zlib reads with a
	// Z_SYNC_FLUSH finish. These assert the same tolerance, and that it stays
	// narrow: a malformed stream still throws with allowTruncated set.
	const data = Buffer.from('Attribute VB_Name = "Slide1"\r\n'.repeat(2000));

	it('returns the decoded prefix when the stream ends without a final block', () => {
		const deflated = zlib.deflateRawSync(data, { level: 6 });
		const cut = deflated.subarray(0, deflated.length - 2);

		const zlibPrefix = zlib.inflateRawSync(cut, { finishFlush: zlib.constants.Z_SYNC_FLUSH });
		const ours = inflateRaw(cut, { allowTruncated: true });

		expect(ours.equals(zlibPrefix)).toBe(true);
		expect(ours.length).toBeGreaterThan(0);
		expect(data.subarray(0, ours.length).equals(ours)).toBe(true);
	});

	it('matches zlib at every truncation point of a real stream', () => {
		const deflated = zlib.deflateRawSync(data, { level: 6 });
		for (let cut = 1; cut < deflated.length; cut += 7) {
			const partial = deflated.subarray(0, cut);
			let expected: Buffer;
			try {
				expected = zlib.inflateRawSync(partial, {
					finishFlush: zlib.constants.Z_SYNC_FLUSH,
				});
			} catch {
				continue; // zlib calls this one malformed, not merely short.
			}
			const ours = inflateRaw(partial, { allowTruncated: true });
			expect(ours.equals(expected), `cut at ${cut} of ${deflated.length}`).toBe(true);
		}
	});

	it('still throws on a malformed stream', () => {
		expect(() => inflateRaw(Buffer.from([0x07]), { allowTruncated: true })).toThrow(
			InflateError,
		);
		expect(() => inflateRaw(Buffer.from([0x03, 0x42]), { allowTruncated: true })).toThrow(
			InflateError,
		);
	});
});

describe('malformed input', () => {
	it('throws rather than looping on a truncated stream', () => {
		const deflated = zlib.deflateRawSync(Buffer.from('x'.repeat(10000)), { level: 9 });
		expect(() => inflateRaw(deflated.subarray(0, deflated.length - 3))).toThrow(InflateError);
	});

	it('throws on a reserved block type', () => {
		// 1 final bit, then block type 3.
		expect(() => inflateRaw(Buffer.from([0x07]))).toThrow(InflateError);
	});

	it('throws on a stored block whose length check fails', () => {
		// final=1, type=0, aligned, then LEN=4 with a wrong NLEN.
		expect(() => inflateRaw(Buffer.from([0x01, 0x04, 0x00, 0x00, 0x00, 1, 2, 3, 4]))).toThrow(
			InflateError,
		);
	});

	it('throws on a back-reference that precedes the output', () => {
		// A fixed-Huffman block whose first symbol is a length/distance pair:
		// final bit, type 01, length symbol 257, distance symbol 1. zlib
		// rejects the same two bytes with "invalid distance too far back".
		const bytes = Buffer.from([0x03, 0x42]);
		expect(() => zlib.inflateRawSync(bytes)).toThrow();
		expect(() => inflateRaw(bytes)).toThrow(InflateError);
	});
});
