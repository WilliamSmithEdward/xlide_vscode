// MS-OVBA compression - [MS-OVBA] 2.4.1.
//
// VBA stores the dir stream and every module's source as a "CompressedContainer":
// a 0x01 signature byte followed by chunks. Each chunk decompresses to at most
// 4096 bytes and is encoded as groups of (flag byte, up to 8 tokens), where a
// token is either a literal byte or a 16-bit back-reference whose length/offset
// bit split widens as the chunk fills.

export class OvbaError extends Error {}

/** Number of bits needed to represent n. */
function bitLength(n: number): number {
	return n === 0 ? 0 : 32 - Math.clz32(n);
}

/**
 * Length/offset masks and offset bit-count for a copy token at the given
 * position within the decompressed chunk ([MS-OVBA] 2.4.1.3.6).
 */
function copyTokenHelp(decompressedCurrent: number, decompressedChunkStart: number): {
	lengthMask: number;
	offsetMask: number;
	bitCount: number;
} {
	const difference = Math.max(1, decompressedCurrent - decompressedChunkStart);
	const bitCount = Math.max(4, bitLength(difference - 1));
	const lengthMask = 0xffff >>> bitCount;
	const offsetMask = (~lengthMask) & 0xffff;
	return { lengthMask, offsetMask, bitCount };
}

/**
 * Decompress an [MS-OVBA] 2.4.1 stream.
 *
 * `maxBytes` stops after the first chunk that reaches that many decompressed
 * bytes and returns the prefix. Copy tokens are chunk-local by construction
 * (`copyTokenHelp` measures from the chunk's decompressed start), so no later
 * chunk can change a byte an earlier one produced: a chunk-aligned prefix is
 * always exactly what a full decompression would have produced there. Callers
 * that only need a module's attribute header use this to skip inflating
 * megabytes of body they will not read.
 */
export function decompress(data: Buffer, streamName = '<unknown>', maxBytes = Infinity): Buffer {
	const err = (msg: string, offset: number): OvbaError =>
		new OvbaError(`${msg} [stream=${streamName}, offset=${offset}]`);

	if (data.length === 0 || data[0] !== 0x01) {
		throw err('Invalid compressed stream: missing 0x01 signature byte.', 0);
	}

	let pos = 1;
	// Growable output; VBA chunks decompress to <= 4096 bytes each.
	let out = Buffer.alloc(Math.max(4096, data.length * 4));
	let len = 0;
	const ensure = (extra: number): void => {
		if (len + extra <= out.length) { return; }
		let size = out.length * 2;
		while (size < len + extra) { size *= 2; }
		const next = Buffer.alloc(size);
		out.copy(next, 0, 0, len);
		out = next;
	};

	while (pos < data.length) {
		if (len >= maxBytes) { break; }
		if (pos + 2 > data.length) {
			throw err('Truncated compressed stream: missing chunk header.', pos);
		}
		const header = data.readUInt16LE(pos);
		const chunkDataSize = (header & 0x0fff) + 1;
		const chunkSignature = (header >> 12) & 0x7;
		const chunkFlag = (header >> 15) & 0x1;
		const headerOffset = pos;
		pos += 2;

		if (chunkSignature !== 0b011) {
			throw err(`Bad compressed chunk signature: expected 0b011, got ${chunkSignature}.`, headerOffset);
		}
		const chunkEnd = pos + chunkDataSize;
		if (chunkEnd > data.length) {
			throw err(
				`Truncated chunk: header announces ${chunkDataSize} bytes but only ${data.length - pos} remain.`,
				headerOffset,
			);
		}
		const decompressedChunkStart = len;

		if (chunkFlag === 0) {
			if (chunkDataSize !== 4096) {
				throw err(`Raw chunk must have exactly 4096 data bytes; got ${chunkDataSize}.`, headerOffset);
			}
			if (pos + 4096 > data.length) {
				throw err('Truncated raw chunk.', pos);
			}
			ensure(4096);
			data.copy(out, len, pos, pos + 4096);
			len += 4096;
			pos += 4096;
			continue;
		}

		while (pos < chunkEnd) {
			if (pos >= data.length) { break; }
			const flagByte = data[pos];
			pos += 1;
			for (let bit = 0; bit < 8; bit++) {
				if (pos >= chunkEnd || pos >= data.length) { break; }
				if ((flagByte >> bit) & 1) {
					if (pos + 2 > data.length) {
						throw err('Truncated copy token.', pos);
					}
					const token = data.readUInt16LE(pos);
					pos += 2;
					const { lengthMask, offsetMask, bitCount } = copyTokenHelp(len, decompressedChunkStart);
					const length = (token & lengthMask) + 3;
					const offset = ((token & offsetMask) >>> (16 - bitCount)) + 1;
					let copySrc = len - offset;
					if (copySrc < 0) {
						throw err('Copy token references before start of output.', pos - 2);
					}
					if (copySrc < decompressedChunkStart) {
						throw err('Copy token references before the start of the current chunk.', pos - 2);
					}
					ensure(length);
					if (offset >= length) {
						// Source and destination do not overlap, so the whole
						// run can move at once.
						out.copy(out, len, copySrc, copySrc + length);
						len += length;
					} else {
						// Overlapping copy: the spec's byte-at-a-time semantics
						// are load-bearing here, since later bytes read back
						// what this same loop just wrote.
						for (let k = 0; k < length; k++) {
							out[len++] = out[copySrc++];
						}
					}
				} else {
					ensure(1);
					out[len++] = data[pos];
					pos += 1;
				}
			}
		}
	}

	return Buffer.from(out.subarray(0, len));
}

export function compress(data: Buffer): Buffer {
	const parts: Buffer[] = [Buffer.from([0x01])];
	let cursor = 0;
	while (cursor < data.length) {
		const chunk = data.subarray(cursor, cursor + 4096);
		cursor += chunk.length;
		const encoded = encodeLz(chunk);
		if (encoded.length <= 4096) {
			const header = Buffer.alloc(2);
			header.writeUInt16LE(0xb000 | (encoded.length - 1), 0);
			parts.push(header, encoded);
			continue;
		}
		// Token encoding overflowed (high-entropy input). A raw chunk is only
		// legal at exactly 4096 data bytes.
		if (chunk.length !== 4096) {
			throw new OvbaError(
				`Final partial chunk (${chunk.length} bytes) encoded to ${encoded.length} bytes, `
				+ 'exceeding the 4096-byte chunk limit.',
			);
		}
		const header = Buffer.alloc(2);
		header.writeUInt16LE(0x3fff, 0);
		parts.push(header, Buffer.from(chunk));
	}
	return Buffer.concat(parts);
}

/**
 * LZ encoder for a single chunk, choosing its copy tokens as Office does
 * (issue #149, from pyOpenVBA 6.3.0): Office's picks are those of the LZNT1
 * standard engine in ntdll (RtlCompressBuffer), so its output is byte for
 * byte what Office writes. Measured on the 160 dir and module streams in
 * this repository's fixtures, from Excel, Word and PowerPoint, .xls, .xlsb
 * and .doc included: all 160 compress back unchanged. The longest-match
 * search this replaced reproduced 22 of them.
 *
 * - Only a position where a token starts is remembered, in one of 4096
 *   buckets chosen from the three bytes that start there. A bucket holds two
 *   positions, the newest and the one before it. A lookup reads both, then
 *   makes the current position the newest.
 * - Each candidate is measured up to the longest match a token here can
 *   hold, capped at the chunk's end. The longer wins and the newer wins a
 *   tie; under three bytes is a literal. A position with fewer than three
 *   bytes left is a literal and is not remembered.
 * - When the chunk's last token fills its flag byte, one more flag byte,
 *   empty, follows. ntdll does not write it; Office does.
 */
function encodeLz(chunk: Buffer): Buffer {
	const chunkLen = chunk.length;
	// A flag byte per eight tokens, a token at most two bytes per input byte
	// when every one is a literal, and the trailing empty flag byte.
	const out = new Uint8Array(chunkLen + Math.ceil(chunkLen / 8) + 2);
	let length = 0;
	const newest = new Int32Array(4096).fill(-1);
	const older = new Int32Array(4096).fill(-1);

	let pos = 0;
	let lastGroupFull = false;
	while (pos < chunkLen) {
		const flagAt = length++;
		let flagBits = 0;
		let bit = 0;
		for (; bit < 8 && pos < chunkLen; bit++) {
			const { lengthMask, bitCount } = copyTokenHelp(pos, 0);
			const maxLength = Math.min(lengthMask + 3, chunkLen - pos);

			let bestLen = 0;
			let bestOffset = 0;
			if (chunkLen - pos >= 3) {
				// 40543 * a 16-bit value passes 2^31 but stays under 2^32, so
				// `>>>` reads it as the unsigned product it is.
				const key = (chunk[pos] << 8) ^ (chunk[pos + 1] << 4) ^ chunk[pos + 2];
				const bucket = ((40543 * key) >>> 4) & 0xfff;
				const first = newest[bucket];
				const second = older[bucket];
				older[bucket] = first;
				newest[bucket] = pos;
				for (const cand of [first, second]) {
					if (cand < 0) { continue; }
					let matchLen = 0;
					while (matchLen < maxLength && chunk[cand + matchLen] === chunk[pos + matchLen]) {
						matchLen++;
					}
					// The newer candidate is measured first, so only a
					// strictly longer older one displaces it.
					if (matchLen > bestLen) {
						bestLen = matchLen;
						bestOffset = pos - cand;
					}
				}
			}

			if (bestLen >= 3) {
				flagBits |= 1 << bit;
				const token = (((bestOffset - 1) << (16 - bitCount)) | (bestLen - 3)) & 0xffff;
				out[length++] = token & 0xff;
				out[length++] = token >>> 8;
				pos += bestLen;
			} else {
				out[length++] = chunk[pos];
				pos += 1;
			}
		}
		out[flagAt] = flagBits;
		lastGroupFull = bit === 8;
	}
	if (lastGroupFull) {
		out[length++] = 0;
	}
	return Buffer.from(out.subarray(0, length));
}
