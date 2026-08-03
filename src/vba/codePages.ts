// ANSI code-page text conversion for VBA project streams.
//
// [MS-OVBA] stores every ANSI string - module sources, dir-stream names, the
// PROJECT stream - in the project's PROJECTCODEPAGE. The first native engine
// implemented only cp1252 and fell back to latin1 for everything else, which
// turned a Russian (cp1251) project's text into mojibake ('Модуль'
// read as 'Ìîäóëü'). Decoding rides TextDecoder, which supports every
// Windows code page VBA uses (official Node and VS Code's Electron ship full
// ICU); encoding uses reverse tables built lazily from the decoder itself, so
// the two directions can never disagree.

import { TextDecoder } from 'util';

/** Windows code page -> WHATWG encoding label. */
const CODE_PAGE_LABELS: Record<number, string> = {
	874: 'windows-874',
	932: 'shift_jis',
	936: 'gbk',
	949: 'euc-kr',
	950: 'big5',
	1250: 'windows-1250',
	1251: 'windows-1251',
	1252: 'windows-1252',
	1253: 'windows-1253',
	1254: 'windows-1254',
	1255: 'windows-1255',
	1256: 'windows-1256',
	1257: 'windows-1257',
	1258: 'windows-1258',
	10000: 'macintosh',
	20866: 'koi8-r',
	21866: 'koi8-u',
	28591: 'iso-8859-1',
	28592: 'iso-8859-2',
	28595: 'iso-8859-5',
	65001: 'utf-8',
};

const DOUBLE_BYTE_PAGES = new Set([932, 936, 949, 950]);

/** Every code page with first-class conversion support, for the CI language matrix. */
export function supportedCodePages(): number[] {
	return Object.keys(CODE_PAGE_LABELS).map(Number);
}

/** WHATWG label for a supported page, so tests can probe ICU directly. */
export function codePageLabel(codePage: number): string | undefined {
	return CODE_PAGE_LABELS[codePage];
}

// cp1252 differs from latin-1 only in 0x80-0x9F. Kept as a table (not via
// TextDecoder) so the hot default path stays allocation-free and byte-exact
// with the engine's original behavior.
const CP1252_HIGH = [
	0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
	0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
	0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
	0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
];

const decoders = new Map<number, TextDecoder | null>();

function decoderFor(codePage: number): TextDecoder | null {
	let cached = decoders.get(codePage);
	if (cached !== undefined) { return cached; }
	const label = CODE_PAGE_LABELS[codePage];
	if (!label) {
		decoders.set(codePage, null);
		return null;
	}
	try {
		cached = new TextDecoder(label);
	} catch {
		// Embedder without full ICU: fall back to the historical behavior.
		cached = null;
	}
	decoders.set(codePage, cached);
	return cached;
}

/** Decode code-page bytes to a string. Unknown pages fall back to cp1252. */
export function decodeCodePage(buf: Buffer, codePage: number): string {
	if (codePage === 1252) { return decodeCp1252(buf); }
	const decoder = decoderFor(codePage);
	if (decoder) { return decoder.decode(buf); }
	return decodeCp1252(buf);
}

function decodeCp1252(buf: Buffer): string {
	let out = '';
	for (const byte of buf) {
		out += byte >= 0x80 && byte <= 0x9f
			? String.fromCharCode(CP1252_HIGH[byte - 0x80])
			: String.fromCharCode(byte);
	}
	return out;
}

interface ReverseTable {
	/** char -> single byte */
	single: Map<string, number>;
	/** char -> [lead, trail] for double-byte pages */
	double: Map<string, [number, number]> | null;
}

const reverseTables = new Map<number, ReverseTable>();

/**
 * Build the char->bytes table by DECODING every byte (and, for CJK pages,
 * every two-byte sequence) through the page's own decoder. Derived rather
 * than transcribed, so encode(decode(x)) is x by construction. Lazy: only a
 * write of non-ASCII text on a non-1252 page pays for it, and only once.
 */
function reverseTableFor(codePage: number, decoder: TextDecoder): ReverseTable {
	let table = reverseTables.get(codePage);
	if (table) { return table; }
	const single = new Map<string, number>();
	const one = Buffer.alloc(1);
	for (let byte = 0; byte <= 0xff; byte++) {
		one[0] = byte;
		const ch = decoder.decode(one);
		if (ch.length === 1 && ch !== '�' && !single.has(ch)) {
			single.set(ch, byte);
		}
	}
	let double: ReverseTable['double'] = null;
	if (DOUBLE_BYTE_PAGES.has(codePage)) {
		double = new Map();
		const two = Buffer.alloc(2);
		for (let lead = 0x81; lead <= 0xfe; lead++) {
			for (let trail = 0x40; trail <= 0xfe; trail++) {
				two[0] = lead;
				two[1] = trail;
				const ch = decoder.decode(two);
				if (ch.length === 1 && ch !== '�' && !double.has(ch)) {
					double.set(ch, [lead, trail]);
				}
			}
		}
	}
	table = { single, double };
	reverseTables.set(codePage, table);
	return table;
}

const ASCII_ONLY = /^[\x00-\x7f]*$/;

/**
 * Encode a string to code-page bytes. Characters the page cannot represent
 * become '?', matching the engine's existing cp1252 behavior (and the old
 * Python backend's errors="replace").
 */
export function encodeCodePage(text: string, codePage: number): Buffer {
	if (codePage === 1252) { return encodeCp1252(text); }
	// UTF-8 is its own encoder; the reverse-table approach below cannot
	// express multibyte sequences.
	if (codePage === 65001) { return Buffer.from(text, 'utf8'); }
	// ASCII is byte-identical in every supported page.
	if (ASCII_ONLY.test(text)) { return Buffer.from(text, 'latin1'); }
	const decoder = decoderFor(codePage);
	if (!decoder) { return encodeCp1252(text); }
	const { single, double } = reverseTableFor(codePage, decoder);
	const out: number[] = [];
	const pushChar = (ch: string): boolean => {
		const one = single.get(ch);
		if (one !== undefined) {
			out.push(one);
			return true;
		}
		const two = double?.get(ch);
		if (two) {
			out.push(two[0], two[1]);
			return true;
		}
		return false;
	};
	// cp1258 (Vietnamese) stores accented vowels as a precomposed base plus a
	// combining tone byte (ệ = ê 0xEA + dot-below 0xF2), which is NOT the
	// char's canonical decomposition, so fold each mark back into the base in
	// turn until one combination maps.
	const pushDecomposed = (ch: string): boolean => {
		const parts = [...ch.normalize('NFD')];
		if (parts.length < 2) { return false; }
		const marks = parts.slice(1);
		if (single.has(parts[0]) && marks.every((m) => single.has(m))) {
			for (const part of parts) { pushChar(part); }
			return true;
		}
		for (let i = 0; i < marks.length; i++) {
			const composed = (parts[0] + marks[i]).normalize('NFC');
			const rest = marks.filter((_, j) => j !== i);
			if (composed.length === 1 && single.has(composed) && rest.every((m) => single.has(m))) {
				pushChar(composed);
				for (const m of rest) { pushChar(m); }
				return true;
			}
		}
		return false;
	};
	for (const ch of text) {
		if (pushChar(ch)) { continue; }
		if (pushDecomposed(ch)) { continue; }
		out.push(0x3f); // '?'
	}
	return Buffer.from(out);
}

function encodeCp1252(text: string): Buffer {
	const out = Buffer.alloc(text.length);
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i);
		if (code < 0x80 || (code >= 0xa0 && code <= 0xff)) {
			out[i] = code;
			continue;
		}
		const high = CP1252_HIGH.indexOf(code);
		out[i] = high >= 0 ? 0x80 + high : 0x3f; // '?'
	}
	return out;
}
