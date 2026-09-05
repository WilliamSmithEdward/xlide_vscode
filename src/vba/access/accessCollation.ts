import { AccessFormatError } from './accessFormat';
import {
	ATTACHED_WEIGHTS,
	COMBINING_RUNS,
	EXPANSIONS,
	KANA_RANGES,
	KANA_SMALL,
	PRIMARY_RUNS,
	UNPRINTABLE,
	WEIGHTS,
} from './accessCollationTable';

/**
 * Text sort keys: the bytes the ACE engine stores for a text value in an
 * index, for sort order 1033 version 0 (the "General" order of Jet 4 and
 * Access 2007 files).
 *
 * The engine's rules, all measured on keys it wrote:
 *
 * - A string is trimmed of trailing spaces and composed: a base letter plus a
 *   combining mark keys like the precomposed letter.
 * - Each character maps to zero or more elements. Most letters are one element
 *   of one or two bytes; expansions such as sharp s (`ss`) or the `ffi`
 *   ligature are several; 19,585 code points are ignored. Case is not encoded.
 * - The key is the element bytes, 0x01, then up to four sections separated by
 *   0x01 with trailing empty ones omitted, then 0x00:
 *   1. one diacritic weight per element, 0x02 standing in for elements without
 *      one, trailing stand-ins trimmed;
 *   2. never seen non-empty;
 *   3. for kana, a bit stream followed by a fixed suffix;
 *   4. for each ignorable-but-recorded character (hyphen, apostrophe,
 *      controls), four bytes naming where it sat.
 *
 * Ported from pyOpenVBA's `_collation.py`; the table is generated, see
 * `accessCollationTable.ts`.
 */

const TEXT_END = 0x01;
const SECTION_END = 0x01;
const KEY_END = 0x00;
const WEIGHT_PLACEHOLDER = 0x02;
const KANA_SUFFIX = [0xff, 0x02, 0x80, 0xff, 0x80];
const KANA_BEFORE_UNPRINTABLE = 0xff;
const UNPRINTABLE_LEAD = 0x80;
const UNPRINTABLE_BASE = 7;
const UNPRINTABLE_STEP = 4;
const UNPRINTABLE_MID = 0x06;
/**
 * The engine stores at most 510 key bytes including the flag byte and cuts
 * longer keys without a clean terminator, which has not been reproduced.
 */
const MAX_KEY_LENGTH = 509;

const SPACE = 0x20;

/** What one character contributes to a key. */
interface CharacterKey {
	/** Each element as its primary bytes and its diacritic weight, 0 for none. */
	elements: Array<{ primary: number[]; weight: number }>;
	kana: boolean;
	smallKana: boolean;
	/** The code recorded for an ignorable character, or undefined. */
	unprintable?: number;
}

const EMPTY: CharacterKey = { elements: [], kana: false, smallKana: false };

/** Whether the code point has a non-zero canonical combining class. */
export function isCombining(codePoint: number): boolean {
	for (let i = 0; i < COMBINING_RUNS.length; i += 2) {
		const start = COMBINING_RUNS[i];
		if (codePoint < start) {
			return false;
		}
		if (codePoint < start + COMBINING_RUNS[i + 1]) {
			return true;
		}
	}
	return false;
}

/** The run covering a code point, by binary search over the run starts. */
function runPrimary(codePoint: number): number[][] {
	let low = 0;
	let high = PRIMARY_RUNS.length / 4 - 1;
	let found = -1;
	while (low <= high) {
		const middle = (low + high) >> 1;
		if (PRIMARY_RUNS[middle * 4] <= codePoint) {
			found = middle;
			low = middle + 1;
		} else {
			high = middle - 1;
		}
	}
	if (found < 0) {
		return [];
	}
	const first = PRIMARY_RUNS[found * 4];
	const count = PRIMARY_RUNS[found * 4 + 1];
	const firstKey = PRIMARY_RUNS[found * 4 + 2];
	const width = PRIMARY_RUNS[found * 4 + 3];
	if (codePoint >= first + count) {
		return [];
	}
	const key = firstKey + (codePoint - first);
	const bytes: number[] = [];
	for (let i = width - 1; i >= 0; i -= 1) {
		bytes.push((key >> (8 * i)) & 0xff);
	}
	return [bytes];
}

function inKanaRange(codePoint: number): boolean {
	for (let i = 0; i < KANA_RANGES.length; i += 2) {
		if (codePoint >= KANA_RANGES[i] && codePoint <= KANA_RANGES[i + 1]) {
			return true;
		}
	}
	return false;
}

function characterKey(codePoint: number): CharacterKey {
	if (codePoint === SPACE) {
		// Trailing spaces are trimmed before encoding; any other space is an
		// element of its own. A lone space keys as empty, so the generated
		// table cannot carry this and the encoder does.
		return { elements: [{ primary: [0x07], weight: 0 }], kana: false, smallKana: false };
	}
	let primaries = EXPANSIONS.get(codePoint)?.map((part) => [...part]) ?? runPrimary(codePoint);
	const weights = WEIGHTS.get(codePoint) ?? [];
	if (primaries.length === 0 && weights.length > 0) {
		// A bare combining mark: a weight with nothing under it.
		primaries = [[]];
	}
	const elements = primaries.map((primary, index) => ({
		primary,
		weight: index < weights.length ? weights[index] : 0,
	}));
	const kana = inKanaRange(codePoint);
	const unprintable = UNPRINTABLE.get(codePoint);
	if (elements.length === 0 && !kana && unprintable === undefined) {
		return EMPTY;
	}
	return {
		elements,
		kana,
		smallKana: KANA_SMALL.has(codePoint),
		...(unprintable === undefined ? {} : { unprintable }),
	};
}

/**
 * Group each character with the combining marks that follow it, folding a
 * mark into the base when a precomposed character exists, which is what the
 * engine does. Whole-string NFC would be wrong: it also rewrites singletons
 * such as U+0387 to U+00B7, and the engine keys those as themselves.
 */
function composeMarks(text: string): Array<{ base: string; marks: string[] }> {
	const out: Array<{ base: string; marks: string[] }> = [];
	for (const char of text) {
		const last = out[out.length - 1];
		if (last && isCombining(char.codePointAt(0)!)) {
			const composed = (last.base + char).normalize('NFC');
			if ([...composed].length === 1 && last.marks.length === 0) {
				last.base = composed;
			} else {
				last.marks.push(char);
			}
			continue;
		}
		out.push({ base: char, marks: [] });
	}
	return out;
}

/**
 * The weight a combining mark adds to the element before it. The first mark
 * on a base with no precomposed form takes the weight that mark gives any
 * precomposed letter; a further mark adds the weight it carries alone.
 */
function markWeight(mark: string, first: boolean): number {
	const codePoint = mark.codePointAt(0)!;
	if (first) {
		const attached = ATTACHED_WEIGHTS.get(codePoint);
		if (attached !== undefined) {
			return attached;
		}
	}
	const weights = WEIGHTS.get(codePoint);
	return weights && weights.length > 0 ? weights[0] : 0;
}

/** Three two-bit kana codes per byte behind a `10` marker, zero-padded. */
function packKana(pairs: number[]): number[] {
	const out: number[] = [];
	for (let i = 0; i < pairs.length; i += 3) {
		const chunk = [pairs[i] ?? 0, pairs[i + 1] ?? 0, pairs[i + 2] ?? 0];
		out.push((0b10 << 6) | (chunk[0] << 4) | (chunk[1] << 2) | chunk[2]);
	}
	return out;
}

/** The stored key for text in an ascending index column, without the flag byte. */
export function encodeTextKey(text: string): Buffer {
	const primary: number[] = [];
	const weights: number[] = [];
	const kanaPairs: number[] = [];
	let lastSmall = -1;
	const unprintables: number[] = [];
	let elementCount = 0;

	for (const { base, marks } of composeMarks(text.replace(/ +$/, ''))) {
		const codePoint = base.codePointAt(0)!;
		if (codePoint > 0xffff) {
			throw new AccessFormatError(
				`U+${codePoint.toString(16).toUpperCase()}: characters outside the Basic `
				+ 'Multilingual Plane have not been measured in an index key.',
			);
		}
		const info = characterKey(codePoint);
		for (const element of info.elements) {
			primary.push(...element.primary);
			weights.push(element.weight);
			elementCount += 1;
		}
		if (marks.length > 0 && info.elements.length > 0) {
			const composedAlready = base.normalize('NFD') !== base;
			marks.forEach((mark, index) => {
				weights[weights.length - 1] += markWeight(mark, index === 0 && !composedAlready);
			});
		} else if (marks.length > 0) {
			for (const mark of marks) {
				for (const element of characterKey(mark.codePointAt(0)!).elements) {
					primary.push(...element.primary);
					weights.push(element.weight);
					elementCount += 1;
				}
			}
		}
		if (info.kana) {
			kanaPairs.push(info.smallKana ? 0b10 : 0b11);
			if (info.smallKana) {
				lastSmall = kanaPairs.length - 1;
			}
		}
		if (info.unprintable !== undefined) {
			unprintables.push(
				UNPRINTABLE_LEAD,
				UNPRINTABLE_BASE + UNPRINTABLE_STEP * elementCount,
				UNPRINTABLE_MID,
				info.unprintable,
			);
		}
	}

	const section1 = weights.map((weight) => weight || WEIGHT_PLACEHOLDER);
	if (section1.some((weight) => weight > 0xff)) {
		// Marks stacked deeply enough to carry the weight past a byte: what the
		// engine stores there has not been measured, and a byte that wrapped
		// would sort somewhere arbitrary.
		throw new AccessFormatError(
			'A diacritic weight in this text exceeds one byte, which no measured key '
			+ 'covers.',
		);
	}
	while (section1.length > 0 && section1[section1.length - 1] === WEIGHT_PLACEHOLDER) {
		section1.pop();
	}
	const section3: number[] = [];
	if (kanaPairs.length > 0) {
		if (lastSmall >= 0) {
			section3.push(...packKana(kanaPairs.slice(0, lastSmall + 1)));
		}
		section3.push(...KANA_SUFFIX);
		if (unprintables.length > 0) {
			section3.push(KANA_BEFORE_UNPRINTABLE);
		}
	}

	const key = [...primary, TEXT_END, ...section1];
	if (section3.length > 0 || unprintables.length > 0) {
		key.push(SECTION_END, SECTION_END); // section 2 is always empty
		key.push(...section3);
	}
	if (unprintables.length > 0) {
		key.push(SECTION_END, ...unprintables);
	}
	key.push(KEY_END);
	if (key.length > MAX_KEY_LENGTH) {
		throw new AccessFormatError(
			`A text key of ${key.length} bytes exceeds the engine's ${MAX_KEY_LENGTH}-byte limit; `
			+ 'the engine truncates such keys in a way that has not been reproduced.',
		);
	}
	return Buffer.from(key);
}
