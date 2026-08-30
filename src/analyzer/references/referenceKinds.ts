// Read/write classification for identifier occurrences (issue #55).
//
// A reference's KIND is decided by the logical statement it sits in, read
// straight off the cached token stream: the assignment family writes its
// target's terminal name, declarations write the names they introduce, and
// everything else reads. The rules are syntactic on purpose - VBA spells
// every write it can prove as a statement shape - and the one gray zone,
// passing a variable to a ByRef parameter, stays a read: claiming a write
// there would need call-site signature resolution and would be wrong for
// every ByVal.
//
// `x = x + 1` is therefore TWO references with distinct kinds, which is the
// shape textDocument/references already produces; `Mid(s, 1, 2) = t` and
// `ReDim Preserve a(n)` are the honest readwrites (they modify in place).

import { tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';

export type ReferenceKind = 'read' | 'write' | 'readwrite';

const DECL_HEADS = new Set(['dim', 'static', 'const', 'private', 'public', 'global', 'friend']);
const SIGNATURE_HEADS = new Set(['sub', 'function', 'property', 'declare', 'enum', 'type', 'event']);
const PARAM_MODIFIERS = new Set(['byval', 'byref', 'optional', 'paramarray']);

function isName(t: VbaToken | undefined): boolean {
	return !!t && (t.kind === 'identifier' || t.kind === 'bracketedIdentifier');
}

function wordOf(t: VbaToken | undefined): string {
	if (!t) { return ''; }
	return (t.kind === 'keyword' ? (t.canonicalText ?? t.rawText) : t.rawText).toLowerCase();
}

/** Index just past a balanced paren group starting at `open`, else open+1. */
function pastParens(seg: readonly VbaToken[], open: number): number {
	let depth = 0;
	for (let i = open; i < seg.length; i++) {
		if (seg[i].rawText === '(') { depth++; }
		else if (seg[i].rawText === ')') {
			depth--;
			if (depth === 0) { return i + 1; }
		}
	}
	return open + 1;
}

/** First index of `raw` at paren depth zero in [from, seg.length), or -1. */
function depthZeroIndexOf(seg: readonly VbaToken[], raw: string, from: number): number {
	let depth = 0;
	for (let i = from; i < seg.length; i++) {
		const t = seg[i].rawText;
		if (t === '(') { depth++; }
		else if (t === ')') { depth--; }
		else if (depth === 0 && t === raw) { return i; }
	}
	return -1;
}

/**
 * The assignment-target rule: within `[from, eq)`, the WRITE lands on the
 * terminal name of the target chain - the name whose only suffix before the
 * `=` is a run of balanced paren groups. `x =`, `x(i) =`, `a.b.c =`,
 * `a(i).b =`, and With's `.y =` all select exactly one name; every other
 * name on the left (receivers, indexes) reads.
 */
function markAssignmentTarget(
	seg: readonly VbaToken[],
	from: number,
	eq: number,
	mark: (t: VbaToken, kind: ReferenceKind) => void,
): void {
	for (let k = from; k < eq; k++) {
		if (!isName(seg[k])) { continue; }
		let j = k + 1;
		while (j < eq && seg[j].rawText === '(') { j = pastParens(seg, j); }
		if (j === eq) { mark(seg[k], 'write'); }
	}
}

function classifySegment(
	seg: readonly VbaToken[],
	wanted: ReadonlySet<number>,
	out: Map<number, ReferenceKind>,
): void {
	if (seg.length === 0) { return; }
	const touches = seg.some((t) => wanted.has(t.start));
	if (!touches) { return; }
	const mark = (t: VbaToken, kind: ReferenceKind): void => {
		if (wanted.has(t.start)) { out.set(t.start, kind); }
	};
	let head = 0;
	let headWord = wordOf(seg[head]);

	// Set / Let are assignment statements with a keyword prefix.
	if (headWord === 'set' || headWord === 'let' || headWord === 'lset' || headWord === 'rset') {
		const eq = depthZeroIndexOf(seg, '=', head + 1);
		if (eq > 0) { markAssignmentTarget(seg, head + 1, eq, mark); }
		return;
	}

	if (headWord === 'for') {
		if (wordOf(seg[1]) === 'each') {
			if (isName(seg[2])) { mark(seg[2], 'write'); }
		} else if (isName(seg[1])) {
			mark(seg[1], 'write');
		}
		return; // the bounds and the collection read
	}

	if (headWord === 'redim') {
		let i = 1;
		let kind: ReferenceKind = 'write';
		if (wordOf(seg[i]) === 'preserve') { kind = 'readwrite'; i++; }
		let expectTarget = true;
		let depth = 0;
		for (; i < seg.length; i++) {
			const raw = seg[i].rawText;
			if (raw === '(') { depth++; continue; }
			if (raw === ')') { depth--; continue; }
			if (depth > 0) { continue; }
			if (raw === ',') { expectTarget = true; continue; }
			if (wordOf(seg[i]) === 'as') { expectTarget = false; continue; }
			if (expectTarget && isName(seg[i])) { mark(seg[i], kind); expectTarget = false; }
		}
		return;
	}

	if (headWord === 'erase') {
		let depth = 0;
		for (let i = 1; i < seg.length; i++) {
			const raw = seg[i].rawText;
			if (raw === '(') { depth++; }
			else if (raw === ')') { depth--; }
			else if (depth === 0 && isName(seg[i])) { mark(seg[i], 'write'); }
		}
		return;
	}

	// Mid/MidB statements modify their first argument in place.
	if ((headWord === 'mid' || headWord === 'mid$' || headWord === 'midb' || headWord === 'midb$')
		&& seg[1]?.rawText === '(') {
		const close = pastParens(seg, 1);
		if (seg[close]?.rawText === '=') {
			for (let i = 2; i < close; i++) {
				if (isName(seg[i])) { mark(seg[i], 'readwrite'); break; }
			}
		}
		return;
	}

	// Input #f, a, b / Line Input #f, s / Get #f, pos, var fill their
	// trailing variables.
	const isLineInput = headWord === 'line' && wordOf(seg[1]) === 'input';
	if ((headWord === 'input' || headWord === 'get' || isLineInput)
		&& seg.some((t) => t.rawText === '#')) {
		const commasNeeded = headWord === 'get' ? 2 : 1;
		let commas = 0;
		let depth = 0;
		for (let i = 1; i < seg.length; i++) {
			const raw = seg[i].rawText;
			if (raw === '(') { depth++; }
			else if (raw === ')') { depth--; }
			else if (depth === 0 && raw === ',') { commas++; }
			else if (depth === 0 && commas >= commasNeeded && isName(seg[i]) && seg[i - 1]?.rawText !== '.') {
				mark(seg[i], 'write');
			}
		}
		return;
	}

	if (DECL_HEADS.has(headWord)) {
		let i = head + 1;
		if (wordOf(seg[i]) === 'const' || headWord === 'const') {
			// Const A = 1, B = 2: names write, initializers read.
			if (headWord !== 'const') { i++; }
			let expectName = true;
			let depth = 0;
			for (; i < seg.length; i++) {
				const raw = seg[i].rawText;
				if (raw === '(') { depth++; }
				else if (raw === ')') { depth--; }
				else if (depth === 0 && raw === ',') { expectName = true; }
				else if (depth === 0 && raw === '=') { expectName = false; }
				else if (depth === 0 && expectName && isName(seg[i])) { mark(seg[i], 'write'); expectName = false; }
			}
			return;
		}
		if (SIGNATURE_HEADS.has(wordOf(seg[i])) || SIGNATURE_HEADS.has(headWord)) {
			classifySignature(seg, mark);
			return;
		}
		// Dim x As Long, y(10) As String, WithEvents app As Excel.Application
		let expectName = true;
		let depth = 0;
		for (; i < seg.length; i++) {
			const w = wordOf(seg[i]);
			const raw = seg[i].rawText;
			if (raw === '(') { depth++; continue; }
			if (raw === ')') { depth--; continue; }
			if (depth > 0) { continue; }
			if (raw === ',') { expectName = true; continue; }
			if (w === 'as' || w === 'new') { expectName = false; continue; }
			if (w === 'withevents') { continue; }
			if (expectName && isName(seg[i])) { mark(seg[i], 'write'); expectName = false; }
		}
		return;
	}

	if (SIGNATURE_HEADS.has(headWord)) {
		classifySignature(seg, mark);
		return;
	}

	// Inline If carries real statements after Then and Else on the same
	// logical line: `If x = 1 Then y = 2 Else z = 3` reads its condition and
	// writes both targets. Each tail classifies as its own statement.
	if (headWord === 'if' || headWord === 'elseif' || headWord === 'else') {
		const from = headWord === 'else' ? head + 1 : (() => {
			let depth = 0;
			for (let i = head + 1; i < seg.length; i++) {
				const raw = seg[i].rawText;
				if (raw === '(') { depth++; }
				else if (raw === ')') { depth--; }
				else if (depth === 0 && wordOf(seg[i]) === 'then') { return i + 1; }
			}
			return seg.length;
		})();
		if (from < seg.length) {
			let start = from;
			let depth = 0;
			for (let i = from; i <= seg.length; i++) {
				const atElse = i < seg.length && depth === 0 && wordOf(seg[i]) === 'else';
				if (i === seg.length || atElse) {
					if (i > start) { classifySegment(seg.slice(start, i), wanted, out); }
					start = i + 1;
					continue;
				}
				const raw = seg[i].rawText;
				if (raw === '(') { depth++; }
				else if (raw === ')') { depth--; }
			}
		}
		return;
	}

	// A plain assignment starts with an expression: a name, Me, or With's
	// leading dot. Anything keyword-led (If, While, Call, Debug...) reads.
	const startsExpression = isName(seg[head]) || wordOf(seg[head]) === 'me' || seg[head].rawText === '.';
	if (startsExpression) {
		const eq = depthZeroIndexOf(seg, '=', head);
		if (eq > 0 && seg[eq].kind === 'operator' && seg[eq].rawText === '=') {
			markAssignmentTarget(seg, head, eq, mark);
		}
	}
}

/** Sub/Function/Property/Event signatures: the name and its parameters. */
function classifySignature(
	seg: readonly VbaToken[],
	mark: (t: VbaToken, kind: ReferenceKind) => void,
): void {
	let i = 0;
	while (i < seg.length && !isName(seg[i])) {
		if (seg[i].rawText === '(') { return; }
		i++;
	}
	// Skip modifier keywords spelled as identifiers never happens: names
	// after the head keywords are the procedure name (Property carries
	// Get/Let/Set first, which are keywords).
	if (isName(seg[i])) { mark(seg[i], 'write'); i++; }
	if (seg[i]?.rawText !== '(') { return; }
	let depth = 0;
	let expectParam = true;
	for (; i < seg.length; i++) {
		const raw = seg[i].rawText;
		const w = wordOf(seg[i]);
		if (raw === '(') { depth++; if (depth === 1) { expectParam = true; } continue; }
		if (raw === ')') { depth--; if (depth === 0) { break; } continue; }
		if (depth !== 1) { continue; }
		if (raw === ',') { expectParam = true; continue; }
		if (w === 'as' || w === 'new') { expectParam = false; continue; }
		if (PARAM_MODIFIERS.has(w)) { continue; }
		if (raw === '=') { expectParam = false; continue; } // Optional defaults read
		if (expectParam && isName(seg[i])) { mark(seg[i], 'write'); expectParam = false; }
	}
}

/**
 * Classifies the identifier occurrences whose ABSOLUTE offsets are given.
 * Unmatched offsets come back 'read' - the honest default for every
 * position that is not provably a write.
 */
export function classifyReferenceKinds(
	source: string,
	offsets: readonly number[],
): Map<number, ReferenceKind> {
	const out = new Map<number, ReferenceKind>();
	if (offsets.length === 0) { return out; }
	const wanted = new Set(offsets);
	const all = tokenizeCached(source);
	let seg: VbaToken[] = [];
	const flush = (): void => {
		if (seg.length > 0) { classifySegment(seg, wanted, out); seg = []; }
	};
	for (const t of all) {
		if (t.kind === 'newline' || t.kind === 'colon') { flush(); continue; }
		if (t.kind === 'comment') { continue; }
		seg.push(t);
	}
	flush();
	for (const offset of offsets) {
		if (!out.has(offset)) { out.set(offset, 'read'); }
	}
	return out;
}
