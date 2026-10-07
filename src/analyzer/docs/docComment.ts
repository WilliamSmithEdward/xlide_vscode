// Inline XML documentation-comment parser.
//
// A documentation comment is a run of contiguous lines whose first non-space
// characters are `'''` (three apostrophes), sitting directly above a procedure,
// type, enum, Declare, or module-level variable declaration - mirroring the
// Visual Studio C# `///` convention. xlide's own directive comments
// (`' @xlide-analysis-*` suppressions, `' @xlide-test*` markers) are
// transparent: the block attaches through them in any stacking order. Object-module docs use the same block above
// the first Option directive because VBA has no source-level class declaration.
// The body is a fragment of XML using the tag vocabulary in docModel.ts
// (<summary>, <param>, <returns>, <remarks>, <example>). Parsing is
// intentionally lenient (regex-based, no XML dependency) so a partially written
// or slightly malformed block still yields useful text.
//
// Pure analyzer code: no `vscode` dependency. See user_guides/vba-doc-comments.md.

import { VbaDoc, VbaDocParam, VbaDocSource } from './docModel';
import type { Span } from '../parser/nodes';
import { lineStartAtAnyBreak, lineEndAtOrAfter } from '../../vbaSourceScan';

/** Decodes the five predefined XML entities. */
function decodeEntities(text: string): string {
	return text
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&');
}

/** Trims and collapses internal whitespace runs (including newlines) to a space. */
function collapse(text: string): string {
	return decodeEntities(text).replace(/\s+/g, ' ').trim();
}

/** Trims surrounding blank lines but preserves internal layout (for examples). */
function dedent(text: string): string {
	const decoded = decodeEntities(text).replace(/\r\n/g, '\n');
	return decoded.replace(/^\n+/, '').replace(/\s+$/, '');
}

/** Extracts the inner text of the first `<tag>...</tag>` in `body`, if present. */
function firstTag(body: string, tag: string): string | undefined {
	const m = firstTagMatch(body, tag);
	return m?.body;
}

interface DocTagMatch {
	attrs: string;
	body: string;
}

/** The same lenient, non-overlapping matches as the paired/self-closing regex. */
function* docTagMatches(body: string, tag: string): Generator<DocTagMatch, void> {
	const opening = new RegExp(`<${tag}\\b`, 'gi');
	let closing: RegExp | undefined;
	let closingMissing = false;
	while (opening.exec(body) !== null) {
		const attrsStart = opening.lastIndex;
		const angle = body.indexOf('>', attrsStart);
		if (angle < 0) {
			return;
		}
		// If this opening cannot match a pair, nested openings in its
		// attributes share its ending and cannot match a pair either.
		opening.lastIndex = angle + 1;
		if (body[angle - 1] === '/') {
			yield { attrs: body.slice(attrsStart, angle - 1), body: '' };
			continue;
		}
		// Once a close is missing it stays missing. If there is one, consume
		// the paired body before finding another opening, as the old regex did.
		if (!closingMissing) {
			closing ??= new RegExp(`</${tag}>`, 'gi');
			closing.lastIndex = angle + 1;
			const close = closing.exec(body);
			if (close) {
				yield { attrs: body.slice(attrsStart, angle), body: body.slice(angle + 1, close.index) };
				opening.lastIndex = closing.lastIndex;
			} else {
				closingMissing = true;
			}
		}
	}
}

function firstTagMatch(body: string, tag: string): DocTagMatch | undefined {
	const match = docTagMatches(body, tag).next().value;
	if (!match) {
		return undefined;
	}
	// Preserve firstTagMatch's extra slash trimming; params use raw attributes.
	return { attrs: match.attrs.replace(/\/\s*$/, ''), body: match.body };
}

interface DocAttribute {
	name: string;
	value: string;
	/** Raw value coordinates relative to the attribute string. */
	start: number;
	length: number;
}

/** Advances over attribute names without retrying every suffix of an unknown word. */
function* docAttributes(raw: string): Generator<DocAttribute, void> {
	const names = /[A-Za-z_][A-Za-z0-9_-]*/g;
	const valueOpening = /\s*=\s*"/y;
	let name: RegExpExecArray | null;
	while ((name = names.exec(raw)) !== null) {
		valueOpening.lastIndex = names.lastIndex;
		if (!valueOpening.exec(raw)) {
			continue;
		}
		const start = valueOpening.lastIndex;
		const end = raw.indexOf('"', start);
		if (end < 0) {
			return;
		}
		names.lastIndex = end + 1;
		yield {
			name: name[0].toLowerCase(),
			value: decodeEntities(raw.slice(start, end)).trim(),
			start,
			length: end - start,
		};
	}
}

function attrsOf(raw: string): Map<string, string> {
	const out = new Map<string, string>();
	for (const attr of docAttributes(raw)) {
		out.set(attr.name, attr.value);
	}
	return out;
}

/** Extracts every `<param name="...">...</param>` entry, in document order. */
function extractParams(body: string): VbaDocParam[] {
	const out: VbaDocParam[] = [];
	for (const match of docTagMatches(body, 'param')) {
		const attrs = attrsOf(match.attrs);
		const name = attrs.get('name') ?? '';
		if (name) {
			const param: VbaDocParam = { name, text: collapse(match.body) };
			const type = attrs.get('type');
			const unit = attrs.get('unit');
			const value = attrs.get('value');
			if (type) {
				param.type = type;
			}
			if (unit) {
				param.unit = unit;
			}
			if (value) {
				param.value = value;
			}
			out.push(param);
		}
	}
	return out;
}

const HAS_TAG_RE = /<(summary|param|returns|remarks|example|signature)\b/i;

/**
 * Parses an XML documentation body (the inner text shared by inline comments and
 * external `<member>` entries) into a {@link VbaDoc}. When the body contains no
 * recognised tags, the whole trimmed text becomes the summary so a plain-text
 * note still produces a tooltip.
 */
export function parseDocBody(body: string, source: VbaDocSource): VbaDoc {
	const doc: VbaDoc = { params: [], source };
	if (!HAS_TAG_RE.test(body)) {
		const plain = collapse(body);
		if (plain) {
			doc.summary = plain;
		}
		return doc;
	}
	const summary = firstTag(body, 'summary');
	if (summary !== undefined) {
		doc.summary = collapse(summary);
	}
	doc.params = extractParams(body);
	const returns = firstTagMatch(body, 'returns');
	if (returns !== undefined) {
		doc.returns = collapse(returns.body);
		const attrs = attrsOf(returns.attrs);
		const type = attrs.get('type');
		const unit = attrs.get('unit');
		const value = attrs.get('value');
		if (type) {
			doc.returnsType = type;
		}
		if (unit) {
			doc.returnsUnit = unit;
		}
		if (value) {
			doc.returnsValue = value;
		}
	}
	const remarks = firstTag(body, 'remarks');
	if (remarks !== undefined) {
		doc.remarks = collapse(remarks);
	}
	const example = firstTag(body, 'example');
	if (example !== undefined) {
		doc.example = dedent(example);
	}
	const signature = firstTag(body, 'signature');
	if (signature !== undefined) {
		doc.signature = collapse(signature);
	}
	return doc;
}

function hasAnyDocContent(doc: VbaDoc): boolean {
	return !!doc.summary ||
		doc.params.length > 0 ||
		!!doc.returns ||
		!!doc.remarks ||
		!!doc.example ||
		!!doc.signature;
}

function docFromLines(docLines: readonly string[]): VbaDoc | undefined {
	if (docLines.length === 0) {
		return undefined;
	}
	const doc = parseDocBody(docLines.join('\n'), 'inline');
	return hasAnyDocContent(doc) ? doc : undefined;
}

function stripDocPrefix(trimmed: string): string {
	let rest = trimmed.slice(3);
	if (rest.startsWith(' ')) {
		rest = rest.slice(1);
	}
	return rest;
}

interface SourceLine {
	text: string;
	start: number;
	end: number;
}

function* sourceLines(source: string): Generator<SourceLine, void> {
	let start = 0;
	for (;;) {
		const end = lineEndAtOrAfter(source, start);
		const next = end + (source[end] === '\r' && source[end + 1] === '\n' ? 2 : 1);
		yield { text: source.slice(start, end), start, end: end === source.length ? end : next - 1 };
		if (end === source.length) { return; }
		start = next;
	}
}

/** Previous physical lines, excluding each CR, CRLF or LF terminator. */
function* precedingLines(source: string, offset: number): Generator<SourceLine, void> {
	let start = lineStartAtAnyBreak(source, offset);
	while (start > 0) {
		let end = start - 1;
		if (source[end] === '\n' && source[end - 1] === '\r') { end--; }
		start = lineStartAtAnyBreak(source, end);
		yield { start, end, text: source.slice(start, end) };
	}
}

function isOrdinaryComment(trimmed: string): boolean {
	return trimmed.startsWith("'") && !trimmed.startsWith("'''");
}

/**
 * xlide's own directive comments - analysis suppressions and test markers -
 * are transparent to the doc scans: a `'''` block attaches to its member
 * through them, whatever order the three comment grammars stack in. Any
 * OTHER intervening line still detaches the block, as documented.
 */
function isXlideDirectiveComment(trimmed: string): boolean {
	return isOrdinaryComment(trimmed) && /^'+\s*@xlide-\S/i.test(trimmed);
}

function isModuleHeaderBoundary(trimmed: string): boolean {
	// A directive is NOT enough separation: the member scan reads through
	// directives, so a block a directive "separates" belongs to the member
	// below - letting the header also claim it would attach it twice.
	return trimmed === '' ||
		(isOrdinaryComment(trimmed) && !isXlideDirectiveComment(trimmed)) ||
		/^Option\b/i.test(trimmed);
}

/**
 * Extracts a module-header documentation block from the top of a module. This
 * supports both object modules with a block directly above `Option Explicit`
 * and standard modules that have no `Option` directive by requiring the header
 * block to be visually separated from the first declaration by a blank line or
 * ordinary comment. A block immediately above a declaration remains declaration
 * documentation.
 */
export function extractModuleHeaderDoc(
	source: string,
	startOffset = 0,
): VbaDoc | undefined {
	const safeOffset = Math.max(0, startOffset);
	if (Number.isNaN(safeOffset) || safeOffset > source.length) { return undefined; }
	const docLines: string[] = [];
	// Read only the header. Most symbol builds need the first few lines, not
	// a line object and substring for every procedure in the whole class.
	for (const line of sourceLines(source)) {
		if (!(line.start >= safeOffset)) { continue; }
		const trimmed = line.text.trimStart();
		if (docLines.length === 0) {
			if (trimmed === '' || isOrdinaryComment(trimmed)) { continue; }
			if (!trimmed.startsWith("'''")) { return undefined; }
		} else if (!trimmed.startsWith("'''")) {
			return isModuleHeaderBoundary(trimmed) ? docFromLines(docLines) : undefined;
		}
		docLines.push(stripDocPrefix(trimmed));
	}
	return docFromLines(docLines);
}

/** One `'''` line of the block above a declaration. */
export interface DocBlockLine {
	/** Offset of the line's first character. */
	start: number;
	/**
	 * Offset of the first of the directive lines directly above it, or `start`
	 * when there are none. A line put before this one goes here, so it does
	 * not come between a directive and the line the directive is about.
	 */
	directivesStart: number;
	/** Offset of the text after `'''` and the one space the parser drops. */
	textStart: number;
	/** That text, to the end of the line. */
	text: string;
}

/**
 * The `'''` lines directly above a declaration, top to bottom. xlide's own
 * directive lines between them are passed over and left out.
 *
 * @param source Full module source text.
 * @param declStart Offset of the first character of the declaration (its full
 *   span start, e.g. the `Public`/`Sub` keyword).
 */
export function leadingDocLines(
	source: string,
	declStart: number,
): DocBlockLine[] {
	// Walk physical lines backward from the declaration's own line. This runs
	// once per declaration while building module symbols, so slicing and
	// splitting the whole module prefix here (the obvious implementation) makes
	// symbol building quadratic in module size.
	const lines: DocBlockLine[] = [];
	for (const { start: prevStart, text: line } of precedingLines(source, declStart)) {
		const trimmed = line.trimStart();
		if (isXlideDirectiveComment(trimmed)) {
			// Suppression and test directives are the product's own grammar;
			// the block attaches through them in any stacking order.
			const below = lines[lines.length - 1];
			if (below) {
				below.directivesStart = prevStart;
			}
			continue;
		}
		if (!trimmed.startsWith("'''")) {
			break;
		}
		const text = stripDocPrefix(trimmed);
		lines.push({
			start: prevStart,
			directivesStart: prevStart,
			textStart: prevStart + line.length - text.length,
			text,
		});
	}
	return lines.reverse();
}

/**
 * Where the lines XLIDE reads as part of a declaration start: its `'''` doc
 * comment and xlide's directive lines, directly above it in any order. The
 * declaration's own line start when it has none. Code that moves or deletes
 * the declaration takes these with it; left behind, they would belong to
 * whatever declaration came next.
 */
export function attachedCommentsStart(source: string, declStart: number): number {
	let start = lineStartAtAnyBreak(source, declStart);
	for (const line of precedingLines(source, declStart)) {
		const trimmed = line.text.trimStart();
		if (!trimmed.startsWith("'''") && !isXlideDirectiveComment(trimmed)) { break; }
		start = line.start;
	}
	return start;
}

/**
 * Scans upward from the start of a declaration to collect a contiguous run of
 * `'''` documentation-comment lines, and parses them into a {@link VbaDoc}.
 * Returns undefined when no such comment immediately precedes the declaration.
 *
 * @param source Full module source text.
 * @param declStart Offset of the first character of the declaration (its full
 *   span start, e.g. the `Public`/`Sub` keyword).
 */
export function extractLeadingDoc(
	source: string,
	declStart: number,
): VbaDoc | undefined {
	return docFromLines(leadingDocLines(source, declStart).map((line) => line.text));
}

/** Where the `<param>` tags above a declaration name `paramName`: their `name` values. */
export function docParamNameSpans(
	source: string,
	declStart: number,
	paramName: string,
): Span[] {
	const lower = paramName.toLowerCase();
	const spans: Span[] = [];
	for (const tag of scanDocTags(leadingDocLines(source, declStart)) ?? []) {
		if (tag.tag === 'param' && tag.nameSpan && tag.name?.toLowerCase() === lower) {
			spans.push(tag.nameSpan);
		}
	}
	return spans;
}

/** A vocabulary tag in a `'''` block, and where it sits in the module. */
export interface DocTagOccurrence {
	/** The tag in lower case: summary, param, returns, remarks, example or signature. */
	tag: string;
	/** From the `<` to the `>` of the opening tag. */
	open: Span;
	/** The `name` attribute, decoded and trimmed, when the tag has a non-empty one. */
	name?: string;
	/** The text between the quotes of that `name` attribute. */
	nameSpan?: Span;
	/** True when a `type`, `unit` or `value` attribute says something. */
	hasHints: boolean;
	/** The tag's text as a hover shows it; undefined when the tag is never closed. */
	text?: string;
	/** Offset just past the closing tag, or the `/>`; undefined when never closed. */
	end?: number;
}

/**
 * The vocabulary tags of a `'''` block in document order, or undefined when
 * it has none: a block of plain text is a note, which the parser reads as a
 * summary. A tag counts as closed the way the parser reads it, by its own
 * closing tag before the next tag of the same name opens.
 */
export function scanDocTags(lines: readonly DocBlockLine[]): DocTagOccurrence[] | undefined {
	const body = lines.map((line) => line.text).join('\n');
	if (!HAS_TAG_RE.test(body)) {
		return undefined;
	}
	const bodyStarts: number[] = [];
	let at = 0;
	for (const line of lines) {
		bodyStarts.push(at);
		at += line.text.length + 1;
	}
	const toSource = (offset: number): number => {
		// Tags and their closing/name spans need not arrive in offset order.
		// Locate the last line starting at or before this offset without
		// rescanning the whole block for every span.
		let low = 0;
		let high = bodyStarts.length;
		while (low < high) {
			const mid = low + Math.floor((high - low) / 2);
			if (bodyStarts[mid] <= offset) {
				low = mid + 1;
			} else {
				high = mid;
			}
		}
		const i = Math.max(0, low - 1);
		return lines[i].textStart + (offset - bodyStarts[i]);
	};
	const lower = body.toLowerCase();
	const closingOffsets = new Map<string, number>();
	const tags: DocTagOccurrence[] = [];
	const opening = new RegExp(HAS_TAG_RE.source, 'gi');
	let m: RegExpExecArray | null;
	while ((m = opening.exec(body)) !== null) {
		const tag = m[1].toLowerCase();
		const angle = body.indexOf('>', opening.lastIndex);
		if (angle < 0) {
			break;
		}
		const openEnd = angle + 1;
		const selfClosing = body[angle - 1] === '/';
		opening.lastIndex = openEnd;
		const occurrence: DocTagOccurrence = {
			tag,
			open: { start: toSource(m.index), end: toSource(openEnd) },
			hasHints: false,
		};
		// Read the attributes as attrsOf does: the last of a repeated name wins.
		const attrs = new Map<string, { value: string; start: number; length: number }>();
		const attrsStart = m.index + 1 + tag.length;
		const rawAttrs = body.slice(attrsStart, selfClosing ? angle - 1 : angle);
		for (const attr of docAttributes(rawAttrs)) {
			attrs.set(attr.name, {
				value: attr.value,
				start: attrsStart + attr.start,
				length: attr.length,
			});
		}
		const name = attrs.get('name');
		if (name?.value) {
			occurrence.name = name.value;
			occurrence.nameSpan = { start: toSource(name.start), end: toSource(name.start + name.length) };
		}
		occurrence.hasHints = ['type', 'unit', 'value'].some((key) => !!attrs.get(key)?.value);
		if (selfClosing) {
			occurrence.text = '';
			occurrence.end = occurrence.open.end;
		} else {
			// Opening tags are visited in order. A cached following close stays
			// valid until we pass it; a missing close never needs another scan.
			let close = closingOffsets.get(tag);
			if (close === undefined || (close >= 0 && close < openEnd)) {
				close = lower.indexOf(`</${tag}>`, openEnd);
				closingOffsets.set(tag, close);
			}
			if (close >= 0) {
				const reopen = new RegExp(`<${tag}\\b`, 'g');
				reopen.lastIndex = openEnd;
				const next = reopen.exec(lower);
				if (!next || close < next.index) {
					occurrence.text = collapse(body.slice(openEnd, close));
					occurrence.end = toSource(close + tag.length + 3);
				}
			}
		}
		tags.push(occurrence);
	}
	return tags;
}
