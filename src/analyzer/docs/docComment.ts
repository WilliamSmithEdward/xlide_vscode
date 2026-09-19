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
import { lineStartAt } from '../../vbaSourceScan';

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

function firstTagMatch(body: string, tag: string): { attrs: string; body: string } | undefined {
	// Accept both the paired form `<tag ...>...</tag>` and the self-closing form
	// `<tag .../>` (empty body) so e.g. `<returns type="Long"/>` is not dropped.
	const re = new RegExp(`<${tag}\\b([^>]*?)(?:/>|>([\\s\\S]*?)<\\/${tag}>)`, 'i');
	const m = re.exec(body);
	if (!m) {
		return undefined;
	}
	// Strip a trailing slash left on the attribute text by a self-closing tag.
	const attrs = (m[1] ?? '').replace(/\/\s*$/, '');
	return { attrs, body: m[2] ?? '' };
}

const ATTRIBUTE_RE = /([A-Za-z_][A-Za-z0-9_-]*)\s*=\s*"([^"]*)"/;

function attrsOf(raw: string): Map<string, string> {
	const out = new Map<string, string>();
	const re = new RegExp(ATTRIBUTE_RE.source, 'g');
	let m: RegExpExecArray | null;
	while ((m = re.exec(raw)) !== null) {
		out.set(m[1].toLowerCase(), decodeEntities(m[2]).trim());
	}
	return out;
}

/** Extracts every `<param name="...">...</param>` entry, in document order. */
function extractParams(body: string): VbaDocParam[] {
	const out: VbaDocParam[] = [];
	// Accept both `<param ...>text</param>` and self-closing `<param .../>`
	// (treated as empty text) so self-closing params are not silently dropped.
	const re = /<param\b([^>]*?)(?:\/>|>([\s\S]*?)<\/param>)/gi;
	let m: RegExpExecArray | null;
	while ((m = re.exec(body)) !== null) {
		const attrs = attrsOf(m[1]);
		const name = attrs.get('name') ?? '';
		if (name) {
			const param: VbaDocParam = { name, text: collapse(m[2] ?? '') };
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

function sourceLines(source: string): SourceLine[] {
	const rawLines = source.split('\n');
	const out: SourceLine[] = [];
	let offset = 0;
	for (const raw of rawLines) {
		const text = raw.replace(/\r$/, '');
		out.push({
			text,
			start: offset,
			end: offset + raw.length,
		});
		offset += raw.length + 1;
	}
	return out;
}

function firstLineIndexAtOrAfter(lines: readonly SourceLine[], offset: number): number {
	const safeOffset = Math.max(0, offset);
	for (let i = 0; i < lines.length; i += 1) {
		const line = lines[i];
		if (line.start >= safeOffset) {
			return i;
		}
		if (safeOffset <= line.end) {
			return i + 1;
		}
	}
	return lines.length;
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
	const lines = sourceLines(source);
	let i = firstLineIndexAtOrAfter(lines, startOffset);
	while (i < lines.length) {
		const trimmed = lines[i].text.trimStart();
		if (trimmed === '' || isOrdinaryComment(trimmed)) {
			i += 1;
			continue;
		}
		break;
	}
	if (i >= lines.length || !lines[i].text.trimStart().startsWith("'''")) {
		return undefined;
	}

	const docLines: string[] = [];
	while (i < lines.length) {
		const trimmed = lines[i].text.trimStart();
		if (!trimmed.startsWith("'''")) {
			break;
		}
		docLines.push(stripDocPrefix(trimmed));
		i += 1;
	}

	if (i < lines.length) {
		const next = lines[i].text.trimStart();
		if (!isModuleHeaderBoundary(next)) {
			return undefined;
		}
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
	let lineStart = lineStartAt(source, declStart);
	const lines: DocBlockLine[] = [];
	while (lineStart > 0) {
		const prevEnd = lineStart - 1; // the '\n' terminating the previous line
		const prevStart = lineStartAt(source, prevEnd);
		const line = source.slice(prevStart, prevEnd).replace(/\r$/, '');
		const trimmed = line.trimStart();
		if (isXlideDirectiveComment(trimmed)) {
			// Suppression and test directives are the product's own grammar;
			// the block attaches through them in any stacking order.
			const below = lines[lines.length - 1];
			if (below) {
				below.directivesStart = prevStart;
			}
			lineStart = prevStart;
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
		lineStart = prevStart;
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
	let lineStart = lineStartAt(source, declStart);
	while (lineStart > 0) {
		const prevStart = lineStartAt(source, lineStart - 1);
		const trimmed = source.slice(prevStart, lineStart - 1).replace(/\r$/, '').trimStart();
		if (!trimmed.startsWith("'''") && !isXlideDirectiveComment(trimmed)) {
			break;
		}
		lineStart = prevStart;
	}
	return lineStart;
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

const OPENING_TAG_RE = /<(summary|param|returns|remarks|example|signature)\b([^>]*?)(\/?)>/;

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
		let i = bodyStarts.length - 1;
		while (i > 0 && bodyStarts[i] > offset) {
			i -= 1;
		}
		return lines[i].textStart + (offset - bodyStarts[i]);
	};
	const lower = body.toLowerCase();
	const tags: DocTagOccurrence[] = [];
	const opening = new RegExp(OPENING_TAG_RE.source, 'gi');
	let m: RegExpExecArray | null;
	while ((m = opening.exec(body)) !== null) {
		const tag = m[1].toLowerCase();
		const openEnd = m.index + m[0].length;
		const occurrence: DocTagOccurrence = {
			tag,
			open: { start: toSource(m.index), end: toSource(openEnd) },
			hasHints: false,
		};
		// Read the attributes as attrsOf does: the last of a repeated name wins.
		const attrs = new Map<string, { value: string; start: number; length: number }>();
		const attrsStart = m.index + 1 + tag.length;
		const attrRe = new RegExp(ATTRIBUTE_RE.source, 'g');
		let attr: RegExpExecArray | null;
		while ((attr = attrRe.exec(m[2])) !== null) {
			attrs.set(attr[1].toLowerCase(), {
				value: decodeEntities(attr[2]).trim(),
				start: attrsStart + attr.index + attr[0].indexOf('"') + 1,
				length: attr[2].length,
			});
		}
		const name = attrs.get('name');
		if (name?.value) {
			occurrence.name = name.value;
			occurrence.nameSpan = { start: toSource(name.start), end: toSource(name.start + name.length) };
		}
		occurrence.hasHints = ['type', 'unit', 'value'].some((key) => !!attrs.get(key)?.value);
		if (m[3] === '/') {
			occurrence.text = '';
			occurrence.end = occurrence.open.end;
		} else {
			const close = lower.indexOf(`</${tag}>`, openEnd);
			const reopen = new RegExp(`<${tag}\\b`, 'g');
			reopen.lastIndex = openEnd;
			const next = reopen.exec(lower);
			if (close >= 0 && (!next || close < next.index)) {
				occurrence.text = collapse(body.slice(openEnd, close));
				occurrence.end = toSource(close + tag.length + 3);
			}
		}
		tags.push(occurrence);
	}
	return tags;
}
